/**
 * normalizer.js - Module chuẩn hóa đề thi tự động từ file Word thô
 * Pure Client-Side Implementation of normalizer.py
 */

(function(window) {
    const {
        NS,
        createWordElement,
        getChildByTagName,
        getChildrenByTagName,
        findDescendants,
        getParagraphText,
        getAllParagraphsInDoc
    } = window.XmlUtils;

    const QUESTION_RE = /^(?:Câu|Question|Bài|câu|bài)\s*(?:hỏi\s+)?(\d+)\s*(?:\([^)]*\))?\s*[:.\-)?\s]?/i;
    const PLAIN_NUMBER_RE = /^(\d+)\s*(?:\([^)]*\))?\s*[.:)]\s*(?!\d)/;
    const SECTION_RE = /^\s*(?:(?:phần|part)\s*([IVX]+|\d+)|([IVX]+|\d+)\s*[.:]\s*(?=.*?(?:trắc nghiệm|tự luận|đúng[\s/]*sai|trả lời ngắn|chọn đ\s*\()))/i;
    const INLINE_MC_RE = /(?<![a-zA-Z])(?:\()?([A-H])\s*[.:)/\-]\s*/g;
    const INLINE_TF_RE = /(?<![a-zA-Z])(?:\()?([a-h])\s*[.:)/\-]\s*/g;

    function splitParagraphsAtBr(xmlDoc) {
        const body = getChildByTagName(xmlDoc.documentElement, "body");
        if (!body) return;

        const paras = Array.from(body.getElementsByTagNameNS(NS.w, "p"));
        for (let i = 0; i < paras.length; i++) {
            const pElem = paras[i];
            const runs = getChildrenByTagName(pElem, "r");
            let hasBr = false;
            for (let r = 0; r < runs.length; r++) {
                if (getChildByTagName(runs[r], "br")) {
                    hasBr = true;
                    break;
                }
            }

            if (!hasBr) continue;
            const parent = pElem.parentNode;
            if (!parent) continue;

            const pPr = getChildByTagName(pElem, "pPr");
            const newParas = [];
            let currentP = createWordElement(xmlDoc, "p");
            if (pPr) currentP.appendChild(pPr.cloneNode(true));
            newParas.push(currentP);

            const children = Array.from(pElem.childNodes);
            for (let c = 0; c < children.length; c++) {
                const child = children[c];
                if (child.localName === "pPr") continue;
                if (child.localName === "r") {
                    const brs = getChildrenByTagName(child, "br");
                    if (brs.length === 0) {
                        currentP.appendChild(child.cloneNode(true));
                    } else {
                        const rPr = getChildByTagName(child, "rPr");
                        let currentR = createWordElement(xmlDoc, "r");
                        if (rPr) currentR.appendChild(rPr.cloneNode(true));
                        currentP.appendChild(currentR);

                        const rChildren = Array.from(child.childNodes);
                        for (let rc = 0; rc < rChildren.length; rc++) {
                            const rChild = rChildren[rc];
                            if (rChild.localName === "rPr") continue;
                            if (rChild.localName === "br") {
                                currentP = createWordElement(xmlDoc, "p");
                                if (pPr) currentP.appendChild(pPr.cloneNode(true));
                                newParas.push(currentP);

                                currentR = createWordElement(xmlDoc, "r");
                                if (rPr) currentR.appendChild(rPr.cloneNode(true));
                                currentP.appendChild(currentR);
                            } else {
                                currentR.appendChild(rChild.cloneNode(true));
                            }
                        }
                    }
                } else {
                    currentP.appendChild(child.cloneNode(true));
                }
            }

            for (let np = 0; np < newParas.length; np++) {
                parent.insertBefore(newParas[np], pElem);
            }
            parent.removeChild(pElem);
        }
    }

    function convertChoiceTablesToParagraphs(xmlDoc) {
        const CHOICE_LABEL_RE = /^\s*(?:\()?([A-Ha-h])\s*[.:)/\-]\s*/;
        const body = getChildByTagName(xmlDoc.documentElement, "body");
        if (!body) return;

        const tables = Array.from(body.getElementsByTagNameNS(NS.w, "tbl"));
        const tablesToProcess = [];

        for (let t = 0; t < tables.length; t++) {
            const table = tables[t];
            let choiceCellsCount = 0;
            let totalCells = 0;
            const cells = table.getElementsByTagNameNS(NS.w, "tc");
            for (let c = 0; c < cells.length; c++) {
                totalCells++;
                const text = cells[c].textContent.trim();
                if (CHOICE_LABEL_RE.test(text)) {
                    choiceCellsCount++;
                }
            }

            if (choiceCellsCount >= 2 && choiceCellsCount >= totalCells * 0.5) {
                tablesToProcess.push(table);
            }
        }

        for (let t = 0; t < tablesToProcess.length; t++) {
            const table = tablesToProcess[t];
            const parent = table.parentNode;
            if (!parent) continue;

            const cells = table.getElementsByTagNameNS(NS.w, "tc");
            for (let c = 0; c < cells.length; c++) {
                const paras = getChildrenByTagName(cells[c], "p");
                for (let p = 0; p < paras.length; p++) {
                    parent.insertBefore(paras[p], table);
                }
            }
            parent.removeChild(table);
        }
    }

    function cleanEmptyParagraphs(xmlDoc) {
        const body = getChildByTagName(xmlDoc.documentElement, "body");
        if (!body) return;

        const paras = Array.from(body.getElementsByTagNameNS(NS.w, "p"));
        for (let i = 0; i < paras.length; i++) {
            const p = paras[i];
            const text = getParagraphText(p).trim();
            if (text.length > 0) continue;
            if (findDescendants(p, "drawing").length > 0) continue;
            if (findDescendants(p, "pict").length > 0) continue;
            if (findDescendants(p, "oMath").length > 0) continue;
            if (findDescendants(p, "object").length > 0) continue;

            if (p.parentNode) {
                p.parentNode.removeChild(p);
            }
        }
    }

    function isShortAnswerValue(val) {
        if (!val) return false;
        const valClean = val.trim().toUpperCase();
        if (valClean.length === 1 && valClean >= 'A' && valClean <= 'H') return false;
        if (/^[ĐSDTRUEFSAI](?:[\s,]+[ĐSDTRUEFSAI])*$/.test(valClean)) return false;
        return true;
    }

    function parseTableAnswerKey(table) {
        const keys = {};
        const rows = getChildrenByTagName(table, "tr");
        if (rows.length === 0) return keys;

        const grid = [];
        for (let r = 0; r < rows.length; r++) {
            const cells = getChildrenByTagName(rows[r], "tc");
            grid.push(cells.map(c => c.textContent.trim()));
        }

        if (grid.length === 0 || grid[0].length === 0) return keys;
        const numRows = grid.length;
        const numCols = grid[0].length;

        // 1. Check for Question Row in horizontal table
        const questionRows = [];
        for (let rIdx = 0; rIdx < numRows; rIdx++) {
            const rowCells = grid[rIdx];
            if (!rowCells || rowCells.length === 0) continue;
            const isCol0Q = /^(?:Câu\s*)?(\d+)$/i.test(rowCells[0]);
            const startIdx = isCol0Q ? 0 : 1;

            let digitsCount = 0;
            let nonEmptyCols = 0;
            for (let c = startIdx; c < rowCells.length; c++) {
                if (rowCells[c]) {
                    nonEmptyCols++;
                    if (/^(?:Câu\s*)?(\d+)$/i.test(rowCells[c])) {
                        digitsCount++;
                    }
                }
            }
            if (digitsCount >= 2 && digitsCount >= nonEmptyCols * 0.7) {
                questionRows.push(rIdx);
            }
        }

        if (questionRows.length > 0) {
            for (let q = 0; q < questionRows.length; q++) {
                const qRowIdx = questionRows[q];
                const colMapping = {};
                const startCol = (grid[qRowIdx] && /^(?:Câu\s*)?(\d+)$/i.test(grid[qRowIdx][0])) ? 0 : 1;

                for (let cIdx = startCol; cIdx < numCols; cIdx++) {
                    const val = grid[qRowIdx][cIdx];
                    const m = /^(?:Câu\s*)?(\d+)$/i.exec(val);
                    if (m) {
                        colMapping[cIdx] = parseInt(m[1], 10);
                    }
                }

                const tfRows = [];
                for (let rIdx = qRowIdx + 1; rIdx < numRows; rIdx++) {
                    if (questionRows.includes(rIdx)) break;
                    let foundLabel = null;
                    for (let c = 0; c < grid[rIdx].length; c++) {
                        const cleaned = grid[rIdx][c].trim().toLowerCase();
                        if (/^[a-h]\s*[\).:]/.test(cleaned)) {
                            foundLabel = cleaned[0];
                            break;
                        }
                    }
                    if (foundLabel) {
                        tfRows.push({ label: foundLabel, rIdx });
                    }
                }

                if (tfRows.length >= 2) {
                    tfRows.sort((a, b) => a.label.localeCompare(b.label));
                    for (const [cIdxStr, qNum] of Object.entries(colMapping)) {
                        const cIdx = parseInt(cIdxStr, 10);
                        const ansList = [];
                        for (let t = 0; t < tfRows.length; t++) {
                            const { rIdx } = tfRows[t];
                            if (cIdx < grid[rIdx].length) {
                                const val = grid[rIdx][cIdx].trim().toUpperCase();
                                if (['Đ', 'ĐÚNG', 'T', 'TRUE', 'D'].includes(val)) {
                                    ansList.push('Đ');
                                } else if (['S', 'SAI', 'F', 'FALSE'].includes(val)) {
                                    ansList.push('S');
                                } else if (val) {
                                    ansList.push(val);
                                }
                            }
                        }
                        if (ansList.length > 0) {
                            keys[qNum] = ansList.join(",");
                        }
                    }
                } else {
                    for (let rIdx = qRowIdx + 1; rIdx < numRows; rIdx++) {
                        if (questionRows.includes(rIdx)) break;
                        const hasVal = grid[rIdx].slice(startCol).some(v => v);
                        if (!hasVal) continue;
                        for (const [cIdxStr, qNum] of Object.entries(colMapping)) {
                            const cIdx = parseInt(cIdxStr, 10);
                            if (cIdx < grid[rIdx].length) {
                                const ansVal = grid[rIdx][cIdx].trim();
                                if (ansVal) {
                                    keys[qNum] = ansVal.toUpperCase();
                                }
                            }
                        }
                        break;
                    }
                }
            }
            return keys;
        }

        // 2. Vertical True/False table
        for (let r = 0; r < rows.length; r++) {
            const cells = getChildrenByTagName(rows[r], "tc");
            if (cells.length >= 2) {
                const tFirst = cells[0].textContent.trim();
                if (/^\d+$/.test(tFirst)) {
                    const qNum = parseInt(tFirst, 10);
                    const tfVals = [];
                    let isTf = false;
                    for (let cIdx = 1; cIdx < cells.length; cIdx++) {
                        const cellTxt = cells[cIdx].textContent.trim();
                        const m = /^([a-h])\s*[\-.:\s]\s*([ĐSđsD])$/i.exec(cellTxt);
                        if (m) {
                            const ansVal = ['Đ', 'D'].includes(m[2].toUpperCase()) ? 'Đ' : 'S';
                            const optIdx = m[1].toLowerCase().charCodeAt(0) - 'a'.charCodeAt(0);
                            tfVals.push({ optIdx, ansVal });
                            isTf = true;
                        } else {
                            const val = cellTxt.toUpperCase();
                            if (['Đ', 'ĐÚNG', 'T', 'TRUE', 'D'].includes(val)) {
                                tfVals.push({ optIdx: cIdx - 1, ansVal: 'Đ' });
                            } else if (['S', 'SAI', 'F', 'FALSE'].includes(val)) {
                                tfVals.push({ optIdx: cIdx - 1, ansVal: 'S' });
                            }
                        }
                    }
                    if (isTf && tfVals.length > 0) {
                        tfVals.sort((a, b) => a.optIdx - b.optIdx);
                        keys[qNum] = tfVals.map(x => x.ansVal).join(",");
                        continue;
                    }
                }
            }
        }

        // 3. Vertical 2-column table
        for (let r = 0; r < rows.length; r++) {
            const cells = getChildrenByTagName(rows[r], "tc");
            for (let i = 0; i < cells.length - 1; i++) {
                const tLeft = cells[i].textContent.trim();
                const tRight = cells[i + 1].textContent.trim();
                if (/^\d+$/.test(tLeft)) {
                    const qNum = parseInt(tLeft, 10);
                    if (!(qNum in keys) && tRight && tRight.length < 15) {
                        keys[qNum] = tRight.toUpperCase();
                    }
                }
            }
        }

        // 4. Combined cell: "1. A" or "Câu 1: B"
        const patCell = /^(?:Câu\s*)?(\d+)\s*[-.:=/\s]+\s*(.+)$/i;
        for (let r = 0; r < rows.length; r++) {
            const cells = getChildrenByTagName(rows[r], "tc");
            for (let c = 0; c < cells.length; c++) {
                const txt = cells[c].textContent.trim();
                if (/^-?\d+[\.,]\d+$/.test(txt)) continue;
                const m = patCell.exec(txt);
                if (m) {
                    const qNum = parseInt(m[1], 10);
                    if (!(qNum in keys)) {
                        keys[qNum] = m[2].trim();
                    }
                }
            }
        }

        return keys;
    }

    function parseTextAnswerKey(text) {
        const keys = {};

        // 1. Part 2 (True/False): 23(Đ,S,Đ,S) or 23(Đ-S-Đ-S)
        const patternTf = /(\d+)\s*[\(\[:]?\s*([ĐSđsD](?:[\s,.\-/]+[ĐSđsD]){1,15})\s*[\)\]]?/gi;
        let mTf;
        while ((mTf = patternTf.exec(text)) !== null) {
            const qNum = parseInt(mTf[1], 10);
            const rawVals = (mTf[2].toUpperCase().match(/[ĐSđsD]/g) || []).map(v => v === 'D' ? 'Đ' : v);
            if (rawVals.length > 0) {
                keys[qNum] = rawVals.join(",");
            }
        }

        // 2. Part 3 (Short answer): 26: 15.5 or 26 = 15.5
        const patternShort = /(\d+)\s*[:=]\s*(-?\d+(?:[\.,]\d+)?)/g;
        let mShort;
        while ((mShort = patternShort.exec(text)) !== null) {
            const qNum = parseInt(mShort[1], 10);
            if (!(qNum in keys)) {
                keys[qNum] = mShort[2];
            }
        }

        // 3. Part 1 (Multiple choice): 1A, 2.B, 3-C, Câu 4: D
        const patternMc = /(?:Câu\s*)?(\d+)\s*[.\-:]?\s*([A-D])(?![a-zA-Z0-9])/gi;
        let mMc;
        while ((mMc = patternMc.exec(text)) !== null) {
            const qNum = parseInt(mMc[1], 10);
            if (!(qNum in keys)) {
                keys[qNum] = mMc[2].toUpperCase();
            }
        }

        return keys;
    }

    function isAnswerTitle(text) {
        const textClean = text.trim();
        if (!textClean) return false;

        const isTitle = (
            /^\s*(?:BẢNG\s+)?ĐÁP\s+ÁN(?:\s+CHI\s+TIẾT)?\s*$/i.test(textClean) ||
            /^\s*HƯỚNG\s+DẪN\s+CHẤM(?:\s+VÀ\s+BIỂU\s+ĐIỂM)?\s*$/i.test(textClean) ||
            /^\s*PHẦN\s+ĐÁP\s+ÁN\s*$/i.test(textClean) ||
            /^\s*ĐÁP\s+ÁN\s+CHI\s+TIẾT\s*$/i.test(textClean) ||
            /^\s*(?:BẢNG\s+)?ĐÁP\s+ÁN\s+CÁC\s+MÃ\s+ĐỀ\s*$/i.test(textClean)
        );
        if (isTitle) return true;

        if (textClean.length < 50) {
            const lower = textClean.toLowerCase();
            const prefixes = ["đáp án", "hướng dẫn chấm", "bảng đáp án", "biểu điểm", "hướng dẫn chấm và biểu điểm", "đáp án và hướng dẫn chấm"];
            if (prefixes.some(p => lower.startsWith(p))) return true;
        }

        if (/[\-\s_*=]{3,}(?:đáp án|hướng dẫn chấm|biểu điểm)[\-\s_*=]{3,}/i.test(textClean)) return true;
        return false;
    }

    function getAnswerForQuestion(answerKeys, sectionIdx, qNum, offsets = { 1: 0, 2: 0, 3: 0 }) {
        if (answerKeys[sectionIdx] && answerKeys[sectionIdx][qNum]) {
            return answerKeys[sectionIdx][qNum];
        }
        const absQNum = qNum + (offsets[sectionIdx] || 0);
        for (const secIdx of [sectionIdx, 1, 2, 3]) {
            if (answerKeys[secIdx] && answerKeys[secIdx][absQNum]) {
                return answerKeys[secIdx][absQNum];
            }
        }
        return null;
    }

    function extractAndRemoveAnswerKey(xmlDoc) {
        const keys = { 1: {}, 2: {}, 3: {} };
        const body = getChildByTagName(xmlDoc.documentElement, "body");
        if (!body) return keys;

        const elements = [];
        for (let i = 0; i < body.childNodes.length; i++) {
            const child = body.childNodes[i];
            if (child.nodeType === 1) {
                if (child.localName === "p") {
                    elements.push({ type: "para", el: child });
                } else if (child.localName === "tbl") {
                    elements.push({ type: "table", el: child });
                }
            }
        }

        let titleIdx = -1;
        for (let i = elements.length - 1; i >= 0; i--) {
            const { type, el } = elements[i];
            if (type === "para") {
                const text = el.textContent.trim();
                if (isAnswerTitle(text)) {
                    titleIdx = i;
                    break;
                }
            } else if (type === "table") {
                let found = false;
                const paras = el.getElementsByTagNameNS(NS.w, "p");
                for (let p = 0; p < paras.length; p++) {
                    const cellTxt = paras[p].textContent.trim();
                    if (!["đáp án", "câu", "ý"].includes(cellTxt.toLowerCase()) && isAnswerTitle(cellTxt)) {
                        titleIdx = i;
                        found = true;
                        break;
                    }
                }
                if (found) break;
            }
        }

        const elementsToRemove = [];

        if (titleIdx !== -1) {
            let currentAnsSectionIdx = 1;
            for (let idx = titleIdx; idx < elements.length; idx++) {
                const { type, el } = elements[idx];
                elementsToRemove.push(el);

                if (type === "para") {
                    const text = el.textContent.trim();
                    const secMatch = SECTION_RE.exec(text);
                    if (secMatch) {
                        const romanToInt = { 'I': 1, 'II': 2, 'III': 3, 'IV': 4, '1': 1, '2': 2, '3': 3, '4': 4 };
                        const partStr = (secMatch[1] || secMatch[2]).toUpperCase();
                        if (romanToInt[partStr]) {
                            currentAnsSectionIdx = romanToInt[partStr];
                        }
                    }

                    if (text.toLowerCase().startsWith("giải thích") || text.toLowerCase().startsWith("lời giải") || text.toLowerCase().startsWith("hd") || text.toLowerCase().startsWith("hướng dẫn")) {
                        continue;
                    }

                    const pKeys = parseTextAnswerKey(text);
                    if (Object.keys(pKeys).length > 0) {
                        if (!keys[currentAnsSectionIdx]) keys[currentAnsSectionIdx] = {};
                        Object.assign(keys[currentAnsSectionIdx], pKeys);
                    }
                } else if (type === "table") {
                    const tKeys = parseTableAnswerKey(el);
                    if (Object.keys(tKeys).length > 0) {
                        if (!keys[currentAnsSectionIdx]) keys[currentAnsSectionIdx] = {};
                        Object.assign(keys[currentAnsSectionIdx], tKeys);
                    }
                }
            }
        } else {
            let answerStarted = false;
            let currentAnsSectionIdx = 1;
            for (let i = elements.length - 1; i >= 0; i--) {
                const { type, el } = elements[i];
                if (type === "table") {
                    const tKeys = parseTableAnswerKey(el);
                    if (Object.keys(tKeys).length > 0) {
                        if (!keys[currentAnsSectionIdx]) keys[currentAnsSectionIdx] = {};
                        Object.assign(keys[currentAnsSectionIdx], tKeys);
                        elementsToRemove.push(el);
                        answerStarted = true;
                    } else if (answerStarted) {
                        break;
                    }
                } else if (type === "para") {
                    const text = el.textContent.trim();
                    const secMatch = SECTION_RE.exec(text);
                    if (secMatch) {
                        const romanToInt = { 'I': 1, 'II': 2, 'III': 3, 'IV': 4, '1': 1, '2': 2, '3': 3, '4': 4 };
                        const partStr = (secMatch[1] || secMatch[2]).toUpperCase();
                        if (romanToInt[partStr]) {
                            currentAnsSectionIdx = romanToInt[partStr];
                        }
                    }

                    if (text.toLowerCase().startsWith("giải thích") || text.toLowerCase().startsWith("lời giải") || text.toLowerCase().startsWith("hd") || text.toLowerCase().startsWith("hướng dẫn")) {
                        if (answerStarted) elementsToRemove.push(el);
                        continue;
                    }

                    const pKeys = parseTextAnswerKey(text);
                    if (Object.keys(pKeys).length > 0) {
                        if (!keys[currentAnsSectionIdx]) keys[currentAnsSectionIdx] = {};
                        Object.assign(keys[currentAnsSectionIdx], pKeys);
                        elementsToRemove.push(el);
                        answerStarted = true;
                    } else if (answerStarted) {
                        const isDivider = !text || /^[\-\s_.*=]+$/.test(text);
                        const isPartHeader = text.length < 100 && (/^(phần|part|đáp án|hướng dẫn|câu)/i.test(text) || !/\w/.test(text));
                        const isQuestion = /^(câu|question)\s*\d+/i.test(text);
                        const isExplanation = text.toLowerCase().startsWith("giải thích") || text.toLowerCase().startsWith("lời giải") || text.toLowerCase().startsWith("hd") || text.toLowerCase().startsWith("hướng dẫn");

                        if (isDivider || isExplanation || (isPartHeader && !isQuestion)) {
                            elementsToRemove.push(el);
                        } else {
                            break;
                        }
                    }
                }
            }
        }

        for (let i = 0; i < elementsToRemove.length; i++) {
            const el = elementsToRemove[i];
            if (el.parentNode) {
                el.parentNode.removeChild(el);
            }
        }

        return keys;
    }

    function preprocessSplitRuns(pNode, splitLabels, xmlDoc, maxDepth = 50) {
        if (maxDepth <= 0) return;
        const suffixPattern = '[.:)/\\-]';
        const runs = getChildrenByTagName(pNode, "r");

        for (let r = 0; r < runs.length; r++) {
            const rElem = runs[r];
            const tEl = getChildByTagName(rElem, "t");
            if (tEl && tEl.textContent) {
                const text = tEl.textContent;
                for (let l = 0; l < splitLabels.length; l++) {
                    const label = splitLabels[l];
                    const pat = new RegExp(`(?<![a-zA-Z])(${label}\\s*${suffixPattern})`);
                    const m = pat.exec(text);
                    if (m && m.index > 0) {
                        const t1 = text.slice(0, m.index);
                        const t2 = text.slice(m.index);
                        tEl.textContent = t1;

                        const newR = rElem.cloneNode(true);
                        const newT = getChildByTagName(newR, "t");
                        if (newT) newT.textContent = t2;

                        if (rElem.nextSibling) {
                            pNode.insertBefore(newR, rElem.nextSibling);
                        } else {
                            pNode.appendChild(newR);
                        }

                        preprocessSplitRuns(pNode, splitLabels, xmlDoc, maxDepth - 1);
                        return;
                    }
                }
            }
        }
    }

    function mergeAdjacentRuns(pNode) {
        const children = Array.from(pNode.childNodes).filter(n => n.nodeType === 1);
        let i = 0;
        let modified = false;
        let guard = 500;
        while (i < children.length - 1 && guard-- > 0) {
            const child1 = children[i];
            const child2 = children[i + 1];
            if (child1.localName === "r" && child2.localName === "r") {
                const t1 = getChildByTagName(child1, "t");
                const t2 = getChildByTagName(child2, "t");

                if (t1 && t2) {
                    const t1Text = t1.textContent || "";
                    const t2Text = t2.textContent || "";

                    const special1 = Array.from(child1.childNodes).some(c => !["rPr", "t", "tab", "br"].includes(c.localName));
                    const special2 = Array.from(child2.childNodes).some(c => !["rPr", "t", "tab", "br"].includes(c.localName));

                    if (special1 || special2) {
                        i++;
                        continue;
                    }

                    const rPr1 = getChildByTagName(child1, "rPr");
                    const rPr2 = getChildByTagName(child2, "rPr");
                    const rPr1Str = rPr1 ? (rPr1.outerHTML || window.XmlUtils.serializeXml(rPr1)) : "";
                    const rPr2Str = rPr2 ? (rPr2.outerHTML || window.XmlUtils.serializeXml(rPr2)) : "";

                    const sameRpr = (rPr1Str === rPr2Str);
                    const isCombinable = sameRpr || (t2Text.trim() === ".");

                    if (isCombinable) {
                        t1.textContent = t1Text + t2Text;
                        pNode.removeChild(child2);
                        children.splice(i + 1, 1);
                        modified = true;
                        continue;
                    }
                }
            }
            i++;
        }
        return modified;
    }

    function normalizeParagraphPrefix(pNode, currentQNum) {
        const text = getParagraphText(pNode);
        let qMatch = QUESTION_RE.exec(text) || PLAIN_NUMBER_RE.exec(text);
        if (!qMatch) return;

        const prefixLen = qMatch[0].length;
        const remainder = text.slice(prefixLen);
        const remainderClean = remainder.replace(/^\s+/, '');
        const totalReplaceLen = text.length - remainderClean.length;

        let accumulatedLen = 0;
        const runsToModify = [];
        const runs = getChildrenByTagName(pNode, "r");

        for (let r = 0; r < runs.length; r++) {
            const run = runs[r];
            const tEl = getChildByTagName(run, "t");
            const runText = tEl ? tEl.textContent : "";
            runsToModify.push({ run, tEl, runText });
            accumulatedLen += runText.length;
            if (accumulatedLen >= totalReplaceLen) break;
        }

        if (runsToModify.length > 0) {
            const { tEl: firstT, runText: firstText } = runsToModify[0];
            if (firstText.length >= totalReplaceLen) {
                if (firstT) firstT.textContent = `Câu ${currentQNum}. ` + firstText.slice(totalReplaceLen);
            } else {
                if (firstT) firstT.textContent = `Câu ${currentQNum}. `;
                let remainingToClear = totalReplaceLen - firstText.length;
                for (let j = 1; j < runsToModify.length; j++) {
                    const { tEl, runText } = runsToModify[j];
                    if (tEl) {
                        if (runText.length <= remainingToClear) {
                            tEl.textContent = "";
                            remainingToClear -= runText.length;
                        } else {
                            tEl.textContent = runText.slice(remainingToClear);
                            break;
                        }
                    }
                }
            }
        }
    }

    function normalizeSectionPrefix(pNode, currentSectionIdx) {
        const text = getParagraphText(pNode);
        const secMatch = SECTION_RE.exec(text);
        if (!secMatch) return;

        const prefixLen = secMatch[0].length;
        const remainder = text.slice(prefixLen);
        const remainderClean = remainder.replace(/^[:.\-\s]+/, '');
        const totalReplaceLen = text.length - remainderClean.length;

        let accumulatedLen = 0;
        const runsToModify = [];
        const runs = getChildrenByTagName(pNode, "r");

        for (let r = 0; r < runs.length; r++) {
            const run = runs[r];
            const tEl = getChildByTagName(run, "t");
            const runText = tEl ? tEl.textContent : "";
            runsToModify.push({ run, tEl, runText });
            accumulatedLen += runText.length;
            if (accumulatedLen >= totalReplaceLen) break;
        }

        if (runsToModify.length > 0) {
            const { tEl: firstT, runText: firstText } = runsToModify[0];
            if (firstText.length >= totalReplaceLen) {
                if (firstT) firstT.textContent = `PHẦN ${currentSectionIdx}. ` + firstText.slice(totalReplaceLen);
            } else {
                if (firstT) firstT.textContent = `PHẦN ${currentSectionIdx}. `;
                let remainingToClear = totalReplaceLen - firstText.length;
                for (let j = 1; j < runsToModify.length; j++) {
                    const { tEl, runText } = runsToModify[j];
                    if (tEl) {
                        if (runText.length <= remainingToClear) {
                            tEl.textContent = "";
                            remainingToClear -= runText.length;
                        } else {
                            tEl.textContent = runText.slice(remainingToClear);
                            break;
                        }
                    }
                }
            }
        }
    }

    function splitInlineAnswers(xmlDoc, pNode, text, isLower = false) {
        const pattern = isLower ? /(?<![a-zA-Z])(?:\()?([a-h])\s*[.:)/\-]\s*/g : /(?<![a-zA-Z])(?:\()?([A-H])\s*[.:)/\-]\s*/g;
        const matches = [];
        let m;
        while ((m = pattern.exec(text)) !== null) {
            matches.push({ label: m[1], index: m.index });
        }
        if (matches.length === 0) return false;

        const firstMatch = matches[0];
        const prefixText = text.slice(0, firstMatch.index).trim();
        const splitLabels = (prefixText.length > 0) ? matches.map(x => x.label) : matches.slice(1).map(x => x.label);
        if (splitLabels.length === 0) return false;

        preprocessSplitRuns(pNode, splitLabels, xmlDoc);

        const parent = pNode.parentNode;
        if (!parent) return false;

        const pPr = getChildByTagName(pNode, "pPr");
        const splitBuckets = [[]];
        let currentBucket = splitBuckets[0];
        let labelIdx = 0;

        const children = Array.from(pNode.childNodes);
        for (let c = 0; c < children.length; c++) {
            const child = children[c];
            if (child.localName === "pPr") continue;

            let startsNewOption = false;
            if (labelIdx < splitLabels.length) {
                const nextLabel = splitLabels[labelIdx];
                const childText = child.textContent.trim();
                const matchPattern = new RegExp(`^\\s*(?:\\()?${nextLabel}\\s*[.:)/\\-]`, 'i');
                if (matchPattern.test(childText)) {
                    startsNewOption = true;
                }
            }

            if (startsNewOption) {
                currentBucket = [];
                splitBuckets.push(currentBucket);
                labelIdx++;
            }

            currentBucket.push(child);
        }

        if (splitBuckets.length <= 1) return false;

        function createSplitPara() {
            const newP = createWordElement(xmlDoc, "p");
            if (pPr) newP.appendChild(pPr.cloneNode(true));
            return newP;
        }

        const newParagraphs = [];
        for (let b = 0; b < splitBuckets.length; b++) {
            const newP = createSplitPara();
            for (let k = 0; k < splitBuckets[b].length; k++) {
                newP.appendChild(splitBuckets[b][k]);
            }
            newParagraphs.push(newP);
        }

        for (let p = 0; p < newParagraphs.length; p++) {
            const np = newParagraphs[p];
            const runs = getChildrenByTagName(np, "r");
            for (let r = 0; r < runs.length; r++) {
                const tEl = getChildByTagName(runs[r], "t");
                if (tEl && tEl.textContent && tEl.textContent.trim()) {
                    tEl.textContent = tEl.textContent.replace(/^\s+/, '');
                    break;
                }
            }
            parent.insertBefore(np, pNode);
        }

        parent.removeChild(pNode);
        return true;
    }

    function addAsteriskToLabel(pNode, label, isLower = false) {
        const wtElems = findDescendants(pNode, "t");
        if (wtElems.length === 0) return false;

        const suffixPattern = '[.:)/\\-]';
        const pat = new RegExp(`^(\\s*)(?:\\()?${label}(\\s*${suffixPattern})`, 'i');

        for (let i = 0; i < wtElems.length; i++) {
            const wt = wtElems[i];
            const t = wt.textContent || "";
            if (pat.test(t)) {
                wt.textContent = t.replace(pat, `$1${label}*$2`);
                return true;
            }
        }

        for (let i = 0; i < wtElems.length; i++) {
            const wt = wtElems[i];
            const t = wt.textContent || "";
            if (t.trim() === label || t.trim() === `${label}.`) {
                wt.textContent = t.replace(label, `${label}*`);
                return true;
            }
        }

        return false;
    }

    function isParaMarkedAsCorrect(pNode) {
        const runs = getChildrenByTagName(pNode, "r");
        for (let r = 0; r < runs.length; r++) {
            const rPr = getChildByTagName(runs[r], "rPr");
            if (!rPr) continue;

            // 1. In đậm
            if (getChildByTagName(rPr, "b")) return true;

            // 2. Gạch chân
            if (getChildByTagName(rPr, "u")) return true;

            // 3. Chữ màu đỏ
            const colorEl = getChildByTagName(rPr, "color");
            if (colorEl) {
                const val = colorEl.getAttributeNS(NS.w, "val") || colorEl.getAttribute("w:val") || "";
                if (val.toUpperCase() === "FF0000" || val.toUpperCase() === "RED") return true;
            }

            // 4. Bôi nền
            if (getChildByTagName(rPr, "highlight") || getChildByTagName(rPr, "shd")) return true;
        }
        return false;
    }

    async function normalizeDocx(zip, options = {}) {
        const autoKey = options.auto_key !== false;
        const boldAsCorrect = options.bold_as_correct === true;
        const choiceLayout = options.choice_layout || "split";

        const docXmlStr = await zip.file("word/document.xml").async("string");
        const xmlDoc = window.XmlUtils.parseXml(docXmlStr);

        splitParagraphsAtBr(xmlDoc);
        convertChoiceTablesToParagraphs(xmlDoc);
        cleanEmptyParagraphs(xmlDoc);

        let answerKeys = {};
        let keyCount = 0;
        if (autoKey) {
            answerKeys = extractAndRemoveAnswerKey(xmlDoc);
            keyCount = Object.values(answerKeys).reduce((acc, sec) => acc + Object.keys(sec).length, 0);
        }

        const body = getChildByTagName(xmlDoc.documentElement, "body");
        let paragraphs = getAllParagraphsInDoc(body);

        // Calculate offsets
        const secQuestions = { 1: [], 2: [], 3: [] };
        let tempSecIdx = 0;
        for (let i = 0; i < paragraphs.length; i++) {
            const pText = getParagraphText(paragraphs[i]).trim();
            if (!pText) continue;
            if (SECTION_RE.test(pText)) {
                tempSecIdx++;
                continue;
            }
            if (tempSecIdx >= 1 && tempSecIdx <= 3) {
                const qm = QUESTION_RE.exec(pText) || PLAIN_NUMBER_RE.exec(pText);
                if (qm) {
                    secQuestions[tempSecIdx].push(parseInt(qm[1], 10));
                }
            }
        }

        let isAlreadyContinuous = true;
        if (secQuestions[1].length > 0 && secQuestions[2].length > 0) {
            if (Math.min(...secQuestions[2]) <= Math.max(...secQuestions[1])) {
                isAlreadyContinuous = false;
            }
        }
        if (isAlreadyContinuous && secQuestions[2].length > 0 && secQuestions[3].length > 0) {
            if (Math.min(...secQuestions[3]) <= Math.max(...secQuestions[2])) {
                isAlreadyContinuous = false;
            }
        }

        const offsets = { 1: 0, 2: 0, 3: 0 };
        if (!isAlreadyContinuous) {
            offsets[2] = secQuestions[1].length;
            offsets[3] = secQuestions[1].length + secQuestions[2].length;
        }

        let currentSectionIdx = 0;
        let currentQNum = 0;
        let pIdx = 0;
        const splitDoneIds = new Set();
        let loopGuard = paragraphs.length * 10 + 5000;

        while (pIdx < paragraphs.length && loopGuard-- > 0) {
            const p = paragraphs[pIdx];
            mergeAdjacentRuns(p);
            const text = getParagraphText(p).trim();

            if (!text) {
                pIdx++;
                continue;
            }

            // A. Check Section
            const secMatch = SECTION_RE.exec(text);
            if (secMatch) {
                currentSectionIdx++;
                normalizeSectionPrefix(p, currentSectionIdx);
                const firstR = getChildByTagName(p, "r");
                if (firstR) {
                    const rPr = window.XmlUtils.getOrCreateRPr(firstR, xmlDoc);
                    if (!getChildByTagName(rPr, "b")) {
                        rPr.appendChild(createWordElement(xmlDoc, "b"));
                    }
                }
                pIdx++;
                continue;
            }

            // B. Check Question
            const qMatch = QUESTION_RE.exec(text) || PLAIN_NUMBER_RE.exec(text);
            if (qMatch) {
                currentQNum = parseInt(qMatch[1], 10);
                normalizeParagraphPrefix(p, currentQNum);
            }

            // E. Inline answers split
            if (currentQNum > 0 && !splitDoneIds.has(p)) {
                splitDoneIds.add(p);
                if (splitInlineAnswers(xmlDoc, p, text, currentSectionIdx === 2)) {
                    paragraphs = getAllParagraphsInDoc(body);
                    continue;
                }
            }

            // C. Multiple Choice (A. B. C. D.)
            const mUpper = /^(?:\()?([A-H])\s*[.:)/\-]\s*/.exec(text);
            if (mUpper && currentQNum > 0) {
                const label = mUpper[1].toUpperCase();
                let isCorrect = false;
                const correctVal = getAnswerForQuestion(answerKeys, currentSectionIdx, currentQNum, offsets);
                if (correctVal && label === correctVal.toUpperCase()) {
                    isCorrect = true;
                }
                if (!correctVal && boldAsCorrect && isParaMarkedAsCorrect(p)) {
                    isCorrect = true;
                }
                if (isCorrect) {
                    addAsteriskToLabel(p, label, false);
                }
                pIdx++;
                continue;
            }

            // D. True/False Part 2 (a) b) c) d))
            const mLower = /^([a-h])\s*[\).]\s*/.exec(text);
            if (mLower && currentQNum > 0) {
                const label = mLower[1].toLowerCase();
                let isCorrect = false;
                const correctVal = getAnswerForQuestion(answerKeys, currentSectionIdx, currentQNum, offsets);
                if (correctVal) {
                    const vals = correctVal.split(',').map(x => x.trim());
                    const idxVal = label.charCodeAt(0) - 'a'.charCodeAt(0);
                    if (idxVal < vals.length && ['Đ', 'D'].includes(vals[idxVal])) {
                        isCorrect = true;
                    }
                }
                if (!correctVal && boldAsCorrect && isParaMarkedAsCorrect(p)) {
                    isCorrect = true;
                }
                if (isCorrect) {
                    addAsteriskToLabel(p, label, true);
                }
                pIdx++;
                continue;
            }

            pIdx++;
        }

        // 3. Short answer insert for Part 3
        if (currentSectionIdx >= 3 || keyCount > 0) {
            paragraphs = getAllParagraphsInDoc(body);
            pIdx = 0;
            let part3Guard = paragraphs.length * 5 + 1000;
            while (pIdx < paragraphs.length && part3Guard-- > 0) {
                const p = paragraphs[pIdx];
                const text = getParagraphText(p).trim();
                const qMatch = QUESTION_RE.exec(text);
                if (qMatch) {
                    const qNum = parseInt(qMatch[1], 10);
                    let secNum = 0;
                    for (let j = pIdx; j >= 0; j--) {
                        const sm = SECTION_RE.exec(getParagraphText(paragraphs[j]).trim());
                        if (sm) {
                            secNum = parseInt(sm[1] || sm[2], 10);
                            break;
                        }
                    }

                    const correctVal = getAnswerForQuestion(answerKeys, secNum, qNum, offsets);
                    if (correctVal && (secNum === 3 || isShortAnswerValue(correctVal))) {
                        let hasAnswerLine = false;
                        if (pIdx + 1 < paragraphs.length) {
                            const nextText = getParagraphText(paragraphs[pIdx + 1]).trim();
                            if (nextText.startsWith("Đáp án")) {
                                hasAnswerLine = true;
                            }
                        }

                        if (!hasAnswerLine && p.parentNode) {
                            const newP = createWordElement(xmlDoc, "p");
                            const newR = createWordElement(xmlDoc, "r");
                            const newT = createWordElement(xmlDoc, "t");
                            newT.textContent = `Đáp án: ${correctVal}`;
                            newR.appendChild(newT);
                            newP.appendChild(newR);

                            if (p.nextSibling) {
                                p.parentNode.insertBefore(newP, p.nextSibling);
                            } else {
                                p.parentNode.appendChild(newP);
                            }
                            paragraphs = getAllParagraphsInDoc(body);
                            pIdx++;
                        }
                    }
                }
                pIdx++;
            }
        }

        // 4. Choice Compacting
        let compactedMsg = "";
        if (choiceLayout && choiceLayout !== "split") {
            const { compactedCount } = window.ChoiceCompactor.compactDocumentChoices(xmlDoc, choiceLayout);
            if (compactedCount > 0) {
                compactedMsg = ` Đã dồn dòng ${compactedCount} câu hỏi.`;
            }
        }

        const newDocXmlStr = window.XmlUtils.serializeXml(xmlDoc);
        zip.file("word/document.xml", newDocXmlStr);

        return {
            success: true,
            keyCount,
            answerKeys,
            message: `Chuẩn hóa thành công! Đã quét được ${keyCount} đáp án.${compactedMsg}`
        };
    }

    window.Normalizer = {
        splitParagraphsAtBr,
        convertChoiceTablesToParagraphs,
        extractAndRemoveAnswerKey,
        isShortAnswerValue,
        getAnswerForQuestion,
        normalizeDocx
    };
})(window);
