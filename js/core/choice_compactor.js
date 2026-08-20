/**
 * choice_compactor.js - Module dồn dòng / gộp các phương án trắc nghiệm trong file Word (.docx)
 * Pure Client-Side Implementation of choice_compactor.py
 * 
 * Hỗ trợ các chế độ:
 * - 'auto': Tự động phân tích độ dài từng phương án để xếp 4, 2 hoặc 1 phương án / dòng (Auto 4-2-1).
 * - '4_per_line': Gộp 4 phương án trên 1 dòng.
 * - '2_per_line': Gộp 2 phương án trên 1 dòng (2 dòng: A-B và C-D).
 * - '1_per_line' / 'split': Giữ nguyên mỗi phương án 1 dòng riêng.
 * 
 * Bảo toàn 100% công thức toán học MathType / OMML, hình ảnh, in đậm/nghiêng, màu sắc và dấu hoa thị (*).
 */

(function(window) {
    const {
        NS,
        createWordElement,
        getChildByTagName,
        getChildrenByTagName,
        findDescendants,
        getParagraphText
    } = window.XmlUtils;

    const CHOICE_UPPER_RE = /^\s*(?:\()?(\*?\s*[A-H]\s*\*?)(?:\s*[.:)/\-]\s*|\s*[\).]\s*\*?)/i;
    const CHOICE_LOWER_RE = /^\s*(?:\()?(\*?\s*[a-h]\s*\*?)(?:\s*[\).]\s*\*?|\s*[.:)/\-]\s*)/;

    const TAB_STOPS_4_COLS = [2100, 4200, 6300]; // ~3.70cm, ~7.40cm, ~11.10cm
    const TAB_STOPS_2_COLS = [4500];             // ~7.94cm
    const TAB_STOPS_3_COLS = [2900, 5800];       // ~5.11cm, ~10.23cm

    function extractChoiceLabelInfo(pNode) {
        const text = getParagraphText(pNode).trim();
        if (!text) return null;

        const mUp = CHOICE_UPPER_RE.exec(text);
        if (mUp) {
            const raw = mUp[1].replace(/\*/g, '').trim().toUpperCase();
            if (raw.length === 1 && raw >= 'A' && raw <= 'H') {
                return {
                    label: raw,
                    isLower: false,
                    index: raw.charCodeAt(0) - 'A'.charCodeAt(0)
                };
            }
        }

        const mLow = CHOICE_LOWER_RE.exec(text);
        if (mLow) {
            const raw = mLow[1].replace(/\*/g, '').trim().toLowerCase();
            if (raw.length === 1 && raw >= 'a' && raw <= 'h') {
                return {
                    label: raw,
                    isLower: true,
                    index: raw.charCodeAt(0) - 'a'.charCodeAt(0)
                };
            }
        }

        return null;
    }

    function hasHeavyContent(pNode) {
        const drawings = findDescendants(pNode, "drawing");
        if (drawings.length > 0) return true;

        const picts = findDescendants(pNode, "pict");
        for (let i = 0; i < picts.length; i++) {
            const ole = findDescendants(picts[i], "OLEObject");
            if (ole.length === 0) return true;
        }
        return false;
    }

    function getParagraphEstimatedLength(pNode) {
        const text = getParagraphText(pNode).trim();
        let textLen = text.length;

        let oleExtraLen = 0;
        const objects = findDescendants(pNode, "object");
        for (let i = 0; i < objects.length; i++) {
            const obj = objects[i];
            const dxaOrig = obj.getAttributeNS(NS.w, "dxaOrig") || obj.getAttribute("w:dxaOrig");
            if (dxaOrig && /^\d+$/.test(dxaOrig)) {
                oleExtraLen += Math.max(5, Math.floor(parseInt(dxaOrig, 10) / 95));
            } else {
                const shapes = obj.getElementsByTagName("*");
                let shape = null;
                for (let j = 0; j < shapes.length; j++) {
                    if (shapes[j].localName === "shape") {
                        shape = shapes[j];
                        break;
                    }
                }
                if (shape) {
                    const styleStr = shape.getAttribute("style") || "";
                    const mW = /width\s*:\s*([\d.]+)\s*pt/i.exec(styleStr);
                    if (mW) {
                        const ptVal = parseFloat(mW[1]);
                        oleExtraLen += Math.max(5, Math.floor((ptVal * 20) / 95));
                    } else {
                        oleExtraLen += 20;
                    }
                } else {
                    oleExtraLen += 20;
                }
            }
        }

        let ommlExtraLen = 0;
        const oMaths = findDescendants(pNode, "oMath");
        for (let i = 0; i < oMaths.length; i++) {
            const mTxt = oMaths[i].textContent || "";
            ommlExtraLen += Math.max(8, mTxt.length + 4);
        }

        const drawings = findDescendants(pNode, "drawing").length;
        const imgExtraLen = drawings * 40;

        return textLen + oleExtraLen + ommlExtraLen + imgExtraLen;
    }

    function cleanTrailingWhitespaceAndTabs(pNode) {
        let outerGuard = 50;
        while (outerGuard-- > 0) {
            const children = Array.from(pNode.childNodes).filter(n => n.nodeType === 1);
            if (children.length === 0) break;
            const lastC = children[children.length - 1];
            if (lastC.localName === "pPr") break;
            if (lastC.localName === "r") {
                let innerGuard = 50;
                while (innerGuard-- > 0) {
                    const rChildren = Array.from(lastC.childNodes).filter(n => n.nodeType === 1);
                    if (rChildren.length === 0) break;
                    const lastRC = rChildren[rChildren.length - 1];
                    if (["tab", "br", "cr", "noBreakHyphen"].includes(lastRC.localName)) {
                        lastC.removeChild(lastRC);
                    } else if (lastRC.localName === "t") {
                        if (lastRC.textContent) {
                            lastRC.textContent = lastRC.textContent.replace(/[\s\t\r\n\xa0]+$/, "");
                        }
                        if (!lastRC.textContent) {
                            lastC.removeChild(lastRC);
                        } else {
                            break;
                        }
                    } else if (lastRC.localName === "rPr") {
                        break;
                    } else {
                        break;
                    }
                }
                const meaningful = Array.from(lastC.childNodes).filter(n => n.nodeType === 1 && n.localName !== "rPr");
                if (meaningful.length === 0) {
                    pNode.removeChild(lastC);
                    continue;
                } else {
                    break;
                }
            } else {
                break;
            }
        }
    }

    function cleanLeadingWhitespaceAndTabs(pNode) {
        let outerGuard = 50;
        while (outerGuard-- > 0) {
            const children = Array.from(pNode.childNodes).filter(n => n.nodeType === 1);
            if (children.length === 0) break;
            let firstC = null;
            for (let i = 0; i < children.length; i++) {
                if (children[i].localName !== "pPr") {
                    firstC = children[i];
                    break;
                }
            }
            if (!firstC) break;
            if (firstC.localName === "r") {
                let innerGuard = 50;
                while (innerGuard-- > 0) {
                    const rChildren = Array.from(firstC.childNodes).filter(n => n.nodeType === 1);
                    if (rChildren.length === 0) break;
                    let firstRC = null;
                    for (let i = 0; i < rChildren.length; i++) {
                        if (rChildren[i].localName !== "rPr") {
                            firstRC = rChildren[i];
                            break;
                        }
                    }
                    if (!firstRC) break;
                    if (["tab", "br", "cr", "noBreakHyphen"].includes(firstRC.localName)) {
                        firstC.removeChild(firstRC);
                    } else if (firstRC.localName === "t") {
                        if (firstRC.textContent) {
                            firstRC.textContent = firstRC.textContent.replace(/^[\s\t\r\n\xa0]+/, "");
                        }
                        if (!firstRC.textContent) {
                            firstC.removeChild(firstRC);
                        } else {
                            break;
                        }
                    } else {
                        break;
                    }
                }
                const meaningful = Array.from(firstC.childNodes).filter(n => n.nodeType === 1 && n.localName !== "rPr");
                if (meaningful.length === 0) {
                    pNode.removeChild(firstC);
                    continue;
                } else {
                    break;
                }
            } else {
                break;
            }
        }
    }

    function collapseConsecutiveTabs(pNode) {
        let prevWasTab = false;
        const runs = Array.from(pNode.childNodes).filter(n => n.nodeType === 1 && n.localName === "r");
        for (let i = 0; i < runs.length; i++) {
            const r = runs[i];
            const rChildren = Array.from(r.childNodes).filter(n => n.nodeType === 1);
            for (let j = 0; j < rChildren.length; j++) {
                const child = rChildren[j];
                if (child.localName === "tab") {
                    if (prevWasTab) {
                        r.removeChild(child);
                    } else {
                        prevWasTab = true;
                    }
                } else if (child.localName === "t") {
                    if (child.textContent && child.textContent.trim()) {
                        prevWasTab = false;
                    }
                } else if (child.localName !== "rPr") {
                    prevWasTab = false;
                }
            }
            const meaningful = Array.from(r.childNodes).filter(n => n.nodeType === 1 && n.localName !== "rPr");
            if (meaningful.length === 0) {
                pNode.removeChild(r);
            }
        }
    }

    function resetParagraphFormattingForCompact(pNode, xmlDoc) {
        let pPr = getChildByTagName(pNode, "pPr");
        if (!pPr) {
            pPr = createWordElement(xmlDoc, "pPr");
            pNode.insertBefore(pPr, pNode.firstChild);
        }

        let ind = getChildByTagName(pPr, "ind");
        if (ind) pPr.removeChild(ind);
        ind = createWordElement(xmlDoc, "ind");
        ind.setAttributeNS(NS.w, "w:left", "0");
        ind.setAttributeNS(NS.w, "w:firstLine", "0");
        ind.setAttributeNS(NS.w, "w:right", "0");
        pPr.appendChild(ind);

        let jc = getChildByTagName(pPr, "jc");
        if (jc) pPr.removeChild(jc);
        jc = createWordElement(xmlDoc, "jc");
        jc.setAttributeNS(NS.w, "w:val", "left");
        pPr.appendChild(jc);
    }

    function setParagraphTabs(pNode, xmlDoc, tabPositions) {
        let pPr = getChildByTagName(pNode, "pPr");
        if (!pPr) {
            pPr = createWordElement(xmlDoc, "pPr");
            pNode.insertBefore(pPr, pNode.firstChild);
        }

        let oldTabs = getChildByTagName(pPr, "tabs");
        if (oldTabs) pPr.removeChild(oldTabs);

        if (!tabPositions || tabPositions.length === 0) return;

        const tabsElem = createWordElement(xmlDoc, "tabs");
        for (let i = 0; i < tabPositions.length; i++) {
            const pos = tabPositions[i];
            const tabElem = createWordElement(xmlDoc, "tab");
            tabElem.setAttributeNS(NS.w, "w:val", "left");
            tabElem.setAttributeNS(NS.w, "w:pos", pos.toString());
            tabsElem.appendChild(tabElem);
        }
        pPr.appendChild(tabsElem);
    }

    function mergeParagraphsWithTab(targetP, sourceParas, tabPositions, xmlDoc) {
        resetParagraphFormattingForCompact(targetP, xmlDoc);
        cleanTrailingWhitespaceAndTabs(targetP);

        for (let i = 0; i < sourceParas.length; i++) {
            const srcP = sourceParas[i];
            cleanLeadingWhitespaceAndTabs(srcP);
            cleanTrailingWhitespaceAndTabs(srcP);

            const parent = srcP.parentNode;

            // 1. Thêm một run chứa ký tự Tab vào targetP
            const tabR = createWordElement(xmlDoc, "r");
            const tabTag = createWordElement(xmlDoc, "tab");
            tabR.appendChild(tabTag);
            targetP.appendChild(tabR);

            // 2. Chuyển toàn bộ các phần tử con của srcP (trừ w:pPr) sang targetP
            const children = Array.from(srcP.childNodes);
            for (let j = 0; j < children.length; j++) {
                const child = children[j];
                if (child.nodeType === 1 && child.localName === "pPr") continue;
                targetP.appendChild(child);
            }

            // 3. Xóa source paragraph khỏi parent
            if (parent) {
                parent.removeChild(srcP);
            }
        }

        cleanTrailingWhitespaceAndTabs(targetP);
        collapseConsecutiveTabs(targetP);
        setParagraphTabs(targetP, xmlDoc, tabPositions);
    }

    function decideGroupLayout(paras, mode) {
        const n = paras.length;
        if (n <= 1) return paras.map(p => [p]);

        if (paras.some(hasHeavyContent)) {
            return paras.map(p => [p]);
        }

        const lengths = paras.map(getParagraphEstimatedLength);
        const maxLen = Math.max(...lengths);
        const sumLen = lengths.reduce((a, b) => a + b, 0);

        if (mode === "4_per_line") {
            return [paras];
        } else if (mode === "2_per_line") {
            if (n === 4) {
                return [[paras[0], paras[1]], [paras[2], paras[3]]];
            } else if (n === 2) {
                return [paras];
            } else if (n === 3) {
                return [[paras[0], paras[1]], [paras[2]]];
            } else {
                const rows = [];
                for (let i = 0; i < n; i += 2) {
                    rows.push(paras.slice(i, i + 2));
                }
                return rows;
            }
        } else if (mode === "1_per_line" || mode === "split") {
            return paras.map(p => [p]);
        } else {
            // Chế độ 'auto' (Auto 4 - 2 - 1)
            if (n === 4) {
                const can4Cols = (
                    maxLen <= 16 &&
                    sumLen <= 58 &&
                    lengths[0] <= 15 &&
                    lengths[1] <= 15 &&
                    lengths[2] <= 15 &&
                    lengths[3] <= 20
                );
                if (can4Cols) {
                    return [paras];
                }

                const can2Cols = (
                    lengths[0] <= 32 &&
                    lengths[1] <= 36 &&
                    (lengths[0] + lengths[1] <= 65) &&
                    lengths[2] <= 32 &&
                    lengths[3] <= 36 &&
                    (lengths[2] + lengths[3] <= 65)
                );
                if (can2Cols) {
                    return [[paras[0], paras[1]], [paras[2], paras[3]]];
                }

                return paras.map(p => [p]);
            } else if (n === 2) {
                if (lengths[0] <= 32 && lengths[1] <= 36 && sumLen <= 65) {
                    return [paras];
                } else {
                    return paras.map(p => [p]);
                }
            } else if (n === 3) {
                if (lengths[0] <= 18 && lengths[1] <= 18 && lengths[2] <= 22 && sumLen <= 54) {
                    return [paras];
                } else {
                    return paras.map(p => [p]);
                }
            } else {
                return paras.map(p => [p]);
            }
        }
    }

    function compactDocumentChoices(xmlDoc, layoutMode = "auto") {
        if (layoutMode === "split" || layoutMode === "1_per_line") {
            return { totalGroups: 0, compactedCount: 0 };
        }

        const body = getChildByTagName(xmlDoc.documentElement, "body");
        if (!body) return { totalGroups: 0, compactedCount: 0 };

        const allParas = window.XmlUtils.getAllParagraphsInDoc(body);
        const groups = [];
        let currentGroup = [];
        let currentIsLower = false;
        let currentExpectedIdx = 0;
        let currentParent = null;

        for (let i = 0; i < allParas.length; i++) {
            const p = allParas[i];
            const info = extractChoiceLabelInfo(p);
            const pParent = p.parentNode;

            if (info) {
                let isContinuation = false;
                if (currentGroup.length > 0 && info.isLower === currentIsLower && pParent === currentParent) {
                    if (info.index === currentExpectedIdx) {
                        isContinuation = true;
                    }
                }

                if (isContinuation) {
                    currentGroup.push(p);
                    currentExpectedIdx++;
                } else {
                    if (currentGroup.length >= 2) {
                        groups.push({ isLower: currentIsLower, paras: currentGroup });
                    }
                    if (info.index === 0) {
                        currentGroup = [p];
                        currentIsLower = info.isLower;
                        currentExpectedIdx = 1;
                        currentParent = pParent;
                    } else {
                        currentGroup = [];
                        currentParent = null;
                    }
                }
            } else {
                if (currentGroup.length >= 2) {
                    groups.push({ isLower: currentIsLower, paras: currentGroup });
                }
                currentGroup = [];
                currentParent = null;
            }
        }

        if (currentGroup.length >= 2) {
            groups.push({ isLower: currentIsLower, paras: currentGroup });
        }

        const totalGroups = groups.length;
        let compactedCount = 0;

        for (let i = 0; i < groups.length; i++) {
            const { paras } = groups[i];
            const rows = decideGroupLayout(paras, layoutMode);

            if (rows.length < paras.length) {
                compactedCount++;
            }

            for (let j = 0; j < rows.length; j++) {
                const row = rows[j];
                if (row.length <= 1) {
                    resetParagraphFormattingForCompact(row[0], xmlDoc);
                    continue;
                }

                const targetP = row[0];
                const sourceParas = row.slice(1);
                const numCols = row.length;
                let tabs;
                if (numCols === 4) {
                    tabs = TAB_STOPS_4_COLS;
                } else if (numCols === 2) {
                    tabs = TAB_STOPS_2_COLS;
                } else if (numCols === 3) {
                    tabs = TAB_STOPS_3_COLS;
                } else {
                    tabs = [];
                    for (let c = 1; c < numCols; c++) {
                        tabs.push(Math.round((9000 / numCols) * c));
                    }
                }

                mergeParagraphsWithTab(targetP, sourceParas, tabs, xmlDoc);
            }
        }

        return { totalGroups, compactedCount };
    }

    async function compactDocxChoices(zip, layoutMode = "auto") {
        const docXmlStr = await zip.file("word/document.xml").async("string");
        const xmlDoc = window.XmlUtils.parseXml(docXmlStr);

        const { totalGroups, compactedCount } = compactDocumentChoices(xmlDoc, layoutMode);

        const newDocXmlStr = window.XmlUtils.serializeXml(xmlDoc);
        zip.file("word/document.xml", newDocXmlStr);

        const modeMap = {
            "auto": "Tự động (Auto 4-2-1)",
            "4_per_line": "4 phương án / dòng",
            "2_per_line": "2 phương án / dòng",
            "1_per_line": "1 phương án / dòng",
            "split": "Tách 1 phương án / dòng"
        };
        const modeText = modeMap[layoutMode] || layoutMode;
        return {
            success: true,
            totalGroups,
            compactedCount,
            message: `Đã dồn dòng ${compactedCount}/${totalGroups} câu hỏi theo chế độ '${modeText}'.`
        };
    }

    window.ChoiceCompactor = {
        extractChoiceLabelInfo,
        hasHeavyContent,
        getParagraphEstimatedLength,
        decideGroupLayout,
        compactDocumentChoices,
        compactDocxChoices
    };
})(window);
