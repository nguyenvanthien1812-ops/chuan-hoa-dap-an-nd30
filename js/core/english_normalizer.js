/**
 * english_normalizer.js - Module chuẩn hóa đề thi tiếng Anh chuyên sâu
 * Pure Client-Side Implementation of english_normalizer.py
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

    const QUESTION_RE = /^[\s.:\-*]*(?:Câu|Questions?|Bài|câu|bài)\s*(?:hỏi\s+)?(\d+)\s*(?:\([^)]*\))?\s*[:.\-)?\s]?/i;
    const PLAIN_NUMBER_RE = /^[\s.:\-*]*(\d+)\s*(?:\([^)]*\))?\s*[.:)]\s*(?!\d)/;
    const SECTION_RE = /^\s*(?:(?:phần|part)\s*([IVXLCDM]+|\d+|[A-D])|([IVXLCDM]+|\d+|[A-D])\s*[.:]\s*(?=.*?(?:trắc nghiệm|tự luận|đúng[\s/]*sai|trả lời ngắn|chọn đ\s*\()))/i;
    const INLINE_MC_RE = /(?<![A-Z])(?:\()?(\*?\s*[A-H]\s*\*?)\s*[.:)/\-]\s*\*?/g;

    function convertAllQuestionTablesToParagraphs(xmlDoc) {
        const body = getChildByTagName(xmlDoc.documentElement, "body");
        if (!body) return;

        const tables = Array.from(body.getElementsByTagNameNS(NS.w, "tbl"));
        const tablesToProcess = [];

        for (let t = 0; t < tables.length; t++) {
            const table = tables[t];
            let hasQuestion = false;
            let hasChoices = false;

            const cells = table.getElementsByTagNameNS(NS.w, "tc");
            for (let c = 0; c < cells.length; c++) {
                const text = cells[c].textContent.trim();
                if (/^[\s.:\-*]*(?:Câu|Questions?|Bài|câu|bài)\s*(?:hỏi\s+)?(\d+)/i.test(text)) {
                    hasQuestion = true;
                }
                if (/(?<![A-Z])(?:\()?(\*?\s*[A-H]\s*\*?)\s*[.:)/\-]/i.test(text)) {
                    hasChoices = true;
                }
            }

            if (hasQuestion) {
                const isTf = (cells.length >= 3 && !hasChoices);
                if (!isTf) {
                    tablesToProcess.push(table);
                }
            }
        }

        for (let t = 0; t < tablesToProcess.length; t++) {
            const table = tablesToProcess[t];
            const parent = table.parentNode;
            if (!parent) continue;

            const paras = table.getElementsByTagNameNS(NS.w, "p");
            for (let p = 0; p < paras.length; p++) {
                parent.insertBefore(paras[p], table);
            }
            parent.removeChild(table);
        }
    }

    function breakTfTablesToParagraphs(xmlDoc) {
        const body = getChildByTagName(xmlDoc.documentElement, "body");
        if (!body) return;

        const tables = Array.from(body.getElementsByTagNameNS(NS.w, "tbl"));
        for (let t = 0; t < tables.length; t++) {
            const table = tables[t];
            const rows = getChildrenByTagName(table, "tr");
            let isTfTable = false;

            for (let r = 0; r < rows.length; r++) {
                const cells = getChildrenByTagName(rows[r], "tc");
                if (cells.length >= 2) {
                    const t0 = cells[0].textContent.trim();
                    if (/^(?:Questions?|Câu)\s*\d+/i.test(t0) || /^[a-d]\s*[\).]/i.test(t0)) {
                        isTfTable = true;
                        break;
                    }
                }
            }

            if (isTfTable) {
                const parent = table.parentNode;
                if (!parent) continue;
                for (let r = 0; r < rows.length; r++) {
                    const cells = getChildrenByTagName(rows[r], "tc");
                    for (let c = 0; c < cells.length; c++) {
                        const paras = getChildrenByTagName(cells[c], "p");
                        for (let p = 0; p < paras.length; p++) {
                            parent.insertBefore(paras[p], table);
                        }
                    }
                }
                parent.removeChild(table);
            }
        }
    }

    async function normalizeEnglishDocx(zip, options = {}) {
        const autoKey = options.auto_key !== false;
        const boldAsCorrect = options.bold_as_correct === true;
        const choiceLayout = options.choice_layout || "split";
        const choiceIndent = options.choice_indent !== false;
        const indentDxa = choiceIndent ? (options.indent_dxa !== undefined ? options.indent_dxa : 567) : 0;

        const docXmlStr = await zip.file("word/document.xml").async("string");
        const xmlDoc = window.XmlUtils.parseXml(docXmlStr);

        window.Normalizer.splitParagraphsAtBr(xmlDoc);
        convertAllQuestionTablesToParagraphs(xmlDoc);
        breakTfTablesToParagraphs(xmlDoc);
        window.Normalizer.convertChoiceTablesToParagraphs(xmlDoc);

        let answerKeys = {};
        let keyCount = 0;
        if (autoKey) {
            answerKeys = window.Normalizer.extractAndRemoveAnswerKey(xmlDoc);
            keyCount = Object.values(answerKeys).reduce((acc, sec) => acc + Object.keys(sec).length, 0);
        }

        const body = getChildByTagName(xmlDoc.documentElement, "body");
        let paragraphs = getAllParagraphsInDoc(body);

        // Offsets
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

        while (pIdx < paragraphs.length) {
            const p = paragraphs[pIdx];
            const text = getParagraphText(p).trim();

            if (!text) {
                pIdx++;
                continue;
            }

            // Section
            const secMatch = SECTION_RE.exec(text);
            if (secMatch) {
                currentSectionIdx++;
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

            // Question
            const qMatch = QUESTION_RE.exec(text) || PLAIN_NUMBER_RE.exec(text);
            if (qMatch) {
                currentQNum = parseInt(qMatch[1], 10);
            }

            // Multiple Choice (A. B. C. D.)
            const mUpper = /^(?:\()?([A-H])\s*[.:)/\-]\s*/.exec(text);
            if (mUpper && currentQNum > 0) {
                const label = mUpper[1].toUpperCase();
                let isCorrect = false;
                const correctVal = window.Normalizer.getAnswerForQuestion(answerKeys, currentSectionIdx, currentQNum, offsets);
                if (correctVal && label === correctVal.toUpperCase()) {
                    isCorrect = true;
                }
                if (!correctVal && boldAsCorrect) {
                    // Check bold
                    const runs = getChildrenByTagName(p, "r");
                    for (let r = 0; r < runs.length; r++) {
                        const rPr = getChildByTagName(runs[r], "rPr");
                        if (rPr && getChildByTagName(rPr, "b")) {
                            isCorrect = true;
                            break;
                        }
                    }
                }
                if (isCorrect) {
                    const wtElems = findDescendants(p, "t");
                    for (let w = 0; w < wtElems.length; w++) {
                        const t = wtElems[w].textContent || "";
                        const pat = new RegExp(`^(\\s*)(?:\\()?${label}(\\s*[.:)/\\-])`, 'i');
                        if (pat.test(t)) {
                            wtElems[w].textContent = t.replace(pat, `$1${label}*$2`);
                            break;
                        }
                    }
                }
                pIdx++;
                continue;
            }

            pIdx++;
        }

        // Choice Compacting
        let compactedMsg = "";
        if (choiceLayout && (choiceLayout !== "split" || indentDxa > 0)) {
            const { compactedCount } = window.ChoiceCompactor.compactDocumentChoices(xmlDoc, choiceLayout, indentDxa);
            if (compactedCount > 0) {
                compactedMsg = ` Đã dồn dòng ${compactedCount} câu hỏi tiếng Anh.`;
            }
        }

        const newDocXmlStr = window.XmlUtils.serializeXml(xmlDoc);
        zip.file("word/document.xml", newDocXmlStr);

        return {
            success: true,
            keyCount,
            answerKeys,
            message: `Chuẩn hóa đề thi tiếng Anh thành công! Đã quét được ${keyCount} đáp án.${compactedMsg}`
        };
    }

    window.EnglishNormalizer = {
        normalizeEnglishDocx
    };
})(window);
