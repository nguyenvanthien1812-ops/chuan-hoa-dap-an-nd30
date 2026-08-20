/**
 * nd30_normalizer.js - Module chuẩn hóa văn bản hành chính theo Nghị định 30/2020/NĐ-CP
 * Pure Client-Side Implementation of nd30_normalizer.py
 */

(function(window) {
    const {
        NS,
        createWordElement,
        getChildByTagName,
        getChildrenByTagName,
        findDescendants,
        getParagraphText,
        setRunFont,
        setParagraphAlignment,
        setParagraphIndent,
        setParagraphSpacing,
        getAllParagraphsInDoc
    } = window.XmlUtils;

    const HEADING_RE = /^\s*(?:Chương\s+[IVXivx\d]+|Mục\s+\d+|Điều\s+\d+|[A-ZĐ\d]+\.\s+|[a-zđ]\)\s+|Part\s+[A-Z\d])/i;
    const MAIN_HEADING_RE = /^\s*(?:Chương\s+[IVXivx\d]+|Mục\s+\d+|Điều\s+\d+|[A-ZĐ\d]+\.\s+)/i;
    const SUB_HEADING_RE = /^\s*([a-zđ]\)\s+)/i;

    function applyMarginSettings(xmlDoc) {
        const body = getChildByTagName(xmlDoc.documentElement, "body");
        if (!body) return;

        const sectPrs = Array.from(xmlDoc.getElementsByTagNameNS(NS.w, "sectPr"));
        for (let i = 0; i < sectPrs.length; i++) {
            const sectPr = sectPrs[i];

            // 1. Page Size (A4: 21.0cm x 29.7cm = 11906 x 16838 dxa)
            let pgSz = getChildByTagName(sectPr, "pgSz");
            if (!pgSz) {
                pgSz = createWordElement(xmlDoc, "pgSz");
                sectPr.appendChild(pgSz);
            }
            pgSz.setAttributeNS(NS.w, "w:w", "11906");
            pgSz.setAttributeNS(NS.w, "w:h", "16838");

            // 2. Margins (Top 2.0cm, Bottom 2.0cm, Left 3.0cm, Right 1.5cm)
            // 1 cm = 567 dxa -> 2.0cm = 1134, 3.0cm = 1701, 1.5cm = 850
            let pgMar = getChildByTagName(sectPr, "pgMar");
            if (!pgMar) {
                pgMar = createWordElement(xmlDoc, "pgMar");
                sectPr.appendChild(pgMar);
            }
            pgMar.setAttributeNS(NS.w, "w:top", "1134");
            pgMar.setAttributeNS(NS.w, "w:bottom", "1134");
            pgMar.setAttributeNS(NS.w, "w:left", "1701");
            pgMar.setAttributeNS(NS.w, "w:right", "850");
            pgMar.setAttributeNS(NS.w, "w:header", "720");
            pgMar.setAttributeNS(NS.w, "w:footer", "720");
            pgMar.setAttributeNS(NS.w, "w:gutter", "0");
        }
    }

    async function applyPageNumbering(zip, xmlDoc) {
        const body = getChildByTagName(xmlDoc.documentElement, "body");
        if (!body) return;

        // 1. Bật different_first_page_header_footer (<w:titlePg/>) trong tất cả sectPr
        const sectPrs = Array.from(xmlDoc.getElementsByTagNameNS(NS.w, "sectPr"));
        for (let i = 0; i < sectPrs.length; i++) {
            const sectPr = sectPrs[i];
            let titlePg = getChildByTagName(sectPr, "titlePg");
            if (!titlePg) {
                titlePg = createWordElement(xmlDoc, "titlePg");
                sectPr.appendChild(titlePg);
            }
        }

        // 2. Kiểm tra xem có header hiện tại không
        let headerFile = zip.file("word/header1.xml") || zip.file("word/header2.xml");
        let headerDoc;
        let headerPath = "word/header1.xml";

        if (headerFile) {
            const headerXmlStr = await headerFile.async("string");
            headerDoc = window.XmlUtils.parseXml(headerXmlStr);
            headerPath = headerFile.name;
        } else {
            // Tạo header1.xml mới
            const newHeaderXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
</w:hdr>`;
            headerDoc = window.XmlUtils.parseXml(newHeaderXml);
        }

        const hdrRoot = headerDoc.documentElement;
        // Xóa các p cũ
        while (hdrRoot.firstChild) {
            hdrRoot.removeChild(hdrRoot.firstChild);
        }

        // Tạo đoạn văn số trang căn giữa, font Times New Roman 13pt
        const p = createWordElement(headerDoc, "p");
        const pPr = createWordElement(headerDoc, "pPr");
        const jc = createWordElement(headerDoc, "jc");
        jc.setAttributeNS(NS.w, "w:val", "center");
        pPr.appendChild(jc);

        const spacing = createWordElement(headerDoc, "spacing");
        spacing.setAttributeNS(NS.w, "w:before", "0");
        spacing.setAttributeNS(NS.w, "w:after", "0");
        pPr.appendChild(spacing);
        p.appendChild(pPr);

        const r = createWordElement(headerDoc, "r");
        setRunFont(r, headerDoc, "Times New Roman", 13);

        const fldSimple = createWordElement(headerDoc, "fldSimple");
        fldSimple.setAttributeNS(NS.w, "w:instr", "PAGE");
        r.appendChild(fldSimple);
        p.appendChild(r);

        hdrRoot.appendChild(p);

        zip.file(headerPath, window.XmlUtils.serializeXml(headerDoc));

        // Đảm bảo document.xml.rels và sectPr liên kết tới header
        let relsFile = zip.file("word/_rels/document.xml.rels");
        if (relsFile) {
            const relsStr = await relsFile.async("string");
            const relsDoc = window.XmlUtils.parseXml(relsStr);
            const relsRoot = relsDoc.documentElement;

            let existingHeaderRelId = null;
            const rels = relsRoot.getElementsByTagName("Relationship");
            for (let i = 0; i < rels.length; i++) {
                const type = rels[i].getAttribute("Type");
                if (type && type.includes("/header")) {
                    existingHeaderRelId = rels[i].getAttribute("Id");
                    break;
                }
            }

            if (!existingHeaderRelId) {
                existingHeaderRelId = "rIdHeaderNd30";
                const newRel = relsDoc.createElement("Relationship");
                newRel.setAttribute("Id", existingHeaderRelId);
                newRel.setAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header");
                newRel.setAttribute("Target", "header1.xml");
                relsRoot.appendChild(newRel);
                zip.file("word/_rels/document.xml.rels", window.XmlUtils.serializeXml(relsDoc));
            }

            // Gán headerReference vào sectPr
            for (let i = 0; i < sectPrs.length; i++) {
                const sectPr = sectPrs[i];
                let headerRef = getChildByTagName(sectPr, "headerReference");
                if (!headerRef) {
                    headerRef = createWordElement(xmlDoc, "headerReference");
                    headerRef.setAttributeNS(NS.w, "w:type", "default");
                    headerRef.setAttributeNS(NS.r, "r:id", existingHeaderRelId);
                    sectPr.insertBefore(headerRef, sectPr.firstChild);
                }
            }
        }
    }

    async function normalizeDocumentNd30(zip, options = {}) {
        const optMargin = options.margin !== false;
        const optPageNum = options.page_num !== false;
        const optFont = options.font !== false;
        const optParagraph = options.paragraph !== false;
        const optHeadings = options.headings !== false;
        const optConvertFonts = options.convert_fonts !== false;
        const optCleanPunctuation = options.clean_punctuation !== false;
        const optRemoveEmptyParas = options.remove_empty_paras !== false;

        const docXmlStr = await zip.file("word/document.xml").async("string");
        const xmlDoc = window.XmlUtils.parseXml(docXmlStr);

        // 0. Bác sĩ văn bản & Chuyển mã phông cũ TCVN3/VNI sang Unicode
        let cleanInfo = { cleanedCount: 0, convertedFontsCount: 0 };
        if (window.TextCleaner) {
            cleanInfo = window.TextCleaner.cleanDocumentFormatting(xmlDoc, {
                convert_fonts: optConvertFonts,
                clean_punctuation: optCleanPunctuation,
                remove_empty_paras: optRemoveEmptyParas
            });
        }

        // 1. Page Margins & A4
        if (optMargin) {
            applyMarginSettings(xmlDoc);
        }

        // 2. Page Numbering
        if (optPageNum) {
            await applyPageNumbering(zip, xmlDoc);
        }

        const body = getChildByTagName(xmlDoc.documentElement, "body");
        if (!body) {
            return { success: true, message: "Đã hoàn thành chuẩn hóa NĐ 30." };
        }

        // 3. Format Body Paragraphs
        if (optFont || optParagraph || optHeadings) {
            const paragraphs = Array.from(body.getElementsByTagNameNS(NS.w, "p"));

            for (let i = 0; i < paragraphs.length; i++) {
                const p = paragraphs[i];
                const text = getParagraphText(p).trim();
                if (!text) continue;

                const textLower = text.toLowerCase();
                const isSlogan1 = textLower.includes("cộng hòa xã hội chủ nghĩa việt nam");
                const isSlogan2 = textLower.includes("độc lập - tự do - hạnh phúc");

                // Slogan & National Motto
                if (isSlogan1 || isSlogan2) {
                    setParagraphAlignment(p, xmlDoc, "center");
                    setParagraphIndent(p, xmlDoc, 0, 0, 0);
                    setParagraphSpacing(p, xmlDoc, 0, 2, 1.15);

                    const size = isSlogan1 ? 13 : 14;
                    const runs = getChildrenByTagName(p, "r");
                    for (let r = 0; r < runs.length; r++) {
                        setRunFont(runs[r], xmlDoc, "Times New Roman", size);
                        const rPr = window.XmlUtils.getOrCreateRPr(runs[r], xmlDoc);
                        if (!getChildByTagName(rPr, "b")) {
                            rPr.appendChild(createWordElement(xmlDoc, "b"));
                        }
                    }
                    continue;
                }

                // Main Title & Headings
                const isMainTitle = (text === text.toUpperCase() && text.length < 100 && /[A-ZĐÀ-Ỹ]/.test(text));
                const isHeading = HEADING_RE.test(text);

                if (isMainTitle || isHeading) {
                    if (optHeadings) {
                        setParagraphIndent(p, xmlDoc, 0, 0, 0);
                        setParagraphSpacing(p, xmlDoc, 6, 4, 1.15);

                        if (isMainTitle) {
                            setParagraphAlignment(p, xmlDoc, "center");
                            const runs = getChildrenByTagName(p, "r");
                            for (let r = 0; r < runs.length; r++) {
                                setRunFont(runs[r], xmlDoc, "Times New Roman", 14);
                                const rPr = window.XmlUtils.getOrCreateRPr(runs[r], xmlDoc);
                                if (!getChildByTagName(rPr, "b")) rPr.appendChild(createWordElement(xmlDoc, "b"));
                            }
                        } else if (MAIN_HEADING_RE.test(text)) {
                            const runs = getChildrenByTagName(p, "r");
                            for (let r = 0; r < runs.length; r++) {
                                setRunFont(runs[r], xmlDoc, "Times New Roman", 14);
                                const rPr = window.XmlUtils.getOrCreateRPr(runs[r], xmlDoc);
                                if (!getChildByTagName(rPr, "b")) rPr.appendChild(createWordElement(xmlDoc, "b"));
                            }
                        } else if (SUB_HEADING_RE.test(text)) {
                            const runs = getChildrenByTagName(p, "r");
                            for (let r = 0; r < runs.length; r++) {
                                setRunFont(runs[r], xmlDoc, "Times New Roman", 13);
                                const rPr = window.XmlUtils.getOrCreateRPr(runs[r], xmlDoc);
                                if (!getChildByTagName(rPr, "i")) rPr.appendChild(createWordElement(xmlDoc, "i"));
                            }
                        } else {
                            const runs = getChildrenByTagName(p, "r");
                            for (let r = 0; r < runs.length; r++) {
                                setRunFont(runs[r], xmlDoc, "Times New Roman", 13);
                                const rPr = window.XmlUtils.getOrCreateRPr(runs[r], xmlDoc);
                                if (!getChildByTagName(rPr, "b")) rPr.appendChild(createWordElement(xmlDoc, "b"));
                            }
                        }
                    } else if (optFont) {
                        const runs = getChildrenByTagName(p, "r");
                        for (let r = 0; r < runs.length; r++) {
                            setRunFont(runs[r], xmlDoc, "Times New Roman");
                        }
                    }
                    continue;
                }

                // General body paragraph
                if (optParagraph) {
                    const pPr = getChildByTagName(p, "pPr");
                    const jc = pPr ? getChildByTagName(pPr, "jc") : null;
                    const alignVal = jc ? (jc.getAttributeNS(NS.w, "val") || jc.getAttribute("w:val")) : "left";

                    if (alignVal !== "center" && alignVal !== "right") {
                        setParagraphAlignment(p, xmlDoc, "both");
                        // Thụt đầu dòng 1.0cm = 567 dxa
                        setParagraphIndent(p, xmlDoc, 0, 567, 0);
                    }
                    setParagraphSpacing(p, xmlDoc, 0, 6, 1.2);
                }

                if (optFont) {
                    const runs = getChildrenByTagName(p, "r");
                    for (let r = 0; r < runs.length; r++) {
                        setRunFont(runs[r], xmlDoc, "Times New Roman", 14);
                    }
                }
            }
        }

        // 4. Format Tables
        if (optFont) {
            const tables = Array.from(body.getElementsByTagNameNS(NS.w, "tbl"));
            for (let t = 0; t < tables.length; t++) {
                const paras = tables[t].getElementsByTagNameNS(NS.w, "p");
                for (let p = 0; p < paras.length; p++) {
                    setParagraphIndent(paras[p], xmlDoc, 0, 0, 0);
                    setParagraphSpacing(paras[p], xmlDoc, 0, 2, 1.15);

                    const runs = getChildrenByTagName(paras[p], "r");
                    for (let r = 0; r < runs.length; r++) {
                        setRunFont(runs[r], xmlDoc, "Times New Roman", 12);
                    }
                }
            }
        }

        const newDocXmlStr = window.XmlUtils.serializeXml(xmlDoc);
        zip.file("word/document.xml", newDocXmlStr);

        let extraMsg = "";
        if (cleanInfo.convertedFontsCount > 0) {
            extraMsg += ` Đã chuyển ${cleanInfo.convertedFontsCount} đoạn phông cũ sang Unicode.`;
        }
        if (cleanInfo.cleanedCount > 0) {
            extraMsg += ` Đã sửa ${cleanInfo.cleanedCount} lỗi dấu câu.`;
        }

        return {
            success: true,
            cleanedCount: cleanInfo.cleanedCount,
            convertedFontsCount: cleanInfo.convertedFontsCount,
            message: `Đã hoàn thành chuẩn hóa văn bản theo Nghị định 30/2020/NĐ-CP!${extraMsg}`
        };
    }

    window.Nd30Normalizer = {
        applyMarginSettings,
        applyPageNumbering,
        normalizeDocumentNd30
    };
})(window);
