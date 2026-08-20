/**
 * text_cleaner.js - Module chuyển mã phông cũ (TCVN3, VNI) sang Unicode & Bác sĩ văn bản tiếng Việt
 * Pure Client-Side Implementation
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

    // =========================================================================
    // 1. BẢNG ÁNH XẠ KÝ TỰ TCVN3 (.VnTime) -> UNICODE
    // =========================================================================
    const TCVN3_TO_UNICODE = {
        // Nguyên âm có dấu trong TCVN3
        "\u00B8": "á", "\u00B5": "à", "\u00B6": "ả", "\u00B7": "ã", "\u00B9": "ạ",
        "\u00A8": "ă", "\u00BE": "ắ", "\u00BB": "ằ", "\u00BC": "ẳ", "\u00BD": "ẵ", "\u00C6": "ặ",
        "\u00A9": "â", "\u00CA": "ấ", "\u00C7": "ầ", "\u00C8": "ẩ", "\u00C9": "ẫ", "\u00CB": "ậ",
        "\u00D0": "é", "\u00CC": "è", "\u00CE": "ẻ", "\u00CF": "ẽ", "\u00D1": "ẹ",
        "\u00AA": "ê", "\u00D5": "ế", "\u00D2": "ề", "\u00D3": "ể", "\u00D4": "ễ", "\u00D6": "ệ",
        "\u00DD": "í", "\u00D7": "ì", "\u00D8": "ỉ", "\u00DC": "ĩ", "\u00DE": "ị",
        "\u00DF": "ó", "\u00E1": "ò", "\u00E2": "ỏ", "\u00E3": "õ", "\u00E4": "ọ",
        "\u00AB": "ô", "\u00E8": "ố", "\u00E5": "ồ", "\u00E6": "ổ", "\u00E7": "ỗ", "\u00E9": "ộ",
        "\u00AC": "ơ", "\u00ED": "ớ", "\u00EA": "ờ", "\u00EB": "ở", "\u00EC": "ỡ", "\u00EE": "ợ",
        "\u00F3": "ú", "\u00EF": "ù", "\u00F1": "ủ", "\u00F2": "ũ", "\u00F4": "ụ",
        "\u00AD": "ư", "\u00F8": "ứ", "\u00F5": "ừ", "\u00F6": "ử", "\u00F7": "ữ", "\u00F9": "ự",
        "\u00FD": "ý", "\u00FA": "ỳ", "\u00FB": "ỷ", "\u00FC": "ỹ", "\u00FE": "ỵ",
        "\u00AE": "đ",

        // Ký tự chữ hoa TCVN3 (.VnTimeH)
        "\u00A1": "Ă", "\u00A2": "Â", "\u00A3": "Ê", "\u00A4": "Ô", "\u00A5": "Ơ", "\u00A6": "Ư", "\u00A7": "Đ"
    };

    // =========================================================================
    // 2. BẢNG ÁNH XẠ KÝ TỰ VNI-TIMES -> UNICODE
    // =========================================================================
    const VNI_TO_UNICODE_PAIRS = [
        // Chữ hoa 2 ký tự và có dấu
        [/AÊN/g, "ĂN"], [/AÉ/g, "Ắ"], [/AÈ/g, "Ằ"], [/AÚ/g, "Ẳ"], [/AÜ/g, "Ẵ"], [/AË/g, "Ặ"], [/AÊ/g, "Ă"],
        [/AÁ/g, "Ấ"], [/AÀ/g, "Ầ"], [/AẨ/g, "Ẩ"], [/AẪ/g, "Ẫ"], [/AẬ/g, "Ậ"], [/AÂ/g, "Â"],
        [/AÙ/g, "Á"], [/AØ/g, "À"], [/AÛ/g, "Ả"], [/AÕ/g, "Ã"], [/AÏ/g, "Ạ"],
        [/EÁ/g, "Ế"], [/EÀ/g, "Ề"], [/EẨ/g, "Ể"], [/EẪ/g, "Ễ"], [/EẬ/g, "Ệ"], [/EÂ/g, "Ê"],
        [/EÙ/g, "É"], [/EØ/g, "È"], [/EÛ/g, "Ẻ"], [/EÕ/g, "Ẽ"], [/EÏ/g, "Ẹ"],
        [/IÙ/g, "Í"], [/IØ/g, "Ì"], [/IÛ/g, "Ỉ"], [/IÕ/g, "Ĩ"], [/IÏ/g, "Ị"],
        [/OÁ/g, "Ố"], [/OÀ/g, "Ồ"], [/OẨ/g, "Ổ"], [/OẪ/g, "Ỗ"], [/OẬ/g, "Ộ"], [/OÂ/g, "Ô"],
        [/ÔÙ/g, "Ớ"], [/ÔØ/g, "Ờ"], [/ÔÛ/g, "Ở"], [/ÔÕ/g, "Ỡ"], [/ÔÏ/g, "Ợ"], [/Ô/g, "Ơ"],
        [/OÙ/g, "Ó"], [/OØ/g, "Ò"], [/OÛ/g, "Ỏ"], [/OÕ/g, "Õ"], [/OÏ/g, "Ọ"],
        [/ƯÙ/g, "Ứ"], [/ƯØ/g, "Ừ"], [/ƯÛ/g, "Ử"], [/ƯÕ/g, "Ữ"], [/ƯÏ/g, "Ự"], [/ƯÙ/g, "Ứ"], [/Ư/g, "Ư"],
        [/UÙ/g, "Ú"], [/UØ/g, "Ù"], [/UÛ/g, "Ủ"], [/UÕ/g, "Ũ"], [/UÏ/g, "Ụ"],
        [/YÙ/g, "Ý"], [/YØ/g, "Ỳ"], [/YÛ/g, "Ỷ"], [/YÕ/g, "Ỹ"], [/YÏ/g, "Ỵ"],
        [/Ñ/g, "Đ"],

        // Chữ thường 2 ký tự và có dấu
        [/aé/g, "ắ"], [/aè/g, "ằ"], [/aú/g, "ẳ"], [/aü/g, "ẵ"], [/aë/g, "ặ"], [/aê/g, "ă"],
        [/aá/g, "ấ"], [/aà/g, "ầ"], [/aå/g, "ẩ"], [/aã/g, "ẫ"], [/aä/g, "ậ"], [/aâ/g, "â"],
        [/aù/g, "á"], [/aø/g, "à"], [/aû/g, "ả"], [/aõ/g, "ã"], [/aï/g, "ạ"],
        [/eá/g, "ế"], [/eà/g, "ề"], [/eå/g, "ể"], [/eã/g, "ễ"], [/eä/g, "ệ"], [/eâ/g, "ê"],
        [/eù/g, "é"], [/eø/g, "è"], [/eû/g, "ẻ"], [/eõ/g, "ẽ"], [/eï/g, "ẹ"],
        [/oá/g, "ố"], [/oà/g, "ồ"], [/oå/g, "ổ"], [/oã/g, "ỗ"], [/oä/g, "ộ"], [/oâ/g, "ô"],
        [/ôù/g, "ớ"], [/ôø/g, "ờ"], [/ôû/g, "ở"], [/ôõ/g, "ỡ"], [/ôï/g, "ợ"], [/ô/g, "ơ"],
        [/où/g, "ó"], [/oø/g, "ò"], [/oû/g, "ỏ"], [/oõ/g, "õ"], [/oï/g, "ọ"],
        [/uù/g, "ú"], [/uø/g, "ù"], [/uû/g, "ủ"], [/uõ/g, "ũ"], [/uï/g, "ụ"],
        [/öù/g, "ứ"], [/öø/g, "ừ"], [/öû/g, "ử"], [/öõ/g, "ữ"], [/öï/g, "ự"], [/ö/g, "ư"],
        [/yù/g, "ý"], [/yø/g, "ỳ"], [/yû/g, "ỷ"], [/yõ/g, "ỹ"], [/yï/g, "ỵ"], [/î/g, "ỵ"],
        [/iù/g, "í"], [/iø/g, "ì"], [/iû/g, "ỉ"], [/iõ/g, "ĩ"], [/iï/g, "ị"],
        [/æ/g, "ỉ"], [/ó/g, "ĩ"], [/ñ/g, "đ"]
    ];

    function isTcvn3Font(fontName) {
        if (!fontName) return false;
        const fn = fontName.toLowerCase();
        return fn.startsWith(".vn") || fn.startsWith("vn") || fn.includes("vntime") || fn.includes("vnarial");
    }

    function isVniFont(fontName) {
        if (!fontName) return false;
        const fn = fontName.toLowerCase();
        return fn.startsWith("vni-") || fn.startsWith("vni ");
    }

    function convertTcvn3ToUnicode(text) {
        if (!text) return "";
        let result = "";
        for (let i = 0; i < text.length; i++) {
            const char = text[i];
            result += TCVN3_TO_UNICODE[char] !== undefined ? TCVN3_TO_UNICODE[char] : char;
        }
        return result;
    }

    function convertVniToUnicode(text) {
        if (!text) return "";
        let res = text;
        for (let i = 0; i < VNI_TO_UNICODE_PAIRS.length; i++) {
            const [pat, rep] = VNI_TO_UNICODE_PAIRS[i];
            res = res.replace(pat, rep);
        }
        return res;
    }

    // =========================================================================
    // 3. BÁC SĨ VĂN BẢN (VIETNAMESE TYPOGRAPHY & PUNCTUATION CLEANER)
    // =========================================================================
    function cleanVietnamesePunctuation(text) {
        if (!text) return "";

        let clean = text;

        // 1. Xóa khoảng trắng thừa giữa các từ (nhiều space liên tiếp -> 1 space)
        clean = clean.replace(/[ \t\xa0]{2,}/g, ' ');

        // 2. Xóa khoảng trắng đứng TRƯỚC các dấu câu: , . : ; ? ! ) ] } ” " %
        clean = clean.replace(/[ \t\xa0]+([,.:;?!%\)\]\}”])/g, '$1');

        // 3. Xóa khoảng trắng đứng SAU các dấu mở: ( [ { “
        clean = clean.replace(/([\(\[\{“])[ \t\xa0]+/g, '$1');

        // 4. Đảm bảo CÓ 1 khoảng trắng SAU các dấu câu , : ; ? ! nếu theo sau là chữ cái
        clean = clean.replace(/([,;?!])([^\s\d,;?!%\)\]\}”"'\n\r])/g, '$1 $2');

        // 5. Đối với dấu hai chấm (:): Thêm khoảng trắng sau nếu không phải định dạng giờ (12:30) hay url (http:)
        clean = clean.replace(/(?<!\d):(?!\d)(?!\/)(?!\s)([^\s\n\r])/g, ': $1');

        // 6. Đối với dấu chấm (.): Thêm khoảng trắng sau nếu không phải số thập phân (3.14) hay số thứ tự (1. 2.) hay tên file
        clean = clean.replace(/(?<!\d)\.(?!\d)(?![a-zA-Z0-9_\-]+\.[a-zA-Z0-9])([^\s\d\.\)\]\}”"'\n\r])/g, '. $1');

        // 7. Chuẩn hóa dấu ngoặc kép chuẩn tiếng Việt: thay thế các cặp dấu hỗn loạn
        clean = clean.replace(/“\s+/g, '“').replace(/\s+”/g, '”');

        // 8. Chuẩn hóa dấu gạch ngang trong ngày tháng năm (VD: ngày 05 - 03 - 2020 -> ngày 05/03/2020)
        clean = clean.replace(/ngày\s+(\d{1,2})\s*[-/]\s*(\d{1,2})\s*[-/]\s*(\d{4})/gi, 'ngày $1 tháng $2 năm $3');

        // 9. Xóa khoảng trắng ở đầu và cuối dòng
        return clean.trim();
    }

    // =========================================================================
    // 4. LÀM SẠCH VÀ CHUYỂN ĐỔI TOÀN BỘ FILE WORD XML
    // =========================================================================
    function cleanDocumentFormatting(xmlDoc, options = {}) {
        const convertFonts = options.convert_fonts !== false;
        const cleanPunctuation = options.clean_punctuation !== false;
        const removeEmptyParas = options.remove_empty_paras !== false;

        const body = getChildByTagName(xmlDoc.documentElement, "body");
        if (!body) return { cleanedCount: 0, convertedFontsCount: 0 };

        let convertedFontsCount = 0;
        let cleanedCount = 0;

        const allParas = getAllParagraphsInDoc(body);

        // 1. Quét từng đoạn văn để chuyển mã phông và làm sạch dấu câu
        for (let i = 0; i < allParas.length; i++) {
            const p = allParas[i];
            const runs = getChildrenByTagName(p, "r");

            for (let r = 0; r < runs.length; r++) {
                const run = runs[r];
                const rPr = getChildByTagName(run, "rPr");
                const tEl = getChildByTagName(run, "t");

                if (!tEl || !tEl.textContent) continue;

                let originalText = tEl.textContent;
                let text = originalText;
                let fontName = "";

                if (rPr) {
                    const rFonts = getChildByTagName(rPr, "rFonts");
                    if (rFonts) {
                        fontName = rFonts.getAttributeNS(NS.w, "ascii") ||
                                   rFonts.getAttribute("w:ascii") ||
                                   rFonts.getAttributeNS(NS.w, "hAnsi") ||
                                   rFonts.getAttribute("w:hAnsi") || "";
                    }
                }

                // A. Chuyển đổi mã phông cũ
                if (convertFonts && fontName) {
                    if (isTcvn3Font(fontName)) {
                        text = convertTcvn3ToUnicode(text);
                        // Đổi phông chữ sang Times New Roman
                        window.XmlUtils.setRunFont(run, xmlDoc, "Times New Roman");
                        convertedFontsCount++;
                    } else if (isVniFont(fontName)) {
                        text = convertVniToUnicode(text);
                        window.XmlUtils.setRunFont(run, xmlDoc, "Times New Roman");
                        convertedFontsCount++;
                    }
                }

                // B. Sửa lỗi chính tả & dấu câu tiếng Việt
                if (cleanPunctuation) {
                    const cleaned = cleanVietnamesePunctuation(text);
                    if (cleaned !== text) {
                        text = cleaned;
                        cleanedCount++;
                    }
                }

                if (text !== originalText) {
                    tEl.textContent = text;
                }
            }
        }

        // 2. Xóa các đoạn văn rỗng liên tiếp (nhiều Enter thừa) & trang trắng ở cuối
        if (removeEmptyParas) {
            let consecutiveEmpty = 0;
            const directParas = Array.from(body.childNodes).filter(n => n.nodeType === 1 && n.localName === "p");

            for (let i = 0; i < directParas.length; i++) {
                const p = directParas[i];
                const text = getParagraphText(p).trim();
                const hasSpecial = findDescendants(p, "drawing").length > 0 ||
                                   findDescendants(p, "object").length > 0 ||
                                   findDescendants(p, "pict").length > 0;

                if (!text && !hasSpecial) {
                    consecutiveEmpty++;
                    // Giữ tối đa 1 dòng trống, xóa nếu có 2 dòng trống liên tiếp trở lên
                    if (consecutiveEmpty > 1) {
                        if (p.parentNode) {
                            p.parentNode.removeChild(p);
                        }
                    }
                } else {
                    consecutiveEmpty = 0;
                }
            }

            // Xóa các đoạn văn rỗng ở cuối tài liệu (Trailing Empty Paragraphs)
            let guard = 50;
            while (guard-- > 0) {
                const children = Array.from(body.childNodes).filter(n => n.nodeType === 1);
                if (children.length === 0) break;
                const lastEl = children[children.length - 1];

                if (lastEl.localName === "p") {
                    const text = getParagraphText(lastEl).trim();
                    const hasSpecial = findDescendants(lastEl, "drawing").length > 0 ||
                                       findDescendants(lastEl, "object").length > 0 ||
                                       findDescendants(lastEl, "sectPr").length > 0;

                    if (!text && !hasSpecial) {
                        body.removeChild(lastEl);
                    } else {
                        break;
                    }
                } else {
                    break;
                }
            }
        }

        return {
            cleanedCount,
            convertedFontsCount
        };
    }

    // Export namespace
    window.TextCleaner = {
        TCVN3_TO_UNICODE,
        VNI_TO_UNICODE_PAIRS,
        isTcvn3Font,
        isVniFont,
        convertTcvn3ToUnicode,
        convertVniToUnicode,
        cleanVietnamesePunctuation,
        cleanDocumentFormatting
    };

})(typeof window !== 'undefined' ? window : global);
