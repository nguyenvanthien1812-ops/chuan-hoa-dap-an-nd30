/**
 * xml_utils.js - OpenXML DOM Manipulation Utilities for Word (.docx)
 * Pure Client-Side Document Manipulation Engine
 */

const NS = {
    w: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    m: "http://schemas.openxmlformats.org/officeDocument/2006/math",
    v: "urn:schemas-microsoft-com:vml",
    o: "urn:schemas-microsoft-com:office:office",
    wp: "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    a: "http://schemas.openxmlformats.org/drawingml/2006/main",
    pic: "http://schemas.openxmlformats.org/drawingml/2006/picture"
};

const domParser = new DOMParser();
const xmlSerializer = new XMLSerializer();

function parseXml(xmlString) {
    return domParser.parseFromString(xmlString, "application/xml");
}

function serializeXml(xmlDoc) {
    return xmlSerializer.serializeToString(xmlDoc);
}

function createWordElement(xmlDoc, tagName) {
    return xmlDoc.createElementNS(NS.w, "w:" + tagName);
}

function getChildByTagName(parent, localName) {
    if (!parent) return null;
    for (let i = 0; i < parent.childNodes.length; i++) {
        const node = parent.childNodes[i];
        if (node.nodeType === 1 && (node.localName === localName || node.nodeName === "w:" + localName || node.nodeName === localName)) {
            return node;
        }
    }
    return null;
}

function getChildrenByTagName(parent, localName) {
    const results = [];
    if (!parent) return results;
    for (let i = 0; i < parent.childNodes.length; i++) {
        const node = parent.childNodes[i];
        if (node.nodeType === 1 && (node.localName === localName || node.nodeName === "w:" + localName || node.nodeName === localName)) {
            results.push(node);
        }
    }
    return results;
}

function findDescendants(parent, localName) {
    const results = [];
    if (!parent) return results;
    const elements = parent.getElementsByTagName("*");
    for (let i = 0; i < elements.length; i++) {
        const el = elements[i];
        if (el.localName === localName || el.nodeName === "w:" + localName || el.nodeName === localName) {
            results.push(el);
        }
    }
    return results;
}

function getParagraphText(pNode) {
    if (!pNode) return "";
    let text = "";
    const nodes = pNode.getElementsByTagName("*");
    for (let i = 0; i < nodes.length; i++) {
        const el = nodes[i];
        if (el.localName === "t") {
            text += el.textContent || "";
        } else if (el.localName === "tab") {
            text += "\t";
        } else if (el.localName === "br" || el.localName === "cr") {
            text += "\n";
        }
    }
    return text;
}

function getOrCreatePPr(pNode, xmlDoc) {
    let pPr = getChildByTagName(pNode, "pPr");
    if (!pPr) {
        pPr = createWordElement(xmlDoc, "pPr");
        pNode.insertBefore(pPr, pNode.firstChild);
    }
    return pPr;
}

function getOrCreateRPr(rNode, xmlDoc) {
    let rPr = getChildByTagName(rNode, "rPr");
    if (!rPr) {
        rPr = createWordElement(xmlDoc, "rPr");
        rNode.insertBefore(rPr, rNode.firstChild);
    }
    return rPr;
}

function setRunFont(rNode, xmlDoc, fontName = "Times New Roman", fontSizePt = null) {
    const rPr = getOrCreateRPr(rNode, xmlDoc);
    
    let rFonts = getChildByTagName(rPr, "rFonts");
    if (!rFonts) {
        rFonts = createWordElement(xmlDoc, "rFonts");
        rPr.appendChild(rFonts);
    }
    rFonts.setAttributeNS(NS.w, "w:ascii", fontName);
    rFonts.setAttributeNS(NS.w, "w:hAnsi", fontName);
    rFonts.setAttributeNS(NS.w, "w:eastAsia", fontName);
    rFonts.setAttributeNS(NS.w, "w:cs", fontName);
    
    if (fontSizePt !== null) {
        const halfPt = Math.round(fontSizePt * 2);
        let sz = getChildByTagName(rPr, "sz");
        if (!sz) {
            sz = createWordElement(xmlDoc, "sz");
            rPr.appendChild(sz);
        }
        sz.setAttributeNS(NS.w, "w:val", halfPt.toString());
        
        let szCs = getChildByTagName(rPr, "szCs");
        if (!szCs) {
            szCs = createWordElement(xmlDoc, "szCs");
            rPr.appendChild(szCs);
        }
        szCs.setAttributeNS(NS.w, "w:val", halfPt.toString());
    }
}

function setParagraphAlignment(pNode, xmlDoc, alignVal = "left") {
    const pPr = getOrCreatePPr(pNode, xmlDoc);
    let jc = getChildByTagName(pPr, "jc");
    if (jc) pPr.removeChild(jc);
    
    jc = createWordElement(xmlDoc, "jc");
    jc.setAttributeNS(NS.w, "w:val", alignVal);
    pPr.appendChild(jc);
}

function setParagraphIndent(pNode, xmlDoc, leftDxa = 0, firstLineDxa = 0, rightDxa = 0) {
    const pPr = getOrCreatePPr(pNode, xmlDoc);
    let ind = getChildByTagName(pPr, "ind");
    if (ind) pPr.removeChild(ind);
    
    ind = createWordElement(xmlDoc, "ind");
    ind.setAttributeNS(NS.w, "w:left", leftDxa.toString());
    ind.setAttributeNS(NS.w, "w:firstLine", firstLineDxa.toString());
    ind.setAttributeNS(NS.w, "w:right", rightDxa.toString());
    pPr.appendChild(ind);
}

function setParagraphSpacing(pNode, xmlDoc, beforePt = 0, afterPt = 0, lineMultiplier = 1.15) {
    const pPr = getOrCreatePPr(pNode, xmlDoc);
    let spacing = getChildByTagName(pPr, "spacing");
    if (spacing) pPr.removeChild(spacing);
    
    spacing = createWordElement(xmlDoc, "spacing");
    spacing.setAttributeNS(NS.w, "w:before", Math.round(beforePt * 20).toString());
    spacing.setAttributeNS(NS.w, "w:after", Math.round(afterPt * 20).toString());
    spacing.setAttributeNS(NS.w, "w:line", Math.round(lineMultiplier * 240).toString());
    spacing.setAttributeNS(NS.w, "w:lineRule", "auto");
    pPr.appendChild(spacing);
}

function getAllParagraphsInDoc(bodyNode) {
    const paragraphs = [];
    function traverse(node) {
        if (!node) return;
        for (let i = 0; i < node.childNodes.length; i++) {
            const child = node.childNodes[i];
            if (child.nodeType === 1) {
                if (child.localName === "p") {
                    paragraphs.push(child);
                } else if (child.localName === "tbl") {
                    const cells = child.getElementsByTagName("*");
                    for (let j = 0; j < cells.length; j++) {
                        if (cells[j].localName === "tc") {
                            for (let k = 0; k < cells[j].childNodes.length; k++) {
                                if (cells[j].childNodes[k].nodeType === 1 && cells[j].childNodes[k].localName === "p") {
                                    paragraphs.push(cells[j].childNodes[k]);
                                }
                            }
                        }
                    }
                } else if (["sdt", "sdtContent", "customXml", "smartTag"].includes(child.localName)) {
                    traverse(child);
                }
            }
        }
    }
    traverse(bodyNode);
    return paragraphs;
}

// Export global namespace
window.XmlUtils = {
    NS,
    parseXml,
    serializeXml,
    createWordElement,
    getChildByTagName,
    getChildrenByTagName,
    findDescendants,
    getParagraphText,
    getOrCreatePPr,
    getOrCreateRPr,
    setRunFont,
    setParagraphAlignment,
    setParagraphIndent,
    setParagraphSpacing,
    getAllParagraphsInDoc
};
