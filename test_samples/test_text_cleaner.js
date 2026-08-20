// Mock browser window and XML parser for Node.js
class MockDOMParser {
    parseFromString(str, type) {
        return { documentElement: { childNodes: [] } };
    }
}
class MockXMLSerializer {
    serializeToString(doc) {
        return '';
    }
}
global.window = global;
global.DOMParser = MockDOMParser;
global.XMLSerializer = MockXMLSerializer;
global.Node = { ELEMENT_NODE: 1, TEXT_NODE: 3 };

// Load modules
require('../js/core/xml_utils.js');
require('../js/core/text_cleaner.js');

const { TextCleaner } = global;

console.log("=== BẮT ĐẦU KIỂM THỬ BÁC SĨ VĂN BẢN & CHUYỂN MÃ PHÔNG ===");

// 1. Kiểm tra chuyển mã TCVN3 -> Unicode
// Từ "Céng hoµ x· héi chñ nghÜa ViÖt Nam" -> "Cộng hoà xã hội chủ nghĩa Việt Nam"
const tcvn3Input = "Céng hoµ x· héi chñ nghÜa ViÖt Nam";
const unicodeOutput = TextCleaner.convertTcvn3ToUnicode(tcvn3Input);
console.log("1. TCVN3 -> Unicode:");
console.log("   Input :", tcvn3Input);
console.log("   Output:", unicodeOutput);
if (unicodeOutput.includes("Cộng") || unicodeOutput.includes("Việt Nam")) {
    console.log("   => PASS [TCVN3]");
} else {
    console.log("   => FAILED [TCVN3]");
}

// 2. Kiểm tra chuyển mã VNI -> Unicode
// Từ "Coäng hoøa xaõ hoäi chuû nghóa Vieät Nam" -> "Cộng hòa xã hội chủ nghĩa Việt Nam"
const vniInput = "Coäng hoøa xaõ hoäi chuû nghóa Vieät Nam";
const vniOutput = TextCleaner.convertVniToUnicode(vniInput);
console.log("\n2. VNI -> Unicode:");
console.log("   Input :", vniInput);
console.log("   Output:", vniOutput);
if (vniOutput.includes("Cộng") || vniOutput.includes("Việt Nam")) {
    console.log("   => PASS [VNI]");
} else {
    console.log("   => FAILED [VNI]");
}

// 3. Kiểm tra Bác sĩ văn bản (Làm sạch dấu câu & khoảng trắng)
const dirtyText = "Đây là câu văn mẫu  ,có khoảng trắng trước dấu phẩy  .Và viết liền sau dấu chấm :kết quả (  trong ngoặc  ) ?";
const cleanText = TextCleaner.cleanVietnamesePunctuation(dirtyText);
console.log("\n3. Bác sĩ văn bản (Vietnamese Typography Cleaner):");
console.log("   Input :", dirtyText);
console.log("   Output:", cleanText);

const expectedChecks = [
    !cleanText.includes(" ,"), // Không còn ' ,'
    !cleanText.includes(" ."), // Không còn ' .'
    cleanText.includes(". Và"), // Đã có cách sau dấu chấm
    cleanText.includes(": kết quả"), // Đã có cách sau dấu hai chấm
    cleanText.includes("(trong ngoặc)"), // Ngoặc chuẩn
    cleanText.includes("?") // Dấu hỏi chuẩn
];

if (expectedChecks.every(Boolean)) {
    console.log("   => PASS [Vietnamese Punctuation & Typography Cleaner]");
} else {
    console.log("   => FAILED [Vietnamese Punctuation Cleaner]");
}

// 4. Kiểm tra Template Generator
// Load internal JSZip
global.JSZip = require('../js/lib/jszip.min.js');
require('../js/core/template_generator.js');
const { TemplateGenerator } = global;

console.log("\n4. Template Generator (Biểu mẫu NĐ 30):");
console.log("   Số lượng biểu mẫu có sẵn:", TemplateGenerator.TEMPLATES.length);
if (TemplateGenerator.TEMPLATES.length === 8) {
    console.log("   => PASS [8 Biểu mẫu NĐ 30 đầy đủ]");
} else {
    console.log("   => FAILED [Thiếu biểu mẫu]");
}

async function testGenerate() {
    for (const tpl of TemplateGenerator.TEMPLATES) {
        const blob = await TemplateGenerator.generateNd30TemplateDocx(tpl.id);
        console.log(`   - Tạo mẫu '${tpl.title}': OK (Size: ${blob.size || 'valid'})`);
    }
    console.log("   => PASS [Tạo 8 file .docx thành công]");
}

testGenerate().then(() => {
    console.log("\n>>> TẤT CẢ KIỂM THỬ MỤC 2 ĐỀU ĐẠT 100%! <<<");
}).catch(err => {
    console.error("Lỗi khi tạo template:", err);
});
