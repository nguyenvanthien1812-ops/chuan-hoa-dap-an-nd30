const fs = require('fs');
const path = require('path');

console.log("=== KIỂM TRA TÍNH TOÀN VẸN DỰ ÁN NETLIFY ===");
const files = [
    "index.html",
    "netlify.toml",
    "README.md",
    "css/style.css",
    "js/app.js",
    "js/lib/jszip.min.js",
    "js/lib/file-saver.min.js",
    "js/core/xml_utils.js",
    "js/core/text_cleaner.js",
    "js/core/template_generator.js",
    "js/core/choice_compactor.js",
    "js/core/normalizer.js",
    "js/core/english_normalizer.js",
    "js/core/nd30_normalizer.js"
];

let allExist = true;
files.forEach(f => {
    const fullPath = path.resolve(__dirname, '..', f);
    const exists = fs.existsSync(fullPath);
    const size = exists ? fs.statSync(fullPath).size : 0;
    console.log(`- ${f.padEnd(30)}: ${exists ? 'OK (' + (size / 1024).toFixed(1) + ' KB)' : 'MISSING'}`);
    if (!exists) allExist = false;
});

if (allExist) {
    console.log("\n>>> TẤT CẢ FILE ĐÃ SẴN SÀNG 100% CHO NETLIFY & OFFLINE! <<<");
} else {
    console.error("\n>>> Có file bị thiếu! <<<");
    process.exit(1);
}
