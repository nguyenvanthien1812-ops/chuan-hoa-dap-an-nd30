/**
 * template_generator.js - Module tạo biểu mẫu văn bản hành chính chuẩn Nghị định 30/2020/NĐ-CP
 * Pure Client-Side Implementation using JSZip
 */

(function(window) {

    // Danh sách 8 biểu mẫu hành chính thông dụng
    const TEMPLATES = [
        {
            id: "quyet_dinh",
            title: "Quyết Định (Cá Biệt)",
            category: "Văn bản quy phạm & cá biệt",
            badge: "Phổ biến nhất",
            icon: "fa-gavel",
            desc: "Dùng để ban hành quy chế, kiện toàn nhân sự, phê duyệt kế hoạch hoặc khen thưởng/kỷ luật theo NĐ 30.",
            fileName: "Mau_Quyet_Dinh_ND30.docx"
        },
        {
            id: "to_trinh",
            title: "Tờ Trình Đề Xuất",
            category: "Văn bản đề xuất",
            badge: "Hành chính",
            icon: "fa-file-signature",
            desc: "Dùng để trình cấp có thẩm quyền phê duyệt chủ trương, dự toán kinh phí hoặc đề án công tác mới.",
            fileName: "Mau_To_Trinh_ND30.docx"
        },
        {
            id: "cong_van",
            title: "Công Văn Hành Chính",
            category: "Trao đổi công tác",
            badge: "Hàng ngày",
            icon: "fa-envelope-open-text",
            desc: "Dùng để trao đổi, hướng dẫn nghiệp vụ, đôn đốc hoặc phúc đáp công việc giữa các cơ quan, đơn vị.",
            fileName: "Mau_Cong_Van_ND30.docx"
        },
        {
            id: "thong_bao",
            title: "Thông Báo Kết Luận / Triển Khai",
            category: "Thông tin chỉ đạo",
            badge: "Thông dụng",
            icon: "fa-bullhorn",
            desc: "Thông báo kết luận cuộc họp, thông báo lịch làm việc, nghỉ lễ hoặc triển khai nhiệm vụ chuyên môn.",
            fileName: "Mau_Thong_Bao_ND30.docx"
        },
        {
            id: "bien_ban",
            title: "Biên Bản Cuộc Họp / Hội Nghị",
            category: "Ghi nhận sự việc",
            badge: "Quan trọng",
            icon: "fa-clipboard-list",
            desc: "Ghi chép đầy đủ tiến trình, các ý kiến phát biểu và kết luận cuối cùng của cuộc họp, hội nghị.",
            fileName: "Mau_Bien_Ban_Hop_ND30.docx"
        },
        {
            id: "ke_hoach",
            title: "Kế Hoạch Công Tác",
            category: "Kế hoạch định hướng",
            badge: "Chiến lược",
            icon: "fa-calendar-check",
            desc: "Xác định mục tiêu, nội dung nhiệm vụ, tiến độ thực hiện và phân công trách nhiệm chi tiết.",
            fileName: "Mau_Ke_Hoach_Cong_Tac_ND30.docx"
        },
        {
            id: "giay_moi",
            title: "Giấy Mời Họp / Hội Thảo",
            category: "Thư mời",
            badge: "Nghi lễ",
            icon: "fa-envelope",
            desc: "Mời đại biểu, cán bộ tham dự cuộc họp, hội nghị trực tiếp hoặc trực tuyến theo đúng thể thức NĐ 30.",
            fileName: "Mau_Giay_Moi_Hop_ND30.docx"
        },
        {
            id: "bao_cao",
            title: "Báo Cáo Tổng Kết / Sơ Kết",
            category: "Báo cáo công tác",
            badge: "Định kỳ",
            icon: "fa-chart-pie",
            desc: "Đánh giá kết quả công tác đã đạt được, tồn tại hạn chế và đề ra phương hướng, nhiệm vụ thời gian tới.",
            fileName: "Mau_Bao_Cao_ND30.docx"
        }
    ];

    // File cấu trúc khung OpenXML
    const CONTENT_TYPES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
  <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
</Types>`;

    const ROOT_RELS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

    const WORD_RELS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
</Relationships>`;

    const STYLES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="Times New Roman" w:cs="Times New Roman"/>
        <w:sz w:val="28"/>
        <w:szCs w:val="28"/>
        <w:lang w:val="vi-VN"/>
      </w:rPr>
    </w:rPrDefault>
    <w:pPrDefault>
      <w:pPr>
        <w:spacing w:line="288" w:lineRule="auto" w:before="0" w:after="120"/>
        <w:jc w:val="both"/>
      </w:pPr>
    </w:pPrDefault>
  </w:docDefaults>
</w:styles>`;

    const SETTINGS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:defaultTabStop w:val="720"/>
</w:settings>`;

    const HEADER_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:pPr>
      <w:jc w:val="center"/>
      <w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>
    </w:pPr>
    <w:r>
      <w:rPr>
        <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>
        <w:sz w:val="26"/>
      </w:rPr>
      <w:fldSimple w:instr="PAGE"/>
    </w:r>
  </w:p>
</w:hdr>`;

    // Hàm tạo khối Header Quốc Hiệu & Cơ Quan Ban Hành theo NĐ 30 (bằng bảng 2 cột không viền)
    function buildHeaderTableXml(agencyLine1, agencyLine2, codeNumber, locationDate) {
        return `
    <w:tbl>
      <w:tblPr>
        <w:tblW w:w="0" w:type="auto"/>
        <w:tblBorders>
          <w:top w:val="none"/><w:left w:val="none"/><w:bottom w:val="none"/><w:right w:val="none"/>
          <w:insideH w:val="none"/><w:insideV w:val="none"/>
        </w:tblBorders>
      </w:tblPr>
      <w:tblGrid>
        <w:gridCol w:w="4500"/>
        <w:gridCol w:w="4900"/>
      </w:tblGrid>
      <w:tr>
        <w:tc>
          <w:tcPr><w:tcW w:w="4500" w:type="dxa"/></w:tcPr>
          <w:p>
            <w:pPr><w:jc w:val="center"/><w:spacing w:before="0" w:after="40" w:line="240" w:lineRule="auto"/></w:pPr>
            <w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t>${agencyLine1}</w:t></w:r>
          </w:p>
          <w:p>
            <w:pPr><w:jc w:val="center"/><w:spacing w:before="0" w:after="40" w:line="240" w:lineRule="auto"/></w:pPr>
            <w:r><w:rPr><w:b/><w:sz w:val="24"/></w:rPr><w:t>${agencyLine2}</w:t></w:r>
          </w:p>
          <w:p>
            <w:pPr><w:jc w:val="center"/><w:spacing w:before="40" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>
            <w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t>Số: ${codeNumber}</w:t></w:r>
          </w:p>
        </w:tc>
        <w:tc>
          <w:tcPr><w:tcW w:w="4900" w:type="dxa"/></w:tcPr>
          <w:p>
            <w:pPr><w:jc w:val="center"/><w:spacing w:before="0" w:after="40" w:line="240" w:lineRule="auto"/></w:pPr>
            <w:r><w:rPr><w:b/><w:sz w:val="24"/></w:rPr><w:t>CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM</w:t></w:r>
          </w:p>
          <w:p>
            <w:pPr><w:jc w:val="center"/><w:spacing w:before="0" w:after="60" w:line="240" w:lineRule="auto"/></w:pPr>
            <w:r><w:rPr><w:b/><w:sz w:val="26"/><w:u w:val="single"/></w:rPr><w:t>Độc lập - Tự do - Hạnh phúc</w:t></w:r>
          </w:p>
          <w:p>
            <w:pPr><w:jc w:val="right"/><w:spacing w:before="40" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>
            <w:r><w:rPr><w:i/><w:sz w:val="26"/></w:rPr><w:t>${locationDate}</w:t></w:r>
          </w:p>
        </w:tc>
      </w:tr>
    </w:tbl>`;
    }

    // Hàm tạo khối Chữ Ký & Nơi Nhận theo NĐ 30 (bằng bảng 2 cột không viền)
    function buildFooterTableXml(recipientList, signerTitle, signerName) {
        let recipientParas = recipientList.map((r, idx) => `
          <w:p>
            <w:pPr><w:jc w:val="left"/><w:spacing w:before="0" w:after="20" w:line="240" w:lineRule="auto"/></w:pPr>
            <w:r><w:rPr><w:i/><w:sz w:val="22"/></w:rPr><w:t>${r}</w:t></w:r>
          </w:p>`).join("");

        return `
    <w:tbl>
      <w:tblPr>
        <w:tblW w:w="0" w:type="auto"/>
        <w:tblBorders>
          <w:top w:val="none"/><w:left w:val="none"/><w:bottom w:val="none"/><w:right w:val="none"/>
          <w:insideH w:val="none"/><w:insideV w:val="none"/>
        </w:tblBorders>
      </w:tblPr>
      <w:tblGrid>
        <w:gridCol w:w="4500"/>
        <w:gridCol w:w="4900"/>
      </w:tblGrid>
      <w:tr>
        <w:tc>
          <w:tcPr><w:tcW w:w="4500" w:type="dxa"/></w:tcPr>
          <w:p>
            <w:pPr><w:jc w:val="left"/><w:spacing w:before="120" w:after="40" w:line="240" w:lineRule="auto"/></w:pPr>
            <w:r><w:rPr><w:b/><w:i/><w:sz w:val="24"/></w:rPr><w:t>Nơi nhận:</w:t></w:r>
          </w:p>
          ${recipientParas}
        </w:tc>
        <w:tc>
          <w:tcPr><w:tcW w:w="4900" w:type="dxa"/></w:tcPr>
          <w:p>
            <w:pPr><w:jc w:val="center"/><w:spacing w:before="120" w:after="40" w:line="240" w:lineRule="auto"/></w:pPr>
            <w:r><w:rPr><w:b/><w:sz w:val="26"/></w:rPr><w:t>${signerTitle.toUpperCase()}</w:t></w:r>
          </w:p>
          <w:p>
            <w:pPr><w:jc w:val="center"/><w:spacing w:before="0" w:after="40" w:line="240" w:lineRule="auto"/></w:pPr>
            <w:r><w:rPr><w:i/><w:sz w:val="24"/></w:rPr><w:t>(Chữ ký, dấu/chữ ký số)</w:t></w:r>
          </w:p>
          <w:p><w:pPr><w:spacing w:before="1000" w:after="0"/></w:pPr></w:p>
          <w:p>
            <w:pPr><w:jc w:val="center"/><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>
            <w:r><w:rPr><w:b/><w:sz w:val="28"/></w:rPr><w:t>${signerName}</w:t></w:r>
          </w:p>
        </w:tc>
      </w:tr>
    </w:tbl>`;
    }

    // Phần cấu hình khổ giấy A4 và lề trang chuẩn NĐ 30
    const SECT_PR_XML = `
    <w:sectPr>
      <w:headerReference w:type="default" r:id="rId3"/>
      <w:titlePg/>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="1134" w:right="850" w:bottom="1134" w:left="1701" w:header="720" w:footer="720" w:gutter="0"/>
    </w:sectPr>`;

    // =========================================================================
    // NỘI DUNG TỪNG BIỂU MẪU CỤ THỂ
    // =========================================================================
    function getDocumentBodyXml(templateId) {
        switch (templateId) {
            case "quyet_dinh":
                return `
    ${buildHeaderTableXml("ỦY BAN NHÂN DÂN / CƠ QUAN", "TÊN ĐƠN VỊ BAN HÀNH", ".../QĐ-UBND", "Hà Nội, ngày ... tháng ... năm 2026")}
    
    <w:p>
      <w:pPr><w:jc w:val="center"/><w:spacing w:before="240" w:after="80" w:line="288" w:lineRule="auto"/></w:pPr>
      <w:r><w:rPr><w:b/><w:sz w:val="30"/></w:rPr><w:t>QUYẾT ĐỊNH</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:jc w:val="center"/><w:spacing w:before="0" w:after="240" w:line="288" w:lineRule="auto"/></w:pPr>
      <w:r><w:rPr><w:b/><w:sz w:val="28"/></w:rPr><w:t>Về việc ban hành quy chế làm việc của cơ quan</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:jc w:val="center"/><w:spacing w:before="0" w:after="200" w:line="288" w:lineRule="auto"/></w:pPr>
      <w:r><w:rPr><w:b/><w:sz w:val="28"/></w:rPr><w:t>THỦ TRƯỞNG CƠ QUAN / ĐƠN VỊ</w:t></w:r>
    </w:p>

    <w:p>
      <w:pPr><w:ind w:firstLine="567"/><w:spacing w:before="0" w:after="120" w:line="288" w:lineRule="auto"/></w:pPr>
      <w:r><w:rPr><w:i/></w:rPr><w:t>Căn cứ Luật Tổ chức chính quyền địa phương ngày 19 tháng 6 năm 2015;</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:ind w:firstLine="567"/><w:spacing w:before="0" w:after="120" w:line="288" w:lineRule="auto"/></w:pPr>
      <w:r><w:rPr><w:i/></w:rPr><w:t>Căn cứ Nghị định số 30/2020/NĐ-CP ngày 05 tháng 3 năm 2020 của Chính phủ về công tác văn thư;</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:ind w:firstLine="567"/><w:spacing w:before="0" w:after="200" w:line="288" w:lineRule="auto"/></w:pPr>
      <w:r><w:rPr><w:i/></w:rPr><w:t>Theo đề nghị của Trưởng phòng Tổ chức - Hành chính.</w:t></w:r>
    </w:p>

    <w:p>
      <w:pPr><w:jc w:val="center"/><w:spacing w:before="120" w:after="160" w:line="288" w:lineRule="auto"/></w:pPr>
      <w:r><w:rPr><w:b/><w:sz w:val="28"/></w:rPr><w:t>QUYẾT ĐỊNH:</w:t></w:r>
    </w:p>

    <w:p>
      <w:pPr><w:ind w:firstLine="567"/><w:spacing w:before="0" w:after="120" w:line="288" w:lineRule="auto"/></w:pPr>
      <w:r><w:rPr><w:b/></w:rPr><w:t>Điều 1. </w:t></w:r>
      <w:r><w:t>Ban hành kèm theo Quyết định này Quy chế làm việc mới áp dụng cho toàn thể cán bộ, công chức, viên chức và người lao động trong cơ quan.</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:ind w:firstLine="567"/><w:spacing w:before="0" w:after="120" w:line="288" w:lineRule="auto"/></w:pPr>
      <w:r><w:rPr><w:b/></w:rPr><w:t>Điều 2. </w:t></w:r>
      <w:r><w:t>Quyết định này có hiệu lực thi hành kể từ ngày ký và thay thế các văn bản quy định trước đây trái với Quyết định này.</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:ind w:firstLine="567"/><w:spacing w:before="0" w:after="240" w:line="288" w:lineRule="auto"/></w:pPr>
      <w:r><w:rPr><w:b/></w:rPr><w:t>Điều 3. </w:t></w:r>
      <w:r><w:t>Chánh Văn phòng, Trưởng các phòng, ban chuyên môn và các cá nhân có liên quan chịu trách nhiệm thi hành Quyết định này./.</w:t></w:r>
    </w:p>

    ${buildFooterTableXml(["- Như Điều 3;", "- Lãnh đạo cơ quan;", "- Lưu: VT, TCHC."], "GIÁM ĐỐC / CHỦ TỊCH", "Nguyễn Văn A")}`;

            case "to_trinh":
                return `
    ${buildHeaderTableXml("TÊN CƠ QUAN CHỦ QUẢN", "TÊN ĐƠN VỊ TRÌNH", ".../TTr-...", "Hà Nội, ngày ... tháng ... năm 2026")}

    <w:p>
      <w:pPr><w:jc w:val="center"/><w:spacing w:before="240" w:after="80" w:line="288" w:lineRule="auto"/></w:pPr>
      <w:r><w:rPr><w:b/><w:sz w:val="30"/></w:rPr><w:t>TỜ TRÌNH</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:jc w:val="center"/><w:spacing w:before="0" w:after="200" w:line="288" w:lineRule="auto"/></w:pPr>
      <w:r><w:rPr><w:b/><w:sz w:val="28"/></w:rPr><w:t>Về việc đề xuất phê duyệt kế hoạch tổ chức tập huấn chuyên môn</w:t></w:r>
    </w:p>

    <w:p>
      <w:pPr><w:jc w:val="center"/><w:spacing w:before="0" w:after="200" w:line="288" w:lineRule="auto"/></w:pPr>
      <w:r><w:rPr><w:b/></w:rPr><w:t>Kính gửi: </w:t></w:r>
      <w:r><w:t>Lãnh đạo Cơ quan / Cấp có thẩm quyền</w:t></w:r>
    </w:p>

    <w:p>
      <w:pPr><w:ind w:firstLine="567"/><w:spacing w:before="0" w:after="120" w:line="288" w:lineRule="auto"/></w:pPr>
      <w:r><w:t>Thực hiện nhiệm vụ năm công tác 2026, nhằm nâng cao trình độ chuyên môn nghiệp vụ và kỹ năng ứng dụng công nghệ thông tin cho cán bộ, giáo viên và nhân viên;</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:ind w:firstLine="567"/><w:spacing w:before="0" w:after="120" w:line="288" w:lineRule="auto"/></w:pPr>
      <w:r><w:t>Đơn vị kính trình Lãnh đạo cơ quan xem xét, phê duyệt Kế hoạch tổ chức đợt tập huấn với các nội dung trọng tâm như sau:</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:ind w:firstLine="567"/><w:spacing w:before="0" w:after="120" w:line="288" w:lineRule="auto"/></w:pPr>
      <w:r><w:rPr><w:b/></w:rPr><w:t>1. Mục đích, yêu cầu: </w:t></w:r>
      <w:r><w:t>Chuẩn hóa quy trình làm việc, cập nhật các quy định mới theo Nghị định 30/2020/NĐ-CP.</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:ind w:firstLine="567"/><w:spacing w:before="0" w:after="120" w:line="288" w:lineRule="auto"/></w:pPr>
      <w:r><w:rPr><w:b/></w:rPr><w:t>2. Thời gian và địa điểm: </w:t></w:r>
      <w:r><w:t>Dự kiến tổ chức trong 02 ngày, tại Hội trường lớn cơ quan.</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:ind w:firstLine="567"/><w:spacing w:before="0" w:after="120" w:line="288" w:lineRule="auto"/></w:pPr>
      <w:r><w:rPr><w:b/></w:rPr><w:t>3. Dự toán kinh phí: </w:t></w:r>
      <w:r><w:t>Sử dụng nguồn kinh phí đào tạo, bồi dưỡng đã được phân bổ trong năm 2026.</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:ind w:firstLine="567"/><w:spacing w:before="0" w:after="240" w:line="288" w:lineRule="auto"/></w:pPr>
      <w:r><w:t>Kính trình Lãnh đạo cơ quan xem xét, quyết định./.</w:t></w:r>
    </w:p>

    ${buildFooterTableXml(["- Như kính gửi;", "- Lưu: VT, ĐV."], "TRƯỞNG PHÒNG / ĐƠN VỊ", "Trần Thị B")}`;

            case "cong_van":
                return `
    ${buildHeaderTableXml("TÊN CƠ QUAN CHỦ QUẢN", "TÊN CƠ QUAN PHÁT HÀNH", ".../CV-...", "Hà Nội, ngày ... tháng ... năm 2026")}

    <w:p>
      <w:pPr><w:jc w:val="left"/><w:spacing w:before="80" w:after="200" w:line="240" w:lineRule="auto"/></w:pPr>
      <w:r><w:rPr><w:i/><w:sz w:val="24"/></w:rPr><w:t>V/v hướng dẫn công tác chuẩn hóa đề thi và văn bản hành chính</w:t></w:r>
    </w:p>

    <w:p>
      <w:pPr><w:jc w:val="center"/><w:spacing w:before="0" w:after="200" w:line="288" w:lineRule="auto"/></w:pPr>
      <w:r><w:rPr><w:b/></w:rPr><w:t>Kính gửi: </w:t></w:r>
      <w:r><w:t>Các phòng ban, đơn vị trực thuộc</w:t></w:r>
    </w:p>

    <w:p>
      <w:pPr><w:ind w:firstLine="567"/><w:spacing w:before="0" w:after="120" w:line="288" w:lineRule="auto"/></w:pPr>
      <w:r><w:t>Nhằm tăng cường hiệu lực, hiệu quả trong công tác văn thư và nâng cao chất lượng biên soạn đề thi trắc nghiệm, Cơ quan hướng dẫn các đơn vị thực hiện một số nội dung sau:</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:ind w:firstLine="567"/><w:spacing w:before="0" w:after="120" w:line="288" w:lineRule="auto"/></w:pPr>
      <w:r><w:rPr><w:b/></w:rPr><w:t>1. </w:t></w:r>
      <w:r><w:t>Thực hiện nghiêm túc quy định về thể thức và kỹ thuật trình bày văn bản theo Nghị định số 30/2020/NĐ-CP.</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:ind w:firstLine="567"/><w:spacing w:before="0" w:after="120" w:line="288" w:lineRule="auto"/></w:pPr>
      <w:r><w:rPr><w:b/></w:rPr><w:t>2. </w:t></w:r>
      <w:r><w:t>Ứng dụng phần mềm tự động để bóc tách đáp án, dồn dòng phương án và chuẩn hóa đề mục thống nhất.</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:ind w:firstLine="567"/><w:spacing w:before="0" w:after="240" w:line="288" w:lineRule="auto"/></w:pPr>
      <w:r><w:t>Đề nghị Thủ trưởng các đơn vị nghiêm túc triển khai thực hiện./.</w:t></w:r>
    </w:p>

    ${buildFooterTableXml(["- Như trên;", "- Lãnh đạo cơ quan;", "- Lưu: VT."], "THỦ TRƯỞNG ĐƠN VỊ", "Lê Văn C")}`;

            default:
                // Mẫu văn bản chung (Thông báo, Biên bản, Kế hoạch, Giấy mời, Báo cáo)
                const titles = {
                    thong_bao: "THÔNG BÁO",
                    bien_ban: "BIÊN BẢN CUỘC HỌP",
                    ke_hoach: "KẾ HOẠCH CÔNG TÁC",
                    giay_moi: "GIẤY MỜI HỌP",
                    bao_cao: "BÁO CÁO TỔNG KẾT"
                };
                const mainTitle = titles[templateId] || "VĂN BẢN HÀNH CHÍNH";

                return `
    ${buildHeaderTableXml("TÊN CƠ QUAN CẤP TRÊN", "TÊN CƠ QUAN BAN HÀNH", ".../...", "Hà Nội, ngày ... tháng ... năm 2026")}

    <w:p>
      <w:pPr><w:jc w:val="center"/><w:spacing w:before="240" w:after="80" w:line="288" w:lineRule="auto"/></w:pPr>
      <w:r><w:rPr><w:b/><w:sz w:val="30"/></w:rPr><w:t>${mainTitle}</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:jc w:val="center"/><w:spacing w:before="0" w:after="200" w:line="288" w:lineRule="auto"/></w:pPr>
      <w:r><w:rPr><w:b/><w:sz w:val="28"/></w:rPr><w:t>Về việc triển khai nhiệm vụ trọng tâm công tác</w:t></w:r>
    </w:p>

    <w:p>
      <w:pPr><w:ind w:firstLine="567"/><w:spacing w:before="0" w:after="120" w:line="288" w:lineRule="auto"/></w:pPr>
      <w:r><w:t>Căn cứ chương trình kế hoạch công tác và tình hình thực tế triển khai nhiệm vụ chuyên môn;</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:ind w:firstLine="567"/><w:spacing w:before="0" w:after="120" w:line="288" w:lineRule="auto"/></w:pPr>
      <w:r><w:rPr><w:b/></w:rPr><w:t>I. MỤC ĐÍCH VÀ YÊU CẦU</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:ind w:firstLine="567"/><w:spacing w:before="0" w:after="120" w:line="288" w:lineRule="auto"/></w:pPr>
      <w:r><w:t>1. Nâng cao hiệu quả phối hợp và trách nhiệm của từng bộ phận chuyên môn.</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:ind w:firstLine="567"/><w:spacing w:before="0" w:after="120" w:line="288" w:lineRule="auto"/></w:pPr>
      <w:r><w:rPr><w:b/></w:rPr><w:t>II. NỘI DUNG THỰC HIỆN</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:ind w:firstLine="567"/><w:spacing w:before="0" w:after="120" w:line="288" w:lineRule="auto"/></w:pPr>
      <w:r><w:t>Triển khai các giải pháp đổi mới sáng tạo, đảm bảo tiến độ và chất lượng đề ra theo quy chuẩn.</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:ind w:firstLine="567"/><w:spacing w:before="0" w:after="240" w:line="288" w:lineRule="auto"/></w:pPr>
      <w:r><w:t>Yêu cầu các đơn vị có liên quan chủ động phối hợp tổ chức thực hiện nghiêm túc./.</w:t></w:r>
    </w:p>

    ${buildFooterTableXml(["- Như trên;", "- Lưu: VT."], "THỦ TRƯỞNG CƠ QUAN", "Phạm Văn D")}`;
        }
    }

    // =========================================================================
    // HÀM TẠO FILE DOCX TỔNG HỢP VÀ TẢI VỀ
    // =========================================================================
    async function generateNd30TemplateDocx(templateId) {
        const ZipConstructor = window.JSZip || (typeof JSZip !== 'undefined' ? JSZip : null);
        if (!ZipConstructor) {
            throw new Error("Không tìm thấy thư viện JSZip!");
        }
        const zip = new ZipConstructor();

        const bodyXml = getDocumentBodyXml(templateId);
        const fullDocXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    ${bodyXml}
    ${SECT_PR_XML}
  </w:body>
</w:document>`;

        zip.file("[Content_Types].xml", CONTENT_TYPES_XML);
        zip.file("_rels/.rels", ROOT_RELS_XML);
        zip.file("word/_rels/document.xml.rels", WORD_RELS_XML);
        zip.file("word/styles.xml", STYLES_XML);
        zip.file("word/settings.xml", SETTINGS_XML);
        zip.file("word/header1.xml", HEADER_XML);
        zip.file("word/document.xml", fullDocXml);

        const outBlob = await zip.generateAsync({
            type: "blob",
            mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        });

        return outBlob;
    }

    // Export global namespace
    window.TemplateGenerator = {
        TEMPLATES,
        generateNd30TemplateDocx
    };

})(typeof window !== 'undefined' ? window : global);
