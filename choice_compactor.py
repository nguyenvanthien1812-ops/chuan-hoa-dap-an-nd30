"""Module dồn dòng / gộp các phương án trắc nghiệm trong file Word (.docx).

Hỗ trợ các chế độ:
- 'auto': Tự động phân tích độ dài từng phương án để xếp 4, 2 hoặc 1 phương án / dòng (Auto 4-2-1).
- '4_per_line': Gộp 4 phương án trên 1 dòng.
- '2_per_line': Gộp 2 phương án trên 1 dòng (2 dòng: A-B và C-D).
- '1_per_line' / 'split': Giữ nguyên mỗi phương án 1 dòng riêng.

Bảo toàn 100% công thức toán học MathType / OMML, hình ảnh, in đậm/nghiêng, màu sắc và dấu hoa thị (*).
Căn chỉnh Tab Stop chuẩn xác theo lề trang, đo chính xác kích thước MathType OLE, loại bỏ triệt để tab trùng lặp.
"""
import os
import re
import copy
import logging
from typing import List, Tuple, Optional, Union
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph
from docx.table import Table, _Cell

logger = logging.getLogger(__name__)

# Regex nhận dạng nhãn phương án trắc nghiệm
# Upper: A., B., C., D., A*, A), A:, A-, (A)
CHOICE_UPPER_RE = re.compile(
    r'^\s*(?:\()?(\*?\s*[A-H]\s*\*?)(?:\s*[.:)/\-]\s*|\s*[\).]\s*\*?)',
    re.IGNORECASE
)

# Lower: a), b), c), d), a., a*), (a)
CHOICE_LOWER_RE = re.compile(
    r'^\s*(?:\()?(\*?\s*[a-h]\s*\*?)(?:\s*[\).]\s*\*?|\s*[.:)/\-]\s*)'
)

# Vị trí Tab Stop (đơn vị dxa: 1 cm = 567 dxa, 1 inch = 1440 dxa)
# Chuẩn cho khổ giấy A4 (vùng văn bản ~16.5cm):
# 4 cột: Tab 1 tại 3.70cm, Tab 2 tại 7.40cm, Tab 3 tại 11.10cm -> Cột 4 có hơn 5.4cm không bị tràn lề
TAB_STOPS_4_COLS = [2100, 4200, 6300]       # ~3.70cm, ~7.40cm, ~11.10cm
# 2 cột: Tab 1 tại 7.94cm -> Cột 1 có 7.9cm, Cột 2 có hơn 8.5cm
TAB_STOPS_2_COLS = [4500]                   # ~7.94cm
# 3 cột: Tab 1 tại 5.11cm, Tab 2 tại 10.23cm
TAB_STOPS_3_COLS = [2900, 5800]             # ~5.11cm, ~10.23cm


def _extract_choice_label_info(p: Paragraph) -> Tuple[Optional[str], bool, int]:
    """
    Kiểm tra xem đoạn văn có phải là phương án trắc nghiệm hay không.
    
    Returns:
        (label_char, is_lower, letter_index)
        Ví dụ: ('A', False, 0) hoặc ('b', True, 1). Nếu không phải trả về (None, False, -1).
    """
    text = p.text.strip()
    if not text:
        return None, False, -1

    # Kiểm tra nhãn chữ hoa (A, B, C, D...)
    m_up = CHOICE_UPPER_RE.match(text)
    if m_up:
        raw = m_up.group(1).replace('*', '').strip().upper()
        if len(raw) == 1 and 'A' <= raw <= 'H':
            return raw, False, ord(raw) - ord('A')

    # Kiểm tra nhãn chữ thường (a, b, c, d...)
    m_low = CHOICE_LOWER_RE.match(text)
    if m_low:
        raw = m_low.group(1).replace('*', '').strip().lower()
        if len(raw) == 1 and 'a' <= raw <= 'h':
            return raw, True, ord(raw) - ord('a')

    return None, False, -1


def _has_heavy_content(p: Paragraph) -> bool:
    """Kiểm tra xem paragraph có chứa hình ảnh lớn hoặc bảng biểu không."""
    p_elem = p._element
    # Kiểm tra ảnh (w:drawing, w:pict) - trừ MathType OLE
    drawings = p_elem.findall('.//' + qn('w:drawing'))
    picts = p_elem.findall('.//' + qn('w:pict'))
    
    # Nếu có w:drawing (ảnh chèn dạng hiện đại)
    if drawings:
        return True
        
    # Với w:pict, kiểm tra xem có phải ảnh chụp/sơ đồ không (không phải MathType OLE)
    for pict in picts:
        # Nếu pict chứa OLEObject MathType thì không coi là heavy image
        if pict.findall('.//' + qn('o:OLEObject')):
            continue
        return True
        
    return False


def _is_paragraph_empty(p: Paragraph) -> bool:
    """Kiểm tra xem paragraph có hoàn toàn rỗng hay không (không text, không ảnh, không math)."""
    text = p.text.strip()
    if text:
        return False
    p_elem = p._element
    if p_elem.findall('.//' + qn('w:drawing')):
        return False
    if p_elem.findall('.//' + qn('w:pict')):
        return False
    if p_elem.findall('.//' + qn('m:oMath')):
        return False
    if p_elem.findall('.//' + qn('w:object')):
        return False
    return True


def _split_paragraphs_at_br(doc: Document):
    """Tách tất cả các ngắt dòng mềm <w:br/> thành các paragraph <w:p> độc lập."""
    for p in list(doc.paragraphs):
        p_elem = p._element
        runs = p_elem.findall('.//' + qn('w:r'))
        has_br = any(len(r.findall('.//' + qn('w:br'))) > 0 for r in runs)
        if not has_br:
            continue

        parent = p_elem.getparent()
        if parent is None:
            continue

        pPr = p_elem.find(qn('w:pPr'))
        new_paras = []
        curr_p = OxmlElement('w:p')
        if pPr is not None:
            curr_p.append(copy.deepcopy(pPr))
        new_paras.append(curr_p)

        for child in list(p_elem):
            if child.tag == qn('w:pPr'):
                continue
            if child.tag == qn('w:r'):
                brs = child.findall(qn('w:br'))
                if not brs:
                    curr_p.append(child)
                else:
                    rPr = child.find(qn('w:rPr'))
                    curr_r = OxmlElement('w:r')
                    if rPr is not None:
                        curr_r.append(copy.deepcopy(rPr))
                    curr_p.append(curr_r)

                    for rc in list(child):
                        if rc.tag == qn('w:rPr'):
                            continue
                        if rc.tag == qn('w:br'):
                            curr_p = OxmlElement('w:p')
                            if pPr is not None:
                                curr_p.append(copy.deepcopy(pPr))
                            new_paras.append(curr_p)

                            curr_r = OxmlElement('w:r')
                            if rPr is not None:
                                curr_r.append(copy.deepcopy(rPr))
                            curr_p.append(curr_r)
                        else:
                            curr_r.append(rc)
            else:
                curr_p.append(child)

        for np in new_paras:
            p_elem.addprevious(np)
        parent.remove(p_elem)


def _convert_choice_tables_to_paragraphs(doc: Document):
    """Chuyển đổi bảng chứa phương án trắc nghiệm thành các đoạn văn riêng biệt."""
    CHOICE_LABEL_RE = re.compile(r'^\s*(?:\()?([A-Ha-h])\s*[.:)/\-]\s*')
    tables_to_process = []
    for tbl in list(doc.tables):
        choice_cells_count = 0
        total_cells = 0
        for row in tbl.rows:
            for cell in row.cells:
                total_cells += 1
                if CHOICE_LABEL_RE.search(cell.text.strip()):
                    choice_cells_count += 1
        if choice_cells_count >= 2 and choice_cells_count >= total_cells * 0.5:
            tables_to_process.append(tbl)

    for tbl in tables_to_process:
        tbl_elem = tbl._element
        parent = tbl_elem.getparent()
        if parent is None:
            continue
        for row in tbl.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    tbl_elem.addprevious(p._element)
        parent.remove(tbl_elem)


def _get_paragraph_estimated_length(p: Paragraph) -> int:
    """
    Ước lượng độ dài trực quan của paragraph theo số ký tự tương đương.
    Đo chính xác kích thước thực tế của MathType OLE object, OMML math và hình ảnh.
    """
    p_elem = p._element
    text_len = len(p.text.strip())
    
    # 1. Đo kích thước các đối tượng MathType OLE (w:object)
    ole_extra_len = 0
    for obj in p_elem.findall('.//' + qn('w:object')):
        # Thử lấy dxaOrig (đơn vị dxa: 1 ký tự font 11/12 ~ 95 dxa, font 14 ~ 110 dxa)
        dxa_orig = obj.get(qn('w:dxaOrig'))
        if dxa_orig and dxa_orig.isdigit():
            ole_extra_len += max(5, int(dxa_orig) // 95)
        else:
            # Thử đọc width từ style trong v:shape
            shape = obj.find('.//' + qn('v:shape'))
            if shape is not None:
                style_str = shape.get('style') or ""
                m_w = re.search(r'width\s*:\s*([\d.]+)\s*pt', style_str, re.IGNORECASE)
                if m_w:
                    pt_val = float(m_w.group(1))
                    # 1 pt = 20 dxa -> ký tự = pt_val * 20 / 95
                    ole_extra_len += max(5, int(pt_val * 20 // 95))
                else:
                    ole_extra_len += 20
            else:
                ole_extra_len += 20

    # 2. Đếm số lượng công thức OMML (m:oMath)
    omml_extra_len = 0
    for m in p_elem.findall('.//' + qn('m:oMath')):
        m_txt = m.xpath('string(.)')
        omml_extra_len += max(8, len(m_txt) + 4)

    # 3. Đếm số ảnh thông thường
    drawings = len(p_elem.findall('.//' + qn('w:drawing')))
    img_extra_len = drawings * 40

    return text_len + ole_extra_len + omml_extra_len + img_extra_len


def _clean_trailing_whitespace_and_tabs(p: Paragraph):
    """
    Xóa triệt để TẤT CẢ các ký tự tab (<w:tab/>), ngắt dòng (<w:br/>, <w:cr/>),
    và khoảng trắng thừa ở cuối đoạn văn.
    """
    p_elem = p._element
    while True:
        children = list(p_elem)
        if not children:
            break
        last_c = children[-1]
        if last_c.tag == qn('w:pPr'):
            break
        if last_c.tag == qn('w:r'):
            while True:
                r_children = list(last_c)
                if not r_children:
                    break
                last_r_c = r_children[-1]
                if last_r_c.tag in (qn('w:tab'), qn('w:br'), qn('w:cr'), qn('w:noBreakHyphen')):
                    last_c.remove(last_r_c)
                elif last_r_c.tag == qn('w:t'):
                    if last_r_c.text:
                        last_r_c.text = last_r_c.text.rstrip(' \t\r\n\xa0')
                    if not last_r_c.text:
                        last_c.remove(last_r_c)
                    else:
                        break
                elif last_r_c.tag == qn('w:rPr'):
                    break
                else:
                    break
            meaningful = [c for c in list(last_c) if c.tag != qn('w:rPr')]
            if not meaningful:
                p_elem.remove(last_c)
                continue
            else:
                break
        else:
            break


def _clean_leading_whitespace_and_tabs(p: Paragraph):
    """
    Xóa triệt để TẤT CẢ các ký tự tab (<w:tab/>), ngắt dòng, khoảng trắng ở ĐẦU đoạn văn
    trước nhãn phương án.
    """
    p_elem = p._element
    while True:
        children = list(p_elem)
        if not children:
            break
        first_c = None
        for c in children:
            if c.tag != qn('w:pPr'):
                first_c = c
                break
        if first_c is None:
            break
        if first_c.tag == qn('w:r'):
            while True:
                r_children = list(first_c)
                if not r_children:
                    break
                first_r_c = None
                for rc in r_children:
                    if rc.tag != qn('w:rPr'):
                        first_r_c = rc
                        break
                if first_r_c is None:
                    break
                if first_r_c.tag in (qn('w:tab'), qn('w:br'), qn('w:cr'), qn('w:noBreakHyphen')):
                    first_c.remove(first_r_c)
                elif first_r_c.tag == qn('w:t'):
                    if first_r_c.text:
                        first_r_c.text = first_r_c.text.lstrip(' \t\r\n\xa0')
                    if not first_r_c.text:
                        first_c.remove(first_r_c)
                    else:
                        break
                else:
                    break
            meaningful = [c for c in list(first_c) if c.tag != qn('w:rPr')]
            if not meaningful:
                p_elem.remove(first_c)
                continue
            else:
                break
        else:
            break


def _collapse_consecutive_tabs(p: Paragraph):
    """
    Đảm bảo giữa các cụm phương án chỉ có DUY NHẤT 1 tab, không bị nhân đôi (\\t\\t)
    khi chuẩn hóa nhiều lần hoặc từ file gốc đã có tab.
    """
    p_elem = p._element
    prev_was_tab = False
    for r in list(p_elem.findall(qn('w:r'))):
        for child in list(r):
            if child.tag == qn('w:tab'):
                if prev_was_tab:
                    r.remove(child)
                else:
                    prev_was_tab = True
            elif child.tag == qn('w:t'):
                if child.text and child.text.strip():
                    prev_was_tab = False
            elif child.tag not in (qn('w:rPr'),):
                prev_was_tab = False
        meaningful = [c for c in list(r) if c.tag != qn('w:rPr')]
        if not meaningful:
            p_elem.remove(r)


def _reset_paragraph_formatting_for_compact(p: Paragraph, indent_dxa: int = 0):
    """
    Xóa bỏ lề thụt dòng (firstLine, left, right indent) và ép căn trái (left align)
    để các Tab Stop hoạt động chính xác từ mép lề trái, tránh bị co chữ / nhảy dòng.
    Nếu indent_dxa > 0 thì thiết lập lề thụt đầu dòng tương ứng.
    """
    p_elem = p._element
    pPr = p_elem.find(qn('w:pPr'))
    if pPr is None:
        pPr = OxmlElement('w:pPr')
        p_elem.insert(0, pPr)

    # 1. Reset thuộc tính thụt dòng (w:ind)
    ind = pPr.find(qn('w:ind'))
    if ind is not None:
        pPr.remove(ind)
    new_ind = OxmlElement('w:ind')
    new_ind.set(qn('w:left'), str(indent_dxa or 0))
    new_ind.set(qn('w:firstLine'), '0')
    new_ind.set(qn('w:right'), '0')
    pPr.append(new_ind)

    # 2. Ép căn lề trái (w:jc w:val="left"), tuyệt đối không để justify
    jc = pPr.find(qn('w:jc'))
    if jc is not None:
        pPr.remove(jc)
    new_jc = OxmlElement('w:jc')
    new_jc.set(qn('w:val'), 'left')
    pPr.append(new_jc)


def _set_paragraph_tabs(p: Paragraph, tab_positions: List[int]):
    """Cài đặt danh sách Tab Stop vào thuộc tính w:pPr của đoạn văn."""
    p_elem = p._element
    pPr = p_elem.find(qn('w:pPr'))
    if pPr is None:
        pPr = OxmlElement('w:pPr')
        p_elem.insert(0, pPr)

    # Xóa w:tabs cũ nếu có
    old_tabs = pPr.find(qn('w:tabs'))
    if old_tabs is not None:
        pPr.remove(old_tabs)

    if not tab_positions:
        return

    tabs_elem = OxmlElement('w:tabs')
    for pos in tab_positions:
        tab_elem = OxmlElement('w:tab')
        tab_elem.set(qn('w:val'), 'left')
        tab_elem.set(qn('w:pos'), str(pos))
        tabs_elem.append(tab_elem)

    pPr.append(tabs_elem)


def _merge_paragraphs_with_tab(target_p: Paragraph, source_paras: List[Paragraph], tab_positions: List[int], indent_dxa: int = 0):
    """
    Gộp các source_paras vào target_p, phân cách bằng ký tự Tab.
    Sau đó xóa bỏ các source_paras khỏi tài liệu và thiết lập Tab Stops chuẩn.
    """
    _reset_paragraph_formatting_for_compact(target_p, indent_dxa)
    _clean_trailing_whitespace_and_tabs(target_p)
    target_elem = target_p._element

    for src_p in source_paras:
        _clean_leading_whitespace_and_tabs(src_p)
        _clean_trailing_whitespace_and_tabs(src_p)
        src_elem = src_p._element
        parent = src_elem.getparent()

        # 1. Thêm một run chứa ký tự Tab vào target_p
        tab_r = OxmlElement('w:r')
        tab_tag = OxmlElement('w:tab')
        tab_r.append(tab_tag)
        target_elem.append(tab_r)

        # 2. Chuyển toàn bộ các phần tử con của src_p (trừ w:pPr) sang target_p
        for child in list(src_elem):
            if child.tag == qn('w:pPr'):
                continue
            target_elem.append(child)

        # 3. Xóa source paragraph khỏi parent
        if parent is not None:
            parent.remove(src_elem)

    _clean_trailing_whitespace_and_tabs(target_p)
    _collapse_consecutive_tabs(target_p)
    _set_paragraph_tabs(target_p, tab_positions)


def _decide_group_layout(paras: List[Paragraph], mode: str) -> List[List[Paragraph]]:
    """
    Xác định cách phân nhóm dòng cho một nhóm các phương án (A, B, C, D...).
    
    Returns:
        Danh sách các dòng, mỗi dòng là một danh sách Paragraph cần gộp.
        Ví dụ:
        - 1 dòng: [[A, B, C, D]]
        - 2 dòng: [[A, B], [C, D]]
        - 4 dòng: [[A], [B], [C], [D]]
    """
    n = len(paras)
    if n <= 1:
        return [[p] for p in paras]

    # Kiểm tra xem có phương án nào chứa ảnh lớn không
    if any(_has_heavy_content(p) for p in paras):
        return [[p] for p in paras]

    lengths = [_get_paragraph_estimated_length(p) for p in paras]
    max_len = max(lengths)
    sum_len = sum(lengths)

    if mode == "4_per_line":
        if n == 4:
            return [paras]
        elif n == 2:
            return [paras]
        elif n == 3:
            return [paras]
        return [paras]

    elif mode == "2_per_line":
        if n == 4:
            return [[paras[0], paras[1]], [paras[2], paras[3]]]
        elif n == 2:
            return [paras]
        elif n == 3:
            return [[paras[0], paras[1]], [paras[2]]]
        else:
            # Nhóm 2 phần tử một dòng
            rows = []
            for i in range(0, n, 2):
                rows.append(paras[i:i+2])
            return rows

    elif mode == "1_per_line" or mode == "split":
        return [[p] for p in paras]

    else:
        # Chế độ 'auto' (Tự động thông minh Auto 4 - 2 - 1):
        # Tính toán chuẩn xác theo kích thước cột thực tế để không bao giờ bị tràn dòng:
        if n == 4:
            # 1. Điều kiện xếp 4 phương án / 1 dòng:
            # - Mỗi cột có khoảng trống ~3.7cm (tương đương tối đa ~15 ký tự).
            # - Cột D có khoảng trống ~5.4cm (tối đa ~20 ký tự).
            # - Toàn bộ 4 phương án phải ngắn: lengths[0..2] <= 15, lengths[3] <= 20, max_len <= 16, sum <= 58
            can_4_cols = (
                max_len <= 16 and 
                sum_len <= 58 and 
                lengths[0] <= 15 and 
                lengths[1] <= 15 and 
                lengths[2] <= 15 and 
                lengths[3] <= 20
            )
            if can_4_cols:
                return [paras]

            # 2. Điều kiện xếp 2 phương án / 1 dòng (2 dòng: A-B và C-D):
            # - Cột 1 (A hoặc C) có khoảng trống ~7.9cm (tương đương tối đa ~32 ký tự).
            # - Cột 2 (B hoặc D) có khoảng trống ~8.5cm (tương đương tối đa ~36 ký tự).
            # - Từng cặp dòng (A+B) <= 65 ký tự và (C+D) <= 65 ký tự.
            can_2_cols = (
                lengths[0] <= 32 and 
                lengths[1] <= 36 and 
                (lengths[0] + lengths[1] <= 65) and
                lengths[2] <= 32 and 
                lengths[3] <= 36 and 
                (lengths[2] + lengths[3] <= 65)
            )
            if can_2_cols:
                return [[paras[0], paras[1]], [paras[2], paras[3]]]

            # 3. Phương án dài hoặc có câu vượt quá độ rộng cột -> Giữ nguyên 4 dòng riêng biệt
            return [[p] for p in paras]

        elif n == 2:
            if lengths[0] <= 32 and lengths[1] <= 36 and sum_len <= 65:
                return [paras]
            else:
                return [[p] for p in paras]

        elif n == 3:
            if lengths[0] <= 18 and lengths[1] <= 18 and lengths[2] <= 22 and sum_len <= 54:
                return [paras]
            else:
                return [[p] for p in paras]

        else:
            return [[p] for p in paras]


def _clean_empty_paragraphs(doc: Document):
    """Xóa bỏ triệt để tất cả các đoạn văn rỗng rác (không text, không ảnh, không math)."""
    for p in list(doc.paragraphs):
        if _is_paragraph_empty(p):
            p_elem = p._element
            parent = p_elem.getparent()
            if parent is not None:
                parent.remove(p_elem)


def compact_document_choices(doc: Document, layout_mode: str = "auto", indent_dxa: int = 0) -> Tuple[int, int]:
    """
    Quét toàn bộ tài liệu Word và dồn dòng các phương án trắc nghiệm theo layout_mode.
    
    Args:
        doc: Đối tượng docx.Document
        layout_mode: 'auto', '4_per_line', '2_per_line', '1_per_line', 'split'
        indent_dxa: Khoảng cách thụt lề đầu dòng (twips), ví dụ 567 cho 1.0cm
        
    Returns:
        (total_groups_found, total_groups_compacted)
    """
    # 1. Chuyển đổi bảng trắc nghiệm thành paragraph
    _convert_choice_tables_to_paragraphs(doc)

    # 2. Tách các ngắt dòng mềm <w:br/> thành paragraph riêng
    _split_paragraphs_at_br(doc)

    # 3. Dọn dẹp triệt để tất cả các dòng trống rác xen kẽ
    _clean_empty_paragraphs(doc)

    all_paras = list(doc.paragraphs)
    
    # Tìm các chuỗi phương án liên tiếp
    # Mỗi group là một tuple: (is_lower, [Paragraph, ...], [empty_paras])
    groups = []
    current_group = []
    current_empty_paras = []
    current_is_lower = False
    current_expected_idx = 0
    current_parent = None

    for p in all_paras:
        # Nếu paragraph hoàn toàn rỗng do ngắt dòng kép:
        if _is_paragraph_empty(p):
            if current_group:
                current_empty_paras.append(p)
            continue

        label, is_lower, idx = _extract_choice_label_info(p)
        p_parent = p._element.getparent()

        if label is not None:
            # Kiểm tra xem có tiếp nối nhóm hiện tại không
            is_continuation = False
            if current_group and is_lower == current_is_lower and p_parent == current_parent:
                if idx == current_expected_idx:
                    is_continuation = True

            if is_continuation:
                current_group.append(p)
                current_expected_idx += 1
            else:
                # Nếu đã có nhóm trước đó >= 2 phương án, lưu lại
                if len(current_group) >= 2:
                    groups.append((current_is_lower, current_group, current_empty_paras))
                current_empty_paras = []
                
                # Bắt đầu nhóm mới nếu là phương án đầu tiên (A hoặc a)
                if idx == 0:
                    current_group = [p]
                    current_is_lower = is_lower
                    current_expected_idx = 1
                    current_parent = p_parent
                else:
                    current_group = []
                    current_parent = None
        else:
            if len(current_group) >= 2:
                groups.append((current_is_lower, current_group, current_empty_paras))
            current_group = []
            current_empty_paras = []
            current_parent = None

    if len(current_group) >= 2:
        groups.append((current_is_lower, current_group, current_empty_paras))

    total_groups = len(groups)
    compacted_count = 0

    if layout_mode in ("split", "1_per_line"):
        for is_lower, paras, empty_paras in groups:
            for ep in empty_paras:
                ep_parent = ep._element.getparent()
                if ep_parent is not None:
                    ep_parent.remove(ep._element)
            for p in paras:
                _reset_paragraph_formatting_for_compact(p, indent_dxa)
        logger.info("Chế độ bố cục là 'split' / '1_per_line' - đã định dạng thụt lề cho từng phương án.")
        return total_groups, 0

    for is_lower, paras, empty_paras in groups:
        # Dọn dẹp các paragraph rỗng xen kẽ
        for ep in empty_paras:
            ep_parent = ep._element.getparent()
            if ep_parent is not None:
                ep_parent.remove(ep._element)

        # Xác định cách chia dòng cho nhóm này
        rows = _decide_group_layout(paras, layout_mode)
        
        # Nếu số dòng đầu ra nhỏ hơn số paragraph ban đầu -> có dồn dòng
        if len(rows) < len(paras):
            compacted_count += 1

        for row in rows:
            if len(row) <= 1:
                # Reset định dạng ngay cả khi giữ 1 dòng để tránh bị lệch lề
                _reset_paragraph_formatting_for_compact(row[0], indent_dxa)
                continue

            target_p = row[0]
            source_paras = row[1:]
            
            # Xác định vị trí Tab Stops theo số lượng cột trong dòng (cộng thêm indent_dxa)
            num_cols = len(row)
            if num_cols == 4:
                tabs = [pos + indent_dxa for pos in TAB_STOPS_4_COLS]
            elif num_cols == 2:
                tabs = [pos + indent_dxa for pos in TAB_STOPS_2_COLS]
            elif num_cols == 3:
                tabs = [pos + indent_dxa for pos in TAB_STOPS_3_COLS]
            else:
                tabs = [indent_dxa + int(9000 / num_cols * c) for c in range(1, num_cols)]

            _merge_paragraphs_with_tab(target_p, source_paras, tabs, indent_dxa)

    logger.info(f"Đã xử lý dồn dòng: Tìm thấy {total_groups} nhóm câu hỏi, đã dồn {compacted_count} nhóm.")
    return total_groups, compacted_count


def compact_docx_choices(
    input_doc_or_path: Union[str, Document],
    output_path: Optional[str] = None,
    layout_mode: str = "auto",
    indent_dxa: int = 0
) -> Tuple[bool, int, str]:
    """
    Hàm giao diện cấp cao để dồn dòng phương án trắc nghiệm.
    
    Args:
        input_doc_or_path: Đường dẫn file Word (.docx) hoặc đối tượng Document
        output_path: Đường dẫn lưu file kết quả (nếu truyền đường dẫn input)
        layout_mode: 'auto', '4_per_line', '2_per_line', '1_per_line', 'split'
        indent_dxa: Khoảng cách thụt lề đầu dòng (twips)
        
    Returns:
        (success, compacted_count, message)
    """
    try:
        if isinstance(input_doc_or_path, str):
            doc = Document(input_doc_or_path)
            save_needed = True
        else:
            doc = input_doc_or_path
            save_needed = False

        total_groups, compacted_count = compact_document_choices(doc, layout_mode, indent_dxa)

        if save_needed:
            save_target = output_path or input_doc_or_path
            doc.save(save_target)

        mode_name_map = {
            "auto": "Tự động (Auto 4-2-1)",
            "4_per_line": "4 phương án / dòng",
            "2_per_line": "2 phương án / dòng",
            "1_per_line": "1 phương án / dòng",
            "split": "Tách 1 phương án / dòng"
        }
        mode_text = mode_name_map.get(layout_mode, layout_mode)
        indent_info = f" (thụt lề {indent_dxa/567:.2f}cm)" if indent_dxa > 0 else ""
        msg = f"Đã dồn dòng {compacted_count}/{total_groups} câu hỏi theo chế độ '{mode_text}'{indent_info}."
        return True, compacted_count, msg

    except Exception as e:
        logger.exception("Lỗi khi dồn dòng phương án: %s", e)
        return False, 0, str(e)
