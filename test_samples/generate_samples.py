"""Test script to generate a sample Word test document and administrative document."""
import os
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH

os.makedirs("test_samples", exist_ok=True)

# 1. Tạo file đề thi mẫu (có bảng đáp án cuối đề, có câu trả lời ngắn, có inline choices)
doc_exam = Document()

# Header / Tiêu đề đề thi
p_title = doc_exam.add_paragraph("KỲ THI TRUNG HỌC PHỔ THÔNG QUỐC GIA")
p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_sub = doc_exam.add_paragraph("MÔN THI: VẬT LÍ / TOÁN HỌC")
p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER

# Phần 1: Trắc nghiệm 4 lựa chọn
doc_exam.add_paragraph("PHẦN I. Câu trắc nghiệm nhiều phương án lựa chọn.")

doc_exam.add_paragraph("Câu 1: Một vật dao động điều hòa với tần số góc omega. Chu kỳ dao động của vật là:")
doc_exam.add_paragraph("A. T = 2pi / omega")
doc_exam.add_paragraph("B. T = omega / 2pi")
doc_exam.add_paragraph("C. T = 2pi * omega")
doc_exam.add_paragraph("D. T = 1 / omega")

doc_exam.add_paragraph("Câu 2: Đơn vị của điện tích trong hệ SI là gì?")
doc_exam.add_paragraph("A. Vôn (V)    B. Ampe (A)    C. Cu-lông (C)    D. Jun (J)")

# Phần 2: Đúng / Sai
doc_exam.add_paragraph("PHẦN II. Câu trắc nghiệm đúng sai.")
doc_exam.add_paragraph("Câu 3: Cho hàm số f(x) = x^3 - 3x + 2. Xét tính đúng sai của các mệnh đề sau:")
doc_exam.add_paragraph("a) Hàm số đồng biến trên khoảng (1; +vô cùng).")
doc_exam.add_paragraph("b) Hàm số có điểm cực đại tại x = 1.")
doc_exam.add_paragraph("c) Giá trị cực tiểu của hàm số bằng 0.")
doc_exam.add_paragraph("d) Đồ thị hàm số đi qua gốc tọa độ O(0; 0).")

# Phần 3: Trả lời ngắn
doc_exam.add_paragraph("PHẦN III. Câu trắc nghiệm trả lời ngắn.")
doc_exam.add_paragraph("Câu 4: Tìm số nguyên dương m nhỏ nhất để hàm số luôn đồng biến trên R.")

# Bảng đáp án ở cuối đề
doc_exam.add_paragraph("BẢNG ĐÁP ÁN")
table = doc_exam.add_table(rows=3, cols=3)
# Row 0: Header / Question num
table.rows[0].cells[0].text = "Câu"
table.rows[0].cells[1].text = "1"
table.rows[0].cells[2].text = "2"
# Row 1: Keys Part 1
table.rows[1].cells[0].text = "Đáp án"
table.rows[1].cells[1].text = "A"
table.rows[1].cells[2].text = "C"

# Part 2 answer text
doc_exam.add_paragraph("Đáp án Phần II:")
doc_exam.add_paragraph("3(Đ,S,Đ,S)")

# Part 3 answer text
doc_exam.add_paragraph("Đáp án Phần III:")
doc_exam.add_paragraph("4: 15.5")

doc_exam.save("test_samples/de_thi_mau.docx")
print("Đã tạo file test_samples/de_thi_mau.docx")

# 2. Tạo file văn bản hành chính mẫu
doc_admin = Document()
p1 = doc_admin.add_paragraph("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM")
p2 = doc_admin.add_paragraph("Độc lập - Tự do - Hạnh phúc")
p3 = doc_admin.add_paragraph("QUYẾT ĐỊNH")
p4 = doc_admin.add_paragraph("Về việc ban hành quy chế làm việc của cơ quan")
p5 = doc_admin.add_paragraph("Căn cứ Nghị định số 30/2020/NĐ-CP ngày 05 tháng 3 năm 2020 của Chính phủ về công tác văn thư.")
p6 = doc_admin.add_paragraph("Điều 1. Ban hành kèm theo Quyết định này Quy chế làm việc mới áp dụng cho toàn thể cán bộ, nhân viên.")
p7 = doc_admin.add_paragraph("Điều 2. Quyết định này có hiệu lực thi hành kể từ ngày ký.")

doc_admin.save("test_samples/van_ban_mau.docx")
print("Đã tạo file test_samples/van_ban_mau.docx")
