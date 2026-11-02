# -*- coding: utf-8 -*-
import os
import sys
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>')
    tcPr.append(tcMar)

def create_report():
    doc = Document()

    # Configure Margins (Normal: 1 inch = 1440 dxa)
    for section in doc.sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(0.8)

    # Styles
    COLOR_PRIMARY = RGBColor(24, 76, 120)     # Navy Blue
    COLOR_SECONDARY = RGBColor(51, 65, 85)   # Slate 700
    COLOR_DARK = RGBColor(15, 23, 42)        # Slate 900
    COLOR_MUTED = RGBColor(100, 116, 139)    # Slate 500

    # Normal Style
    normal_style = doc.styles['Normal']
    normal_style.font.name = 'Times New Roman'
    normal_style.font.size = Pt(12)
    normal_style.font.color.rgb = COLOR_DARK
    normal_style.paragraph_format.line_spacing = 1.2
    normal_style.paragraph_format.space_after = Pt(4)

    # -------------------------------------------------------------
    # COVER PAGE
    # -------------------------------------------------------------
    p_univ = doc.add_paragraph()
    p_univ.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_univ.paragraph_format.space_after = Pt(2)
    r = p_univ.add_run("BỘ GIÁO DỤC VÀ ĐÀO TẠO\nTRƯỜNG ĐẠI HỌC LẠC HỒNG\nKHOA CÔNG NGHỆ THÔNG TIN")
    r.font.name = 'Times New Roman'
    r.font.size = Pt(13)
    r.font.bold = True
    r.font.color.rgb = COLOR_PRIMARY

    p_line = doc.add_paragraph()
    p_line.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_line.paragraph_format.space_after = Pt(40)
    r_line = p_line.add_run("———————————————***———————————————")
    r_line.font.bold = True
    r_line.font.color.rgb = COLOR_MUTED

    p_title_pre = doc.add_paragraph()
    p_title_pre.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title_pre.paragraph_format.space_after = Pt(8)
    r_tp = p_title_pre.add_run("BÁO CÁO BÀI TẬP CÁ NHÂN")
    r_tp.font.name = 'Times New Roman'
    r_tp.font.size = Pt(16)
    r_tp.font.bold = True
    r_tp.font.color.rgb = RGBColor(180, 83, 9) # Amber 700

    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_after = Pt(12)
    r_t = p_title.add_run("TRIỂN KHAI WEBSITE TỰ ĐỘNG\nVỚI GITHUB ACTIONS")
    r_t.font.name = 'Times New Roman'
    r_t.font.size = Pt(22)
    r_t.font.bold = True
    r_t.font.color.rgb = COLOR_PRIMARY

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub.paragraph_format.space_after = Pt(60)
    r_sub = p_sub.add_run("Môn học: THIẾT KẾ PHẦN MỀM (TKPM)")
    r_sub.font.name = 'Times New Roman'
    r_sub.font.size = Pt(14)
    r_sub.font.italic = True
    r_sub.font.color.rgb = COLOR_SECONDARY

    # Cover Info Table
    table_cover = doc.add_table(rows=4, cols=2)
    table_cover.alignment = WD_TABLE_ALIGNMENT.CENTER
    table_cover.autofit = False

    cover_data = [
        ("Sinh viên thực hiện:", "PHẠM BẢO PHONG"),
        ("Mã số sinh viên (MSSV):", "1822040447"),
        ("Giảng viên hướng dẫn:", "ThS. BÙI MINH PHỤNG"),
        ("Lớp / Chuyên ngành:", "Công Nghệ Thông Tin")
    ]

    for idx, (label, val) in enumerate(cover_data):
        row = table_cover.rows[idx]
        cell_lbl, cell_val = row.cells[0], row.cells[1]
        cell_lbl.width = Inches(2.6)
        cell_val.width = Inches(3.6)
        
        p0 = cell_lbl.paragraphs[0]
        p0.paragraph_format.space_after = Pt(6)
        r0 = p0.add_run(label)
        r0.font.bold = True
        r0.font.size = Pt(12)
        r0.font.color.rgb = COLOR_SECONDARY

        p1 = cell_val.paragraphs[0]
        p1.paragraph_format.space_after = Pt(6)
        r1 = p1.add_run(val)
        r1.font.bold = True
        r1.font.size = Pt(12)
        r1.font.color.rgb = COLOR_PRIMARY

    # Cover Footer
    p_foot = doc.add_paragraph()
    p_foot.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_foot.paragraph_format.space_before = Pt(80)
    p_foot.paragraph_format.space_after = Pt(0)
    r_f = p_foot.add_run("BIÊN HÒA, ĐỒNG NAI — THÁNG 11/2026")
    r_f.font.bold = True
    r_f.font.size = Pt(12)
    r_f.font.color.rgb = COLOR_SECONDARY

    doc.add_page_break()

    # -------------------------------------------------------------
    # HELPER FUNCTIONS FOR REPORT BODY
    # -------------------------------------------------------------
    def add_h1(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(14)
        p.paragraph_format.space_after = Pt(6)
        r = p.add_run(text)
        r.font.name = 'Times New Roman'
        r.font.size = Pt(15)
        r.font.bold = True
        r.font.color.rgb = COLOR_PRIMARY
        return p

    def add_h2(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.space_after = Pt(4)
        r = p.add_run(text)
        r.font.name = 'Times New Roman'
        r.font.size = Pt(13)
        r.font.bold = True
        r.font.color.rgb = RGBColor(14, 116, 144) # Cyan/Blue 700
        return p

    def add_h3(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(6)
        p.paragraph_format.space_after = Pt(2)
        r = p.add_run(text)
        r.font.name = 'Times New Roman'
        r.font.size = Pt(12)
        r.font.bold = True
        r.font.color.rgb = COLOR_SECONDARY
        return p

    def add_caption(text):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(12)
        r = p.add_run(text)
        r.font.name = 'Times New Roman'
        r.font.size = Pt(10.5)
        r.font.italic = True
        r.font.color.rgb = COLOR_SECONDARY
        return p

    def add_code_block(code_text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.left_indent = Inches(0.2)
        p.paragraph_format.line_spacing = 1.05
        
        table = doc.add_table(rows=1, cols=1)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        cell = table.cell(0, 0)
        cell.width = Inches(6.2)
        set_cell_background(cell, "F8FAFC")
        set_cell_margins(cell, top=120, bottom=120, left=150, right=150)
        
        cp = cell.paragraphs[0]
        cp.paragraph_format.space_after = Pt(0)
        cp.paragraph_format.line_spacing = 1.1
        cr = cp.add_run(code_text)
        cr.font.name = 'Consolas'
        cr.font.size = Pt(9.5)
        cr.font.color.rgb = RGBColor(30, 41, 59)

    def add_callout(text, title="LƯU Ý QUAN TRỌNG:"):
        table = doc.add_table(rows=1, cols=1)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        cell = table.cell(0, 0)
        cell.width = Inches(6.2)
        set_cell_background(cell, "EFF6FF")
        set_cell_margins(cell, top=120, bottom=120, left=160, right=160)
        
        cp = cell.paragraphs[0]
        cp.paragraph_format.space_after = Pt(0)
        r_t = cp.add_run(title + " ")
        r_t.font.bold = True
        r_t.font.size = Pt(10.5)
        r_t.font.color.rgb = RGBColor(29, 78, 216)
        
        r_c = cp.add_run(text)
        r_c.font.size = Pt(10.5)
        r_c.font.color.rgb = RGBColor(30, 58, 138)

    # -------------------------------------------------------------
    # MỤC LỤC TÓM TẮT
    # -------------------------------------------------------------
    add_h1("TỔNG QUAN NỘI DUNG BÁO CÁO")
    p_intro = doc.add_paragraph(
        "Báo cáo này trình bày toàn diện quá trình thực hiện bài tập cá nhân môn Thiết Kế Phần Mềm: "
        "\"Triển khai Website tự động với GitHub Actions\", tuân thủ đầy đủ 5 bước theo yêu cầu của giảng viên "
        "từ khâu xây dựng mã nguồn web, cấu hình Git/GitHub, lập trình kịch bản CI/CD pipeline, triển khai "
        "tự động lên GitHub Pages, cho đến thực nghiệm kiểm chứng cơ chế chặn mã nguồn lỗi của hệ thống."
    )
    p_intro.paragraph_format.space_after = Pt(8)

    # Info Summary Box
    tbl_sum = doc.add_table(rows=6, cols=2)
    tbl_sum.alignment = WD_TABLE_ALIGNMENT.CENTER
    summary_items = [
        ("Sinh viên thực hiện", "Phạm Bảo Phong — MSSV: 1822040447"),
        ("Giảng viên hướng dẫn", "ThS. Bùi Minh Phụng"),
        ("Địa điểm đào tạo", "Biên Hòa, Đồng Nai — Trường Đại học Lạc Hồng"),
        ("GitHub Repository", "https://github.com/Phong-Pham-Bao/devops-1822040447"),
        ("Địa chỉ Website Công Khai", "https://phong-pham-bao.github.io/devops-1822040447/"),
        ("Công nghệ sử dụng", "HTML5, CSS3, JavaScript, Git, GitHub Actions, GitHub Pages")
    ]
    for idx, (k, v) in enumerate(summary_items):
        r = tbl_sum.rows[idx]
        c0, c1 = r.cells[0], r.cells[1]
        c0.width = Inches(2.2)
        c1.width = Inches(4.0)
        set_cell_background(c0, "F1F5F9")
        set_cell_margins(c0, 60, 60, 100, 100)
        set_cell_margins(c1, 60, 60, 100, 100)
        
        p0 = c0.paragraphs[0]
        p0.paragraph_format.space_after = Pt(2)
        r0 = p0.add_run(k)
        r0.font.bold = True
        r0.font.size = Pt(11)
        r0.font.color.rgb = COLOR_SECONDARY
        
        p1 = c1.paragraphs[0]
        p1.paragraph_format.space_after = Pt(2)
        r1 = p1.add_run(v)
        r1.font.size = Pt(11)

    doc.add_paragraph().paragraph_format.space_after = Pt(4)

    # -------------------------------------------------------------
    # PHẦN 1: MỤC TIÊU & CƠ SỞ LÝ THUYẾT CI/CD
    # -------------------------------------------------------------
    add_h1("PHẦN 1: MỤC TIÊU VÀ CƠ SỞ LÝ THUYẾT")
    
    add_h2("1.1. Mục tiêu bài tập")
    doc.add_paragraph(
        "Bài tập cá nhân nhằm trang bị và củng cố cho sinh viên các kiến thức thực tiễn cốt lõi:"
    )
    doc.add_paragraph(
        "• Nắm vững quy trình CI/CD cơ bản trong kỹ thuật phần mềm: Viết mã nguồn → Push lên kho lưu trữ GitHub → "
        "Kích hoạt pipeline tự động kiểm tra cú pháp và tính toàn vẹn → Tự động xuất bản (deploy) sản phẩm lên môi trường thực tế.\n"
        "• Hiểu sâu sắc vai trò của Pipeline đóng vai trò như một \"Cổng kiểm soát chất lượng\" (Quality Gate) kiên cố: "
        "ngăn chặn triệt để các phiên bản mã nguồn lỗi hoặc thiếu sót tiếp cận người dùng cuối.\n"
        "• Thành thạo thao tác cấu hình GitHub Actions (Jobs, Steps, Dependencies `needs`, Runner Ubuntu) và tích hợp triển khai GitHub Pages."
    )

    add_h2("1.2. Khái niệm CI/CD và GitHub Actions")
    doc.add_paragraph(
        "CI/CD là viết tắt của Continuous Integration (Tích hợp liên tục) và Continuous Deployment / Delivery (Triển khai liên tục):\n"
        "• Continuous Integration (CI): Cơ chế tự động hóa việc tích hợp các thay đổi mã nguồn từ nhiều thành viên. Mỗi khi có lệnh commit hoặc push, hệ thống tự động build và chạy các bài test định sẵn để phát hiện xung đột hoặc lỗi sớm nhất.\n"
        "• Continuous Deployment (CD): Khi pha CI vượt qua thành công, mã nguồn sẽ tự động được đóng gói thành các Artifacts và chuyển giao trực tiếp lên máy chủ môi trường phát hành (Production Environment) mà không cần sự can thiệp thủ công của con người."
    )
    doc.add_paragraph(
        "GitHub Actions là nền tảng tự động hóa CI/CD mạnh mẽ được tích hợp trực tiếp bên trong GitHub. "
        "Mỗi quy trình làm việc (workflow) được định nghĩa bằng tệp tin cấu hình YAML đặt trong thư mục `.github/workflows/`, "
        "cho phép quản lý vòng đời phát triển ứng dụng một cách chuyên nghiệp, minh bạch và có tính lặp lại cao."
    )

    # -------------------------------------------------------------
    # PHẦN 2: QUY TRÌNH THỰC HIỆN CHI TIẾT (BƯỚC 1 - BƯỚC 5)
    # -------------------------------------------------------------
    add_h1("PHẦN 2: QUY TRÌNH THỰC HIỆN CHI TIẾT")

    # BƯỚC 1
    add_h2("Bước 1: Thiết kế và Xây dựng Website Hoàn chỉnh")
    doc.add_paragraph(
        "Theo yêu cầu của đề bài, thư mục dự án được cấu trúc gồm 3 tệp tin mã nguồn cơ bản: index.html, style.css, và script.js. "
        "Đặc biệt, thẻ tiêu đề chính <h1> phải thể hiện rõ ràng họ tên và MSSV của sinh viên: \"Phạm Bảo Phong - 1822040447\"."
    )
    
    add_h3("a) Cấu trúc tệp tin index.html")
    doc.add_paragraph(
        "Tệp index.html được xây dựng theo chuẩn ngữ nghĩa HTML5, tích hợp đầy đủ thẻ <title>, thẻ meta viewport "
        "phục vụ tương thích đa thiết bị, khối thông tin sinh viên, bảng trạng thái Pipeline và nút bấm kiểm tra JavaScript:"
    )
    add_code_block(
'''<!DOCTYPE html>
<html lang="vi">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>CI/CD Deployment Demo - Phạm Bảo Phong</title>
  <link rel="stylesheet" href="style.css">
</head>
<body>
  <div class="container">
    <header class="card">
      <span class="badge" id="versionBadge">Phiên bản 1.0</span>
      <h1>Phạm Bảo Phong - 1822040447</h1>
      <p class="subtitle">Báo cáo bài tập cá nhân: Triển khai Website tự động với GitHub Actions</p>
    </header>
    <main class="card">
      <!-- Thông tin sinh viên, giảng viên và trạng thái pipeline -->
    </main>
  </div>
  <script src="script.js"></script>
</body>
</html>'''
    )

    add_h3("b) Tệp định dạng giao diện style.css và kịch bản tương tác script.js")
    doc.add_paragraph(
        "• style.css: Được thiết kế theo phong cách hiện đại (Glassmorphism & Gradient), hỗ trợ linh hoạt cả hai chế độ "
        "Dark Mode và Light Mode, căn chỉnh tỷ lệ chuẩn trên cả màn hình máy tính và điện thoại thông minh.\n"
        "• script.js: Xử lý chức năng đồng hồ thời gian thực (Live System Clock), hiệu ứng chuyển đổi chủ đề (Theme Toggle) "
        "và thông báo xác thực tương tác hệ thống khi người dùng nhấn nút."
    )

    img_p1 = r"1822040447_PhamBaoPhong_TKPM\images\img_01_website_live.png"
    if os.path.exists(img_p1):
        doc.add_picture(img_p1, width=Inches(6.0))
        add_caption("Hình 1: Giao diện trực quan của Website với thẻ <h1> chứa họ tên và MSSV Phạm Bảo Phong - 1822040447.")

    # BƯỚC 2
    add_h2("Bước 2: Khởi tạo Git và Đưa Mã Nguồn Lên GitHub")
    doc.add_paragraph(
        "Để quản lý phiên bản và kết nối tới pipeline tự động, sinh viên tiến hành khởi tạo kho lưu trữ Git cục bộ "
        "và tạo một Repository Công khai (Public) trên GitHub với tên quy định là devops-1822040447."
    )
    
    add_h3("a) Các câu lệnh Git thực thi")
    add_code_block(
'''# Di chuyển vào thư mục dự án
cd "d:\\Dự Án\\Bài Học\\1822040447_PhamBaoPhong_TKPM"

# Khởi tạo Git repository trên nhánh main
git init -b main

# Thêm toàn bộ mã nguồn vào staging và tạo commit đầu tiên
git add .
git commit -m "feat: khoi tao website va pipeline GitHub Actions (Phien ban 1.0)"

# Tạo repository Public và đẩy mã nguồn lên GitHub bằng GitHub CLI
gh repo create devops-1822040447 --public --source=. --remote=origin --push'''
    )

    add_h3("b) Cấu hình GitHub Pages sử dụng nguồn GitHub Actions")
    doc.add_paragraph(
        "Truy cập vào kho lưu trữ trên GitHub → Vào mục Settings → Chọn mục Pages ở thanh menu bên trái. "
        "Tại phần Build and deployment → Source: chọn giá trị \"GitHub Actions\" (thay cho lựa chọn mặc định Deploy from a branch). "
        "Điều này báo hiệu cho GitHub biết rằng trang web sẽ được đóng gói và xuất bản thông qua Artifact do Workflow tạo ra."
    )

    img_p2 = r"1822040447_PhamBaoPhong_TKPM\images\img_02_github_repo.png"
    if os.path.exists(img_p2):
        doc.add_picture(img_p2, width=Inches(6.0))
        add_caption("Hình 2: Kho lưu trữ công khai devops-1822040447 được khởi tạo thành công trên tài khoản GitHub Phong-Pham-Bao.")

    # BƯỚC 3
    add_h2("Bước 3: Xây dựng Kịch bản CI/CD Pipeline với GitHub Actions")
    doc.add_paragraph(
        "Pipeline tự động được định nghĩa trong tệp `.github/workflows/deploy.yml`. "
        "Kịch bản gồm 2 Job riêng biệt hoạt động tuần tự với sự phụ thuộc điều kiện:"
    )
    doc.add_paragraph(
        "1. Job `check` (Kiểm tra chất lượng mã nguồn - Code Quality Check):\n"
        "   - Chạy trên môi trường ảo `ubuntu-latest`.\n"
        "   - Tải mã nguồn về Runner bằng action chính thức `actions/checkout@v4`.\n"
        "   - Thực hiện bước kiểm tra 1: Duyệt vòng lặp kiểm tra sự tồn tại của cả 3 tệp bắt buộc: index.html, style.css, script.js. Nếu thiếu bất kỳ tệp nào, trả về lỗi exit 1 và dừng lại.\n"
        "   - Thực hiện bước kiểm tra 2: Sử dụng lệnh `grep -qi \"<title>.*</title>\" index.html` để quét và xác minh tệp index.html có chứa thẻ <title> hợp lệ hay không. Nếu không tìm thấy thẻ title, trả về lỗi exit 1."
    )
    doc.add_paragraph(
        "2. Job `deploy` (Triển khai lên GitHub Pages):\n"
        "   - Được gán phụ thuộc qua khai báo `needs: check`. Nghĩa là Job deploy CHỈ ĐƯỢC PHÉP CHẠY khi Job check đã vượt qua thành công 100%.\n"
        "   - Được cấp quyền bảo mật thông qua `permissions: pages: write, id-token: write, contents: read`.\n"
        "   - Sử dụng các action tiêu chuẩn: `actions/configure-pages@v5`, `actions/upload-pages-artifact@v3` (đóng gói thư mục gốc thành artifact) và `actions/deploy-pages@v4` (xuất bản trực tiếp lên GitHub Pages CDN)."
    )

    add_h3("Nội dung hoàn chỉnh tệp cấu hình deploy.yml:")
    add_code_block(
'''name: Deploy Website to GitHub Pages

on:
  push:
    branches:
      - main
  workflow_dispatch:

permissions:
  contents: read
  pages: write
  id-token: write

concurrency:
  group: "pages"
  cancel-in-progress: false

jobs:
  check:
    name: Kiểm tra mã nguồn (Code Quality Check)
    runs-on: ubuntu-latest
    steps:
      - name: Tải mã nguồn về Runner
        uses: actions/checkout@v4

      - name: Kiểm tra tối thiểu 3 file (index.html, style.css, script.js)
        run: |
          echo "=== KIỂM TRA 3 FILE NGUỒN ==="
          MISSING_FILES=0
          for file in index.html style.css script.js; do
            if [ -f "$file" ]; then
              echo "✓ Đã tìm thấy: $file"
            else
              echo "✗ LỖI: Không tìm thấy: $file"
              MISSING_FILES=1
            fi
          done
          if [ $MISSING_FILES -ne 0 ]; then
            exit 1
          fi

      - name: Kiểm tra file index.html có thẻ <title>
        run: |
          echo "=== KIỂM TRA THẺ <title> ==="
          if grep -qi "<title>.*</title>" index.html; then
            echo "✓ Hợp lệ: Đã tìm thấy thẻ <title>"
            grep -i "<title>.*</title>" index.html
          else
            echo "✗ LỖI: index.html thiếu thẻ <title> bắt buộc!"
            exit 1
          fi

  deploy:
    name: Triển khai lên GitHub Pages (Deploy)
    needs: check
    environment:
      name: github-pages
      url: ${{ steps.deployment.outputs.page_url }}
    runs-on: ubuntu-latest
    steps:
      - name: Tải mã nguồn về Runner
        uses: actions/checkout@v4

      - name: Thiết lập môi trường GitHub Pages
        uses: actions/configure-pages@v5

      - name: Nén và tải lên Artifact
        uses: actions/upload-pages-artifact@v3
        with:
          path: '.'

      - name: Xuất bản lên GitHub Pages
        id: deployment
        uses: actions/deploy-pages@v4'''
    )

    # BƯỚC 4
    add_h2("Bước 4: Triển khai Thành công Lên GitHub Pages")
    doc.add_paragraph(
        "Sau khi đẩy commit đầu tiên lên nhánh `main`, GitHub Actions đã tự động bắt sự kiện `push` "
        "và tiến hành thực thi pipeline. Kết quả ghi nhận tại Run ID #41498641858:"
    )
    doc.add_paragraph(
        "• Job `check`: Hoàn thành trong vòng 5 giây, xác nhận đầy đủ 3 tệp nguồn và thẻ <title> hợp lệ.\n"
        "• Job `deploy`: Tự động kích hoạt sau khi Job check hoàn thành, đóng gói artifact và xuất bản thành công sau 20 giây.\n"
        "• Trang web chính thức có thể truy cập công khai tại URL: https://phong-pham-bao.github.io/devops-1822040447/"
    )

    img_p3 = r"1822040447_PhamBaoPhong_TKPM\images\img_03_actions_list.png"
    if os.path.exists(img_p3):
        doc.add_picture(img_p3, width=Inches(6.0))
        add_caption("Hình 3: Lịch sử thực thi các workflow trong tab Actions trên GitHub.")

    img_p8 = r"1822040447_PhamBaoPhong_TKPM\images\img_08_job_deploy_success_detail.png"
    if os.path.exists(img_p8):
        doc.add_picture(img_p8, width=Inches(6.0))
        add_caption("Hình 4: Chi tiết quá trình thực thi thành công của Job deploy xuất bản lên môi trường github-pages.")

    # BƯỚC 5
    add_h2("Bước 5: Thực nghiệm Pipeline Thất bại và Khôi phục Phiên bản 2.0")
    doc.add_paragraph(
        "Đây là phần thực nghiệm trọng tâm nhằm chứng minh năng lực bảo vệ của CI/CD Pipeline "
        "đối với môi trường sản xuất khi xảy ra lỗi vô ý của lập trình viên."
    )

    add_h3("5.1. Kịch bản lỗi: Cố tình xóa thẻ <title> để kiểm chứng cơ chế chặn")
    doc.add_paragraph(
        "Sinh viên tiến hành xóa dòng thẻ `<title>` khỏi tệp index.html, sau đó tạo commit và push lên nhánh main:\n"
        "• Commit: `test: co tinh xoa the title de kiem thu pipeline that bai` (Mã băm commit: `a81805e`).\n"
        "• Hệ thống GitHub Actions tự động kích hoạt Run ID #41498855062."
    )
    doc.add_paragraph(
        "Kết quả quan sát thực tế:\n"
        "1. Tại Job `check`: Bước kiểm tra 1 (đủ 3 file) vượt qua bình thường. Tuy nhiên, đến bước kiểm tra 2, "
        "lệnh `grep -qi \"<title>.*</title>\"` trả về false, dòng lệnh in ra thông báo: `✗ LỖI: File index.html thiếu thẻ <title> bắt buộc!` "
        "và thoát với mã lỗi exit 1. Toàn bộ Job check bị đánh dấu thất bại (Màu đỏ).\n"
        "2. Tại Job `deploy`: Nhờ khai báo `needs: check`, hệ thống nhận diện điều kiện tiên quyết không thỏa mãn, "
        "Job deploy lập tức chuyển sang trạng thái `skipped` (bỏ qua), hoàn toàn không được khởi chạy.\n"
        "3. Đánh giá tính an toàn: Trang web công khai trên GitHub Pages vẫn giữ nguyên phiên bản cũ ổn định, "
        "mã nguồn lỗi bị chặn đứng tại cổng CI, không gây gián đoạn hay ảnh hưởng tới người dùng!"
    )

    img_p4 = r"1822040447_PhamBaoPhong_TKPM\images\img_04_pipeline_failed.png"
    if os.path.exists(img_p4):
        doc.add_picture(img_p4, width=Inches(6.0))
        add_caption("Hình 5: Pipeline tự động phát hiện mã nguồn vi phạm quy tắc — Job check báo đỏ và Job deploy bị hủy bỏ (Skipped).")

    img_p7 = r"1822040447_PhamBaoPhong_TKPM\images\img_07_job_failure_detail.png"
    if os.path.exists(img_p7):
        doc.add_picture(img_p7, width=Inches(6.0))
        add_caption("Hình 6: Nhật ký lỗi (Log) của bước kiểm tra thẻ <title> thông báo lỗi và dừng tiến trình với exit code 1.")

    add_h3("5.2. Kịch bản khôi phục và nâng cấp lên \"Phiên bản 2.0\"")
    doc.add_paragraph(
        "Sau khi kiểm chứng thành công cơ chế phát hiện lỗi, sinh viên tiến hành sửa chữa và cải tiến trang web:\n"
        "• Khôi phục lại thẻ `<title>` với nội dung: `<title>CI/CD Deployment Demo - Phạm Bảo Phong (Phiên bản 2.0)</title>`.\n"
        "• Cập nhật giao diện với huy hiệu `Phiên bản 2.0` và trạng thái xác nhận pipeline hoạt động.\n"
        "• Tạo commit và push lên nhánh main: `fix: khoi phuc the title va nang cap len Phien ban 2.0` (Mã băm commit: `8946c42`)."
    )
    doc.add_paragraph(
        "Kết quả quan sát thực tế (Run ID #41499011709):\n"
        "• Job `check`: Kiểm tra thành công cả 3 tệp nguồn và thẻ <title>, thời gian thực thi: 7 giây.\n"
        "• Job `deploy`: Tự động tiếp nhận và xuất bản phiên bản mới lên GitHub Pages sau 49 giây.\n"
        "• Truy cập vào địa chỉ https://phong-pham-bao.github.io/devops-1822040447/ xác nhận trang web đã được cập nhật "
        "chính xác lên giao diện Phiên bản 2.0 mà không cần bất kỳ thao tác thủ công nào trên máy chủ."
    )

    img_p5 = r"1822040447_PhamBaoPhong_TKPM\images\img_05_pipeline_success.png"
    if os.path.exists(img_p5):
        doc.add_picture(img_p5, width=Inches(6.0))
        add_caption("Hình 7: Pipeline chuyển sang trạng thái xanh toàn bộ (Success) sau khi khắc phục lỗi và nâng cấp lên Phiên bản 2.0.")

    img_p6 = r"1822040447_PhamBaoPhong_TKPM\images\img_06_deployments_activity.png"
    if os.path.exists(img_p6):
        doc.add_picture(img_p6, width=Inches(6.0))
        add_caption("Hình 8: Nhật ký hoạt động triển khai (Deployment Activity) xác nhận Phiên bản 2.0 đang hoạt động (Active).")

    # -------------------------------------------------------------
    # PHẦN 3: KẾT LUẬN & BÀI HỌC KINH NGHIỆM
    # -------------------------------------------------------------
    add_h1("PHẦN 3: KẾT LUẬN VÀ BÀI HỌC KINH NGHIỆM")
    
    add_h2("3.1. Kết quả đạt được")
    doc.add_paragraph(
        "Thông qua bài tập cá nhân, sinh viên đã hoàn thành xuất sắc 100% mục tiêu đề ra:\n"
        "1. Xây dựng website hoàn chỉnh với đầy đủ 3 tệp tin tiêu chuẩn (HTML, CSS, JS), hiển thị chính xác thông tin định danh sinh viên Phạm Bảo Phong - MSSV: 1822040447.\n"
        "2. Tạo lập kho lưu trữ công khai `devops-1822040447` và kết nối với dịch vụ GitHub Pages.\n"
        "3. Lập trình thành công tệp kịch bản pipeline CI/CD với 2 Job rõ ràng: `check` và `deploy`, sử dụng phụ thuộc `needs: check` để kiểm soát chất lượng.\n"
        "4. Triển khai website tự động, nhanh chóng và ổn định lên môi trường Internet công khai.\n"
        "5. Thực nghiệm thành công hai kịch bản: phát hiện và chặn đứng mã nguồn lỗi (khi thiếu thẻ title), và tự động cập nhật sản phẩm khi sửa lỗi nâng cấp lên Phiên bản 2.0."
    )

    add_h2("3.2. Ý nghĩa thực tiễn của Pipeline trong phát triển phần mềm")
    doc.add_paragraph(
        "Thực nghiệm cho thấy rõ ràng tầm quan trọng của quy trình DevOps và CI/CD trong các dự án công nghệ hiện đại:\n"
        "• Loại bỏ rủi ro do thao tác thủ công: Trong các quy trình truyền thống (sử dụng FTP hoặc copy paste mã nguồn lên server), lập trình viên rất dễ quên file, gửi nhầm code lỗi hoặc ghi đè dữ liệu.\n"
        "• Tiêu chuẩn hóa quy trình kiểm thử (Automated Testing): Bất kỳ đoạn mã nào muốn xuất bản đều phải vượt qua một chuỗi các tiêu chí kiểm tra nghiêm ngặt (kiểm tra cú pháp, đơn vị, liên kết, bảo mật).\n"
        "• Bảo vệ trải nghiệm người dùng cuối: Dù lỗi xảy ra trong quá trình phát triển, hệ thống vẫn duy trì phiên bản trước đó hoạt động bình thường, không làm gián đoạn dịch vụ."
    )

    add_callout(
        "Em xin chân thành cảm ơn Thầy Bùi Minh Phụng đã tận tình giảng dạy và hướng dẫn những kiến thức thực tế vô cùng quý báu về Thiết kế phần mềm và văn hóa CI/CD / DevOps, giúp em có thể làm chủ quy trình phát triển và triển khai phần mềm hiện đại!",
        "LỜI CẢM ƠN:"
    )

    # Save to project folder
    output_docx = r"1822040447_PhamBaoPhong_TKPM\BaoCao_1822040447_PhamBaoPhong_TKPM.docx"
    doc.save(output_docx)
    print(f"Báo cáo Word đã được lưu thành công tại: {output_docx}")

    # Also save a copy in root workspace for convenience
    root_docx = r"BaoCao_1822040447_PhamBaoPhong_TKPM.docx"
    doc.save(root_docx)
    print(f"Đã sao chép 1 bản ra thư mục gốc: {root_docx}")

if __name__ == '__main__':
    create_report()
