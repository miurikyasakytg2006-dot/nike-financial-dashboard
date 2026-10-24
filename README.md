# NIKE FY2020–FY2026: dữ liệu tài chính và dashboard

Kho mã nguồn kỹ thuật của đề tài “Đổi mới và tái cân bằng chiến lược phân phối tại Nike: Từ Direct-to-Consumer (D2C) đến bán buôn và những biến động tài chính giai đoạn FY2020–FY2026”. Người dùng số liệu không cần đọc tài liệu này: xem `handover/GIAI-THICH-QUY-TRINH.md` và các file trong `handover/`.

Toàn bộ số liệu lấy từ 7 báo cáo Form 10-K của NIKE, Inc. nộp SEC (FY2020–FY2026). Sản phẩm học tập của sinh viên, không liên kết với NIKE, Inc.

## Cấu trúc

| Đường dẫn | Nội dung |
| --- | --- |
| `handover/NIKE_Financial_Dataset_FY2020-2026.xlsx` | Bộ dữ liệu chuẩn cho nhóm (tải lên Google Drive). Cũng là đầu vào của dashboard. |
| `pipeline/sources.py` | Danh mục 7 báo cáo 10-K (số accession, ngày nộp, URL SEC). |
| `pipeline/metrics.py` | Định nghĩa mọi chỉ tiêu báo cáo, công thức KPI, kiểm tra đối chiếu. Một công thức dùng chung cho Excel, Python và dashboard. |
| `pipeline/ixbrl.py` | Đọc Inline XBRL trong 10-K: giá trị, số trang, mục. |
| `pipeline/extract.py` | Tải gói XBRL và trích số liệu → `data/extracted/facts.csv`, đối chiếu API XBRL của SEC → `data/extracted/sec_api_check.csv`. |
| `pipeline/recheck_pages.py` | Đọc lại 35 số chỉ có trong MD&A trên trang 10-K hiển thị bằng trình duyệt: 28 tỷ lệ tăng trưởng loại trừ tỷ giá theo vị trí cột, 7 số Digital từ câu văn → `data/extracted/page_check.csv`, ảnh chụp và trang bằng chứng trong `handover/bang-chung-doi-chieu/` (được chép lên dashboard web). |
| `pipeline/build_workbook.py` | Tạo file Excel (giữ lại các cột xác minh nhóm đã điền). |
| `handover/GIAI-THICH-QUY-TRINH.md`, `.pdf` | Giải thích quy trình cho người trình bày và các thành viên (không cần đọc mã). Bản PDF tạo bằng `python -m pipeline.build_guide_pdf` sau mỗi lần sửa file .md. |
| `pipeline/verify_workbook.py` | Tính lại mọi công thức Excel bằng một engine độc lập và so với kết quả Python. |
| `pipeline/build_site_data.py` | Kiểm tra file Excel, tính KPI, xuất `site/data/nike-financials.json`. Dừng nếu có lỗi. |
| `pipeline/build_excel_dashboard.py`, `pipeline/excel_slicers.ps1` | Tạo bản Excel của dashboard (`handover/NIKE_Dashboard_FY2020-2026.xlsx`, mở bằng Excel desktop hoặc Excel for the web trên OneDrive; không mở bằng Google Sheets vì file dùng slicer và hàm AGGREGATE) từ cùng file JSON. Slicer năm và slicer cơ sở tăng trưởng được thêm bằng Excel trên Windows. |
| `pipeline/build_powerbi.py`, `powerbi/` | Tạo dự án Power BI (PBIP: mô hình TMDL có nhúng dữ liệu, báo cáo PBIR) từ cùng file JSON. Bản giao cho nhóm là `handover/NIKE_Dashboard_FY2020-2026.pbix`, lưu từ Power BI Desktop. |
| `site/` | Dashboard tĩnh (HTML, CSS, JavaScript, Apache ECharts 5.6.0). Thư mục được triển khai lên GitHub Pages. |
| `tests/` | Kiểm thử dữ liệu (`test_pipeline.py`) và kiểm thử trình duyệt (`test_dashboard.py`). |
| `.github/workflows/pages.yml` | Kiểm tra dữ liệu, chạy toàn bộ kiểm thử (dữ liệu và trình duyệt) rồi mới triển khai `site/` lên GitHub Pages. |

## Cài đặt

```bash
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt
.venv/bin/python -m playwright install chromium   # chỉ cần cho kiểm thử trình duyệt
```

Kiểm thử trình duyệt dùng Google Chrome nếu máy có cài, nếu không dùng Chromium của Playwright.

## Quy trình dữ liệu

```bash
.venv/bin/python -m pipeline.extract           # tải 10-K, trích số liệu, đối chiếu API SEC
.venv/bin/python -m pipeline.recheck_pages     # đọc lại 35 số MD&A trên trang 10-K (cần Chrome; nghỉ giữa các báo cáo)
.venv/bin/python -m pipeline.build_workbook    # tạo handover/NIKE_Financial_Dataset_FY2020-2026.xlsx
.venv/bin/python -m pipeline.verify_workbook   # tính lại công thức Excel độc lập
.venv/bin/python -m pipeline.build_site_data   # kiểm tra Excel và xuất JSON cho dashboard
.venv/bin/python -m pipeline.build_excel_dashboard   # bản Excel của dashboard (slicer cần Windows + Excel)
.venv/bin/python -m pytest tests -q            # 59 kiểm thử: dữ liệu, 3 bản dashboard và trình duyệt
```

Trên Windows thay `.venv/bin/python` bằng `.venv\Scripts\python.exe`.

Cập nhật bản Power BI sau khi dữ liệu đổi (sau `build_site_data`):

1. `.venv\Scripts\python.exe -m pipeline.build_powerbi`
2. Mở `powerbi\NIKE_Dashboard.pbip` bằng Power BI Desktop, bấm **Refresh**, chọn trang “Chiến lược phân phối”.
3. **File › Save as**, chọn kiểu “Power BI file (*.pbix)”, ghi đè `handover\NIKE_Dashboard_FY2020-2026.pbix`.

- File Excel là nguồn chuẩn. `build_site_data` đọc sheet 02 và 03, so từng giá trị và số trang với dữ liệu trích từ 10-K, và dừng (mã lỗi 1, không xuất JSON) nếu: thiếu dòng, trùng khóa, sai đơn vị hay phạm vi, ô trống không ghi “Thiếu / không công bố”, phần trăm nhập dạng số nguyên, thiếu nguồn, giá trị hoặc số trang khác số trích từ 10-K, “Đã xác minh (người)” thiếu tên hoặc ngày, có “Bằng chứng mâu thuẫn”, hoặc kiểm tra đối chiếu vượt ngưỡng. Bước này bảo vệ dashboard khỏi việc file Excel bị sửa sai; nó không thay cho việc kiểm tra chính số trích xuất (đã làm ở `extract` và `recheck_pages`).
- Khi nhóm đánh dấu “Đã xác minh (người)” trong file Excel trên Google Drive: tải file đó về, chép đè vào `handover/`, chạy lại `build_site_data`. Chạy lại `build_workbook` sẽ giữ nguyên các cột xác minh; nếu số liệu mới khác số đã xác minh, dòng đó chuyển sang “Bằng chứng mâu thuẫn”.
- Gói XBRL được tải từ bản sao chính thức trên NIKE Investor Relations (cùng số accession với EDGAR) vì SEC yêu cầu chương trình tự động khai email liên hệ. Mọi giá trị hợp nhất được đối chiếu với dữ liệu XBRL do SEC công bố (`data.sec.gov`); trước khi chạy `extract`, đặt biến môi trường `SEC_CONTACT_EMAIL`, hoặc mở đường dẫn API trong trình duyệt và lưu file vào `data/cache/companyfacts.json`. Trang NIKE IR chặn tạm thời (HTTP 403) nếu tải dồn dập; `recheck_pages` đã nghỉ giữa các báo cáo.

### Thêm năm tài chính mới (ví dụ FY2027)

1. Thêm một dòng vào `FILINGS` và sửa `LAST_FY` trong `pipeline/sources.py`.
2. Chạy toàn bộ quy trình trên; đọc kỹ các dòng `NOTE:` của `extract` (câu hoặc bảng NIKE đổi cách trình bày sẽ hiện ở đây).
3. Cập nhật các con số cố định trong kiểm thử (`tests/test_pipeline.py`) nếu cần, rồi xem lại dashboard ở cả 3 độ rộng màn hình.

## Xem dashboard trên máy

```bash
.venv/bin/python -m pipeline.build_site_data
.venv/bin/python -m http.server 8000 --directory site
```

Mở `http://localhost:8000`. Mở trực tiếp `site/index.html` từ ổ đĩa sẽ không tải được dữ liệu (trình duyệt chặn `fetch` từ `file://`).

## Triển khai lên GitHub Pages

1. Tạo kho trên GitHub (ví dụ `nike-financial-dashboard`), rồi đẩy mã nguồn:
   ```bash
   git init -b main
   git add .
   git commit -m "NIKE FY2020–FY2026 dataset and dashboard"
   git remote add origin https://github.com/<tai-khoan>/nike-financial-dashboard.git
   git push -u origin main
   ```
2. Trên GitHub: **Settings → Pages → Build and deployment → Source: GitHub Actions**.
3. Workflow `Deploy dashboard` tự chạy: kiểm tra dữ liệu, chạy toàn bộ kiểm thử, rồi triển khai. Đường dẫn có dạng `https://<tai-khoan>.github.io/nike-financial-dashboard/`.
4. Đường dẫn dashboard được ghi vào sheet 01 của file Excel từ hằng số `DASHBOARD_URL` trong `pipeline/build_workbook.py`. Đổi tài khoản hoặc tên kho thì sửa hằng số này rồi chạy lại `build_workbook`.

Trang chỉ dùng đường dẫn tương đối, không có máy chủ, khóa bí mật hay dịch vụ trả phí; khi triển khai, file Excel được chép vào `site/downloads/` và trang bằng chứng vào `site/bang-chung-doi-chieu/`.

## Ghi chú thiết kế

- **Câu hỏi trước, biểu đồ sau.** Hai góc nhìn ứng với hai câu hỏi của đề tài. Biểu đồ chọn theo công việc phân tích (FT Visual Vocabulary): thanh 100% cho cơ cấu kênh, cột nhóm cho quy mô, cột phân kỳ cho tăng trưởng, cột lồng nhau cho Digital trong NIKE Direct, các biểu đồ nhỏ cùng khung cho kết quả hợp nhất. Không dùng trục kép; biểu đồ đường không bắt đầu từ 0 đều ghi rõ.
- **Nguồn là một phần của giao diện.** Mỗi số NIKE công bố trong bảng kèm số trang 10-K có liên kết; số nhóm tự tính in nghiêng; số chưa được kiểm tra có dấu † (dấu và chú thích tự ẩn khi không còn số nào chờ kiểm tra); mỗi biểu đồ có dòng nguồn. Ảnh PNG xuất ra đã gồm tiêu đề, chú giải và nguồn; nút “Tải bảng (Excel)” xuất file .xlsx mở đúng ở mọi thiết lập vùng (khác với CSV). In trang (Ctrl+P) ra cả hai góc nhìn, khổ A4 ngang.
- **Màu theo đối tượng, không theo thứ hạng.** Xanh dương = NIKE Direct, cam = bán buôn ở mọi biểu đồ; xám xanh = số hợp nhất NIKE, Inc. Bảng màu đã chạy qua bộ kiểm tra mù màu (CVD ΔE ≥ 8); chữ đạt WCAG AA.
- **Chữ.** Source Sans 3 cho giao diện và số (chữ số đều độ rộng để cột số thẳng hàng), Source Serif 4 chỉ cho tiêu đề và hai câu hỏi. Cả hai hỗ trợ đầy đủ dấu tiếng Việt, giấy phép SIL OFL 1.1, lưu cục bộ trong `site/assets/fonts/`.
- **Tránh có chủ đích:** nhận diện thương hiệu NIKE, thẻ KPI kiểu SaaS, gradient, hiệu ứng trang trí, nút phóng to (7 điểm dữ liệu năm không cần zoom; bộ lọc năm thay thế).

## Giấy phép thành phần bên thứ ba

- Apache ECharts 5.6.0 — Apache License 2.0 (`site/assets/vendor/echarts-LICENSE.txt`).
- Source Sans 3, Source Serif 4 (gói Fontsource 5.3.0) — SIL Open Font License 1.1 (`site/assets/fonts/OFL-*.txt`).
- Số liệu tài chính: báo cáo công khai của NIKE, Inc. trên SEC EDGAR.
