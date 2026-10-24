# NIKE FY2020–FY2026: dữ liệu tài chính và dashboard

Mã nguồn phần số liệu của đề tài “Đổi mới và tái cân bằng chiến lược phân phối tại Nike: Từ Direct-to-Consumer (D2C) đến bán buôn và những biến động tài chính giai đoạn FY2020–FY2026”. Nếu chỉ cần dùng số liệu, bạn không cần đọc tài liệu này: hãy xem `handover/GIAI-THICH-QUY-TRINH.md` và các file trong thư mục `handover/`.

Toàn bộ số liệu lấy từ 7 báo cáo Form 10-K mà NIKE, Inc. nộp cho SEC (FY2020–FY2026). Đây là sản phẩm học tập của sinh viên, không liên quan đến NIKE, Inc.

## Cấu trúc

| Đường dẫn | Nội dung |
| --- | --- |
| `handover/NIKE_Financial_Dataset_FY2020-2026.xlsx` | Bộ dữ liệu chuẩn của nhóm, đồng thời là đầu vào của dashboard. |
| `pipeline/sources.py` | Danh sách 7 báo cáo 10-K (số accession, ngày nộp, đường dẫn trên SEC). |
| `pipeline/metrics.py` | Định nghĩa các chỉ tiêu, công thức tính chỉ số và các phép kiểm tra. Excel, Python và dashboard dùng chung một bộ công thức này. |
| `pipeline/ixbrl.py` | Đọc dữ liệu Inline XBRL trong 10-K: giá trị, số trang, mục. |
| `pipeline/extract.py` | Tải gói XBRL, trích số liệu ra `data/extracted/facts.csv` và đối chiếu với dữ liệu XBRL của SEC (`data/extracted/sec_api_check.csv`). |
| `pipeline/recheck_pages.py` | Mở trang 10-K bằng trình duyệt và đọc lại 35 số chỉ có trong MD&A: 28 tỷ lệ tăng trưởng loại trừ tỷ giá theo vị trí cột, 7 số Digital từ câu văn. Kết quả ghi vào `data/extracted/page_check.csv`. Ảnh chụp và trang bằng chứng nằm trong `handover/bang-chung-doi-chieu/` (cũng được đưa lên dashboard web). |
| `pipeline/build_workbook.py` | Tạo file Excel bộ dữ liệu, giữ lại các cột xác minh mà nhóm đã điền. |
| `pipeline/verify_workbook.py` | Tính lại mọi công thức Excel bằng một công cụ độc lập rồi so với kết quả Python. |
| `pipeline/build_site_data.py` | Kiểm tra file Excel, tính các chỉ số và xuất `site/data/nike-financials.json`. Dừng lại nếu có lỗi. |
| `pipeline/build_excel_dashboard.py`, `pipeline/excel_slicers.ps1` | Tạo bản Excel của dashboard (`handover/NIKE_Dashboard_FY2020-2026.xlsx`) từ cùng file JSON. Slicer chọn năm và slicer cơ sở tăng trưởng được thêm bằng Excel trên Windows. File mở bằng Excel trên máy hoặc Excel for the web (OneDrive), không mở bằng Google Sheets vì có slicer và hàm AGGREGATE. |
| `pipeline/build_powerbi.py`, `powerbi/` | Tạo dự án Power BI (PBIP) từ cùng file JSON. Bản giao cho nhóm là `handover/NIKE_Dashboard_FY2020-2026.pbix`, lưu ra từ Power BI Desktop. |
| `handover/GIAI-THICH-QUY-TRINH.md`, `.pdf` | Giải thích quy trình cho người trình bày và các thành viên, không cần đọc mã. Sau mỗi lần sửa file .md, tạo lại bản PDF bằng `python -m pipeline.build_guide_pdf`. |
| `site/` | Dashboard web tĩnh (HTML, CSS, JavaScript, thư viện biểu đồ Apache ECharts 5.6.0), được đưa lên GitHub Pages. |
| `tests/` | Kiểm thử dữ liệu (`test_pipeline.py`) và kiểm thử trên trình duyệt (`test_dashboard.py`). |
| `.github/workflows/pages.yml` | Mỗi lần cập nhật: kiểm tra dữ liệu, chạy toàn bộ kiểm thử, đạt hết thì mới đưa `site/` lên GitHub Pages. |

## Cài đặt

```bash
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt
.venv/bin/python -m playwright install chromium   # chỉ cần cho kiểm thử trên trình duyệt
```

Kiểm thử trên trình duyệt dùng Google Chrome nếu máy có cài, nếu không thì dùng Chromium của Playwright.

## Quy trình dữ liệu

```bash
.venv/bin/python -m pipeline.extract           # tải 10-K, trích số liệu, đối chiếu với dữ liệu SEC
.venv/bin/python -m pipeline.recheck_pages     # đọc lại 35 số MD&A trên trang 10-K (cần Chrome, có nghỉ giữa các báo cáo)
.venv/bin/python -m pipeline.build_workbook    # tạo handover/NIKE_Financial_Dataset_FY2020-2026.xlsx
.venv/bin/python -m pipeline.verify_workbook   # tính lại công thức Excel bằng công cụ độc lập
.venv/bin/python -m pipeline.build_site_data   # kiểm tra file Excel và xuất JSON cho dashboard
.venv/bin/python -m pipeline.build_excel_dashboard   # bản Excel của dashboard (slicer cần Windows và Excel)
.venv/bin/python -m pytest tests -q            # 59 kiểm thử: dữ liệu, 3 bản dashboard và trình duyệt
```

Trên Windows, thay `.venv/bin/python` bằng `.venv\Scripts\python.exe`.

Cập nhật bản Power BI sau khi dữ liệu thay đổi (sau bước `build_site_data`):

1. `.venv\Scripts\python.exe -m pipeline.build_powerbi`
2. Mở `powerbi\NIKE_Dashboard.pbip` bằng Power BI Desktop, bấm **Refresh**, chọn trang “Chiến lược phân phối”.
3. **File › Save as**, chọn kiểu “Power BI file (*.pbix)”, ghi đè `handover\NIKE_Dashboard_FY2020-2026.pbix`.

Ghi chú:

- File Excel là nguồn chuẩn. `build_site_data` đọc sheet 02 và 03, so từng giá trị và số trang với dữ liệu trích từ 10-K, rồi dừng lại (mã lỗi 1, không xuất JSON) nếu gặp một trong các lỗi sau: thiếu dòng, trùng khóa, sai đơn vị hoặc phạm vi, ô trống mà không ghi “Thiếu / không công bố”, phần trăm nhập dạng số nguyên, thiếu nguồn, giá trị hoặc số trang khác số trích từ 10-K, “Đã xác minh (người)” thiếu tên hoặc ngày, có “Bằng chứng mâu thuẫn”, hoặc phép kiểm tra vượt ngưỡng. Bước này giữ cho dashboard không bị sai khi file Excel bị sửa nhầm, nhưng không thay cho việc kiểm tra chính các số trích xuất (đã làm ở `extract` và `recheck_pages`).
- Khi nhóm đánh dấu “Đã xác minh (người)” trong file Excel trên Google Drive: tải file đó về, chép đè vào `handover/` rồi chạy lại `build_site_data`. Chạy lại `build_workbook` vẫn giữ nguyên các cột xác minh. Nếu số liệu mới khác số đã xác minh, dòng đó sẽ chuyển sang “Bằng chứng mâu thuẫn”.
- Gói XBRL được tải từ bản sao chính thức trên NIKE Investor Relations (cùng số accession với bản trên EDGAR), vì SEC yêu cầu chương trình tự động phải khai email liên hệ. Mọi số hợp nhất đều được đối chiếu với dữ liệu XBRL do SEC công bố (`data.sec.gov`). Trước khi chạy `extract`, hãy đặt biến môi trường `SEC_CONTACT_EMAIL`, hoặc mở đường dẫn API trên trình duyệt rồi lưu file vào `data/cache/companyfacts.json`. Trang NIKE IR sẽ tạm chặn (HTTP 403) nếu tải dồn dập, nên `recheck_pages` có nghỉ giữa các báo cáo.

### Thêm năm tài chính mới (ví dụ FY2027)

1. Thêm một dòng vào `FILINGS` và sửa `LAST_FY` trong `pipeline/sources.py`.
2. Chạy lại toàn bộ quy trình trên và đọc kỹ các dòng `NOTE:` của `extract`: chỗ nào NIKE đổi cách trình bày câu hoặc bảng sẽ hiện ở đây.
3. Cập nhật các con số cố định trong kiểm thử (`tests/test_pipeline.py`) nếu cần, rồi xem lại dashboard ở cả 3 độ rộng màn hình.

## Xem dashboard trên máy

```bash
.venv/bin/python -m pipeline.build_site_data
.venv/bin/python -m http.server 8000 --directory site
```

Mở `http://localhost:8000`. Nếu mở thẳng `site/index.html` từ ổ đĩa thì trang không tải được dữ liệu, vì trình duyệt chặn `fetch` từ `file://`.

## Đưa lên GitHub Pages

1. Tạo kho trên GitHub (ví dụ `nike-financial-dashboard`) rồi đẩy mã nguồn lên:
   ```bash
   git init -b main
   git add .
   git commit -m "NIKE FY2020–FY2026 dataset and dashboard"
   git remote add origin https://github.com/<tai-khoan>/nike-financial-dashboard.git
   git push -u origin main
   ```
2. Trên GitHub: **Settings → Pages → Build and deployment → Source: GitHub Actions**.
3. Workflow `Deploy dashboard` sẽ tự chạy: kiểm tra dữ liệu, chạy toàn bộ kiểm thử rồi đưa trang lên. Đường dẫn có dạng `https://<tai-khoan>.github.io/nike-financial-dashboard/`.
4. Đường dẫn dashboard được ghi vào file Excel và bản Power BI từ hằng số `DASHBOARD_URL` trong `pipeline/build_workbook.py` và `pipeline/build_powerbi.py`. Nếu đổi tài khoản hoặc tên kho, sửa hằng số này ở cả hai file rồi tạo lại các file.

Trang chỉ dùng đường dẫn tương đối, không cần máy chủ, khóa bí mật hay dịch vụ trả phí. Khi đưa lên, file Excel được chép vào `site/downloads/` và trang bằng chứng vào `site/bang-chung-doi-chieu/`.

## Ghi chú thiết kế

- **Mỗi trang trả lời một câu hỏi.** Dashboard có hai trang, ứng với hai câu hỏi của đề tài. Loại biểu đồ được chọn theo điều cần xem: thanh 100% cho cơ cấu kênh, cột nhóm cho quy mô doanh thu, cột âm dương cho tăng trưởng, cột lồng nhau cho phần Digital trong NIKE Direct, và các biểu đồ nhỏ cùng khung cho kết quả hợp nhất. Không dùng biểu đồ hai trục, và biểu đồ nào có trục tung không bắt đầu từ 0 đều ghi rõ.
- **Số nào cũng có nguồn.** Mỗi số NIKE công bố trong bảng đều kèm số trang 10-K có liên kết, số do nhóm tự tính được in nghiêng, số chưa được kiểm tra có dấu † (dấu này tự ẩn khi không còn số nào chờ kiểm tra), và mỗi biểu đồ có dòng nguồn. Ảnh PNG tải về đã có sẵn tiêu đề, chú giải và nguồn. Nút “Tải bảng (Excel)” xuất file .xlsx mở đúng trên mọi máy, kể cả máy cài tiếng Việt (CSV thì không). In trang (Ctrl+P) sẽ ra cả hai trang, khổ A4 ngang.
- **Màu cố định theo đối tượng.** Xanh dương luôn là NIKE Direct, cam luôn là bán buôn, xám xanh là số hợp nhất của NIKE, Inc. Bảng màu đã được kiểm tra để người mù màu vẫn phân biệt được, và chữ đủ tương phản để dễ đọc.
- **Phông chữ.** Source Sans 3 cho giao diện và các con số (chữ số rộng bằng nhau nên cột số thẳng hàng), Source Serif 4 cho tiêu đề và hai câu hỏi. Cả hai phông hỗ trợ đầy đủ tiếng Việt, dùng giấy phép SIL OFL 1.1 và được lưu sẵn trong `site/assets/fonts/`.
- **Giao diện đơn giản.** Không dùng logo hay màu nhận diện của NIKE, không có hiệu ứng trang trí. Không có nút phóng to, vì mỗi biểu đồ chỉ có 7 năm và bộ lọc năm là đủ.

## Giấy phép thành phần bên thứ ba

- Apache ECharts 5.6.0: Apache License 2.0 (`site/assets/vendor/echarts-LICENSE.txt`).
- Source Sans 3, Source Serif 4 (gói Fontsource 5.3.0): SIL Open Font License 1.1 (`site/assets/fonts/OFL-*.txt`).
- Số liệu tài chính: báo cáo công khai của NIKE, Inc. trên SEC EDGAR.
