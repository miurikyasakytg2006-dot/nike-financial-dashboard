# NIKE FY2019–FY2021: dữ liệu tài chính và dashboard

Mã nguồn phần số liệu của đề tài về chiến lược bán trực tiếp (NIKE Direct) của NIKE, giai đoạn FY2019–FY2021. Nếu chỉ cần dùng số liệu, bạn không cần đọc tài liệu này: các file dùng ngay nằm trong thư mục `handover/`.

Toàn bộ số liệu lấy từ báo cáo Form 10-K (năm), Form 10-Q (quý) và thông cáo kết quả quý 4 (Form 8-K) mà NIKE, Inc. nộp cho SEC. Năm tài chính của NIKE kết thúc ngày 31/5, nên FY2019–FY2021 là từ 01/06/2018 đến 31/05/2021. Đây là sản phẩm học tập của sinh viên, không liên quan đến NIKE, Inc.

## Cấu trúc

| Đường dẫn | Nội dung |
| --- | --- |
| `handover/NIKE_Financial_Dataset_FY2019-2021.xlsx` | Bộ dữ liệu chuẩn của nhóm, đồng thời là đầu vào của dashboard: số liệu năm, số liệu 12 quý, bốn khu vực, các chỉ số tự tính (đều là công thức Excel) và nguồn của từng số. |
| `handover/NIKE_Dashboard_FY2019-2021.xlsx`, `.pbix` | Bản Excel và bản Power BI của dashboard. |
| `handover/bang-chung-doi-chieu/` | Ảnh chụp trang 10-Q và trang “Đối chiếu 36 số” (xem `recheck_pages` ở dưới). |
| `pipeline/sources.py` | 20 tài liệu dùng trong dự án: 3 báo cáo 10-K và 9 báo cáo 10-Q của FY2019–FY2021 là nguồn số liệu; 3 thông cáo kết quả quý 4 (Form 8-K) cho tỷ lệ % của quý 4 và để đối chiếu số quý 4 tự tính; 5 báo cáo của FY2022–FY2023 chỉ dùng để xem NIKE có in lại số cũ khác đi không. |
| `pipeline/metrics.py` | Định nghĩa các chỉ tiêu, công thức tính chỉ số và các phép kiểm tra. Excel, Python và dashboard dùng chung một bộ công thức này. |
| `pipeline/ixbrl.py` | Đọc báo cáo: dữ liệu Inline XBRL (từ FY2020) hoặc tệp XBRL riêng đi kèm bản HTML (các báo cáo FY2019), cùng số trang và mục của từng số. |
| `pipeline/extract.py` | Tải báo cáo, trích số liệu ra `data/extracted/facts.csv` và các đoạn trích dẫn ra `data/extracted/quotes.csv`. Ngoài bảng doanh thu, đọc cả những chỗ NIKE in lại cùng tỷ lệ trong một báo cáo (bảng doanh thu theo khu vực, câu về NIKE Direct trong MD&A). Với thông cáo quý 4: đọc báo cáo kết quả kinh doanh 3 tháng, bảng doanh thu theo khu vực, bảng EBIT và các câu nêu % thay đổi của quý. |
| `pipeline/quotes.py` | Bảy đoạn NIKE tự giải thích chiến lược trong 10-K, kèm ý chính bằng tiếng Việt. |
| `pipeline/dataset.py` | Chọn số cho từng kỳ, đối chiếu chéo giữa các nguồn và tính các chỉ số. |
| `pipeline/recheck_pages.py` | Mở trang 10-Q bằng Chrome và đọc lại 36 tỷ lệ tăng trưởng loại trừ tỷ giá của các quý (không có thẻ XBRL, báo cáo năm sau không in lại; riêng tỷ lệ bán buôn chỉ in một lần trong mỗi 10-Q). Kết quả ghi vào `data/extracted/page_check.csv`, ảnh chụp lưu trong `handover/bang-chung-doi-chieu/`. |
| `pipeline/build_workbook.py` | Tạo file Excel bộ dữ liệu, giữ lại các cột xác minh mà nhóm đã điền. |
| `pipeline/verify_workbook.py` | Tính lại mọi công thức Excel bằng một công cụ độc lập rồi so với kết quả Python. |
| `pipeline/insights.py` | Sáu phát hiện chính và các câu tóm tắt, viết một lần từ số liệu. Cả ba bản dashboard hiện đúng các câu này. |
| `pipeline/build_site_data.py` | Kiểm tra file Excel, tính các chỉ số và xuất `site/data/nike-financials.json`. Dừng lại nếu có lỗi. |
| `pipeline/build_excel_dashboard.py`, `pipeline/excel_save.ps1` | Tạo bản Excel của dashboard từ cùng file JSON. Trên Windows, file được Excel mở và lưu lại một lần để Excel for the web (OneDrive) hiện được biểu đồ. |
| `pipeline/build_powerbi.py`, `powerbi/` | Tạo dự án Power BI (PBIP) từ cùng file JSON. Bản giao cho nhóm là file .pbix, lưu ra từ Power BI Desktop. |
| `site/` | Dashboard web tĩnh (HTML, CSS, JavaScript, thư viện biểu đồ Apache ECharts 5.6.0), được đưa lên GitHub Pages. |
| `tests/` | Kiểm thử dữ liệu và hai bản dashboard Excel, Power BI (`test_pipeline.py`), kiểm thử trên trình duyệt (`test_dashboard.py`). |
| `.github/workflows/pages.yml` | Mỗi lần cập nhật: kiểm tra dữ liệu, chạy toàn bộ kiểm thử, đạt hết thì mới đưa `site/` lên GitHub Pages. |

## Số của mỗi kỳ lấy từ đâu

- **Năm:** từ 10-K của chính năm đó. **Quý 1–3:** từ 10-Q của chính quý đó (cột 3 tháng).
- **Quý 4:** NIKE không nộp 10-Q cho quý 4, nên số quý 4 bằng số cả năm (10-K) trừ số 9 tháng (10-Q quý 3). Số dư hàng tồn kho cuối quý 4 là số dư ngày 31/5 trong 10-K. Thông cáo kết quả quý 4 (8-K) của NIKE in số riêng của quý 4 (báo cáo kết quả kinh doanh, doanh thu theo khu vực, EBIT): số nào có trong đó đều khớp với số tự tính; thông cáo không có NIKE Direct, bán buôn và Digital. Tỷ lệ % của quý 4 lấy từ thông cáo này; tỷ lệ nào NIKE không nêu ở đâu cả (phần lớn tỷ lệ của NIKE Direct, bán buôn, tỷ trọng NIKE Direct) thì ghi “Không công bố”.
- **Tăng trưởng:** so với số cùng kỳ năm trước in trong cùng báo cáo, giống cách NIKE tự tính. Riêng tồn kho cuối quý 1–3: 10-Q chỉ in kèm số dư ngày 31/5, nên số cùng kỳ lấy từ 10-Q cùng quý năm trước. Không tính tăng trưởng Digital, vì số Digital đã làm tròn đến 0,1 tỷ USD làm tăng trưởng lệch với tỷ lệ NIKE nêu (FY2020: 44,7% so với 47%).
- **Đối chiếu:** mỗi số được kiểm tra thêm ít nhất một lần: so với chỗ khác NIKE in cùng số (bảng trong MD&A, câu trong thuyết minh, bảng doanh thu theo khu vực, bảng số liệu quý của 10-K FY2019 và FY2020, thông cáo kết quả quý 4, số in lại ở báo cáo năm sau), so với số lũy kế 6 và 9 tháng, tính lại từ báo cáo tài chính, hoặc đọc lại trên trang. Kết quả hiện tại: 627 số liệu, 613 khớp, 14 là tỷ lệ % NIKE không nêu cho riêng quý 4. Không có số nào bị NIKE in lại khác đi ở các báo cáo sau.

## Cài đặt

```bash
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt
.venv/bin/python -m playwright install chromium   # chỉ cần cho kiểm thử trên trình duyệt
```

Kiểm thử trên trình duyệt dùng Google Chrome nếu máy có cài, nếu không thì dùng Chromium của Playwright.

## Quy trình dữ liệu

```bash
.venv/bin/python -m pipeline.extract           # tải báo cáo (lưu tạm trong data/cache/), trích số liệu và trích dẫn
.venv/bin/python -m pipeline.recheck_pages     # đọc lại 36 tỷ lệ trên trang 10-Q (cần Chrome)
.venv/bin/python -m pipeline.build_workbook    # tạo handover/NIKE_Financial_Dataset_FY2019-2021.xlsx
.venv/bin/python -m pipeline.verify_workbook   # tính lại công thức Excel bằng công cụ độc lập
.venv/bin/python -m pipeline.build_site_data   # kiểm tra file Excel và xuất JSON cho dashboard
.venv/bin/python -m pipeline.build_excel_dashboard   # tạo handover/NIKE_Dashboard_FY2019-2021.xlsx
.venv/bin/python -m pytest tests -q            # 81 kiểm thử: dữ liệu, 3 bản dashboard và trình duyệt
```

Trên Windows, thay `.venv/bin/python` bằng `.venv\Scripts\python.exe`. `recheck_pages --from-saved` chỉ đánh giá lại các lần đọc đã lưu, không mở lại trang.

Cập nhật bản Power BI sau khi dữ liệu thay đổi (sau bước `build_site_data`):

1. `.venv\Scripts\python.exe -m pipeline.build_powerbi`
2. Mở `powerbi\NIKE_Dashboard.pbip` bằng Power BI Desktop, bấm **Refresh**.
3. **File › Save as**, chọn kiểu “Power BI file (*.pbix)”, ghi đè `handover\NIKE_Dashboard_FY2019-2021.pbix`.

Ghi chú:

- File Excel là nguồn chuẩn. `build_site_data` đọc sheet 02, 03 và 04, so từng giá trị và số trang với dữ liệu trích từ báo cáo, rồi dừng lại (mã lỗi 1, không xuất JSON) nếu gặp một trong các lỗi sau: thiếu dòng, trùng khóa, sai đơn vị hoặc phạm vi, ô trống mà không ghi “Thiếu / không công bố”, phần trăm nhập dạng số nguyên, thiếu nguồn, giá trị hoặc số trang khác số trích từ báo cáo, “Đã xác minh (người)” thiếu tên hoặc ngày, có “Bằng chứng mâu thuẫn”, hoặc phép kiểm tra vượt ngưỡng. Bước này giữ cho dashboard không bị sai khi file Excel bị sửa nhầm, nhưng không thay cho việc kiểm tra chính các số trích xuất (đã làm ở `extract`, `dataset` và `recheck_pages`).
- Các cột xác minh của nhóm (người xác minh, ngày, ghi chú) đang ẩn trong sheet 02–04. Khi nhóm đánh dấu “Đã xác minh (người)” trong file Excel trên Google Drive: tải file đó về, chép đè vào `handover/` rồi chạy lại `build_site_data`. Chạy lại `build_workbook` vẫn giữ nguyên các cột xác minh. Nếu số liệu mới khác số đã xác minh, dòng đó sẽ chuyển sang “Bằng chứng mâu thuẫn”.
- Báo cáo được tải từ bản sao chính thức trên trang NIKE Investor Relations (cùng số hồ sơ SEC, tức accession number, với bản trên SEC EDGAR); liên kết trong dashboard và bộ dữ liệu trỏ về SEC EDGAR. Trang NIKE IR có thể tạm chặn (HTTP 403) nếu tải dồn dập; khi đó chờ một lúc rồi chạy lại.

## Xem dashboard trên máy

```bash
.venv/bin/python -m pipeline.build_site_data
.venv/bin/python -m http.server 8000 --directory site
```

Mở `http://localhost:8000`. Nếu mở thẳng `site/index.html` từ ổ đĩa thì trang không tải được dữ liệu, vì trình duyệt chặn `fetch` từ `file://`.

## Đưa lên GitHub Pages

1. Trên GitHub: **Settings → Pages → Build and deployment → Source: GitHub Actions**.
2. Mỗi lần đẩy lên nhánh `main`, workflow `Deploy dashboard` tự chạy: kiểm tra dữ liệu, chạy toàn bộ kiểm thử rồi đưa trang lên. Đường dẫn có dạng `https://<tai-khoan>.github.io/nike-financial-dashboard/`.
3. Đường dẫn dashboard được ghi vào file Excel và bản Power BI từ hằng số `DASHBOARD_URL` trong `pipeline/build_workbook.py` và `pipeline/build_powerbi.py`. Nếu đổi tài khoản hoặc tên kho, sửa hằng số này ở cả hai file rồi tạo lại các file.

Trang chỉ dùng đường dẫn tương đối, không cần máy chủ, khóa bí mật hay dịch vụ trả phí. Khi đưa lên, file Excel được chép vào `site/downloads/` và trang bằng chứng vào `site/bang-chung-doi-chieu/`.

## Ghi chú thiết kế

- **Mỗi phần trả lời một câu hỏi.** Trang mở đầu bằng sáu phát hiện chính, rồi đến kênh phân phối, 12 quý, Digital và cửa hàng, bốn khu vực, chi phí và lợi nhuận, lời giải thích của NIKE và bảng số liệu. Loại biểu đồ chọn theo điều cần xem: thanh 100% cho cơ cấu kênh, thanh âm dương cho phần thay đổi, đường cho 12 quý, thanh chồng cho Digital trong NIKE Direct, biểu đồ điểm nối cho FY2019 so với FY2021 theo khu vực, và biểu đồ thác nước cho biên EBIT. Không dùng biểu đồ hai trục, và biểu đồ nào có trục tung không bắt đầu từ 0 đều ghi rõ.
- **Quý 4 FY2020 (tháng 3–5/2020) được tô xám** trên các biểu đồ theo quý trong bản web, vì đây là quý doanh thu giảm mạnh nhất trong 12 quý.
- **Số nào cũng có nguồn.** Mỗi số NIKE công bố trong bảng đều kèm số trang có liên kết đến báo cáo trên SEC, số do nhóm tự tính được in nghiêng, số ước tính từ số đã làm tròn có dấu ≈, và mỗi biểu đồ có dòng nguồn. Ảnh PNG tải về đã có sẵn tiêu đề, chú giải và nguồn. Nút “Tải bảng (Excel)” xuất file .xlsx mở đúng trên mọi máy, kể cả máy cài tiếng Việt (CSV thì không).
- **Màu cố định theo đối tượng.** Xanh dương luôn là NIKE Direct, cam luôn là bán buôn, xám xanh là số của cả NIKE, Inc. Bảng màu đã được kiểm tra để người mù màu vẫn phân biệt được, và chữ đủ tương phản để dễ đọc.
- **Phông chữ.** Source Sans 3 cho giao diện và các con số (chữ số rộng bằng nhau nên cột số thẳng hàng), Source Serif 4 cho tiêu đề. Cả hai phông hỗ trợ đầy đủ tiếng Việt, dùng giấy phép SIL OFL 1.1 và được lưu sẵn trong `site/assets/fonts/`.
- **Giao diện đơn giản.** Không dùng logo hay màu nhận diện của NIKE, không có hiệu ứng trang trí.

## Giấy phép thành phần bên thứ ba

- Apache ECharts 5.6.0: Apache License 2.0 (`site/assets/vendor/echarts-LICENSE.txt`).
- Source Sans 3, Source Serif 4 (gói Fontsource 5.3.0): SIL Open Font License 1.1 (`site/assets/fonts/OFL-*.txt`).
- Số liệu tài chính: báo cáo công khai của NIKE, Inc. trên SEC EDGAR.
