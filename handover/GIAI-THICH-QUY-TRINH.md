# Giải thích quy trình: số liệu và dashboard NIKE FY2020–FY2026

Tài liệu này dành cho người trình bày: đọc để hiểu dữ liệu được làm ra thế nào và trả lời được khi bị hỏi. Phần kỹ thuật chi tiết (lệnh chạy, cấu trúc mã) nằm ở `README.md` của kho mã.

## 1. Tóm tắt

Toàn bộ số liệu lấy từ **7 báo cáo thường niên Form 10-K** mà NIKE, Inc. nộp cho Ủy ban Chứng khoán Mỹ (SEC), FY2020–FY2026. Một chương trình đọc số từ các báo cáo, kiểm tra chéo bằng nhiều cách rồi ghi vào **một file Excel chuẩn**. Từ file đó tạo ra **3 bản dashboard cùng nội dung**: web, Excel và Power BI.

## 2. Các file và dùng để làm gì

| File | Dùng để làm gì |
| --- | --- |
| `NIKE_Financial_Dataset_FY2020-2026.xlsx` | Bộ dữ liệu gốc. Mỗi dòng là một chỉ tiêu của một năm, kèm link báo cáo, số trang, câu gốc và trạng thái kiểm tra. Dùng khi cần trích số cho báo cáo. |
| Dashboard web: <a href="https://miurikyasakytg2006-dot.github.io/nike-financial-dashboard/">https://<span class="nowrap">miurikyasakytg2006-dot.github.io</span>/<wbr><span class="nowrap">nike-financial-dashboard/</span></a> | Xem nhanh trên trình duyệt, lọc năm, tải ảnh biểu đồ (PNG) để dán vào slide, tải bảng số dạng Excel, in cả hai góc nhìn (Ctrl+P). |
| `NIKE_Dashboard_FY2020-2026.xlsx` | Cùng dashboard, dạng Excel, lọc năm bằng slicer. Mở bằng Excel trên máy hoặc Excel for the web (OneDrive). **Không mở bằng Google Sheets** vì file dùng slicer và hàm riêng của Excel. |
| `NIKE_Dashboard_FY2020-2026.pbix` | Cùng dashboard, dạng Power BI. Mở bằng Power BI Desktop bản mới nhất (miễn phí, chỉ chạy trên Windows). Nên giữ một bản gốc, vì Power BI Desktop có thể tự ghi lại file khi đóng. |
| `bang-chung-doi-chieu/doi-chieu-35-so.html` | Bằng chứng cho 35 số chỉ có trong phần MD&A: bảng so sánh và ảnh chụp đúng chỗ trên trang 10-K. Cũng mở được từ dashboard web (mục “Nguồn và phương pháp”). |

Cả 3 bản dashboard được **tạo từ cùng một file dữ liệu**, nên số liệu giống hệt nhau. Khi dữ liệu thay đổi, cả 3 bản được tạo lại, và các bài kiểm thử tự động sẽ so dữ liệu của bản Excel và bản Power BI với bản web.

## 3. Quy trình 5 bước

**7 báo cáo 10-K → 1. Lấy số → 2. Kiểm tra chéo → 3. File Excel chuẩn → 4. Kiểm tra trước khi xuất → 5. Ba bản dashboard**

**Bước 1: Lấy số.** Báo cáo 10-K nộp SEC có gắn **thẻ XBRL**: mỗi con số trong báo cáo tài chính được gắn một “nhãn” mà máy đọc được (ví dụ nhãn “Doanh thu”). Chương trình đọc các nhãn này để lấy số của báo cáo kết quả kinh doanh, bảng cân đối và thuyết minh doanh thu theo kênh. Những số không có nhãn, nằm trong phần **MD&A** (phần ban điều hành phân tích kết quả, không được kiểm toán), thì đọc từ bảng và câu văn. Với mỗi số, chương trình ghi lại số trang, mục và nguyên văn câu hoặc dòng bảng.

**Bước 2: Kiểm tra chéo.** Mỗi số được kiểm tra lại bằng ít nhất một cách khác:

| Cách kiểm tra | Áp dụng cho | Kết quả |
| --- | --- | --- |
| So với dữ liệu XBRL do SEC công bố cho cùng báo cáo | Số hợp nhất của NIKE, Inc., kể cả số năm trước in kèm trong mỗi báo cáo | 247/247 khớp |
| So hai chỗ khác nhau trong cùng báo cáo: thuyết minh “Revenues” và bảng doanh thu trong MD&A | Doanh thu theo kênh | Khớp toàn bộ |
| Cộng trừ: doanh thu − giá vốn = lợi nhuận gộp; Direct + bán buôn + GBD = NIKE Brand; … (7 phép) | Mọi năm | Chênh lệch bằng 0 ở cả 7 phép và 7 năm (ngưỡng cho phép ±1–2 triệu USD) |
| So tỷ lệ nhóm tự tính với % NIKE công bố | Tăng trưởng, tỷ trọng | Đạt, chênh lệch tối đa 0,43 điểm % (NIKE làm tròn đến số nguyên) |
| Đọc lại trực tiếp trên trang 10-K, chụp ảnh đúng chỗ | 35 số chỉ có trong MD&A (tăng trưởng loại trừ tỷ giá, doanh số Digital) | 35/35 khớp |

Lưu ý khi trình bày cách kiểm tra đầu tiên: dữ liệu SEC công bố chính là dữ liệu XBRL NIKE nộp, nên phép so này xác nhận chương trình **đọc đúng nhãn và đúng kỳ**. Đây không phải một nguồn số liệu độc lập thứ hai.

Với 35 số cuối: **28 tỷ lệ tăng trưởng loại trừ tỷ giá** được đọc lại theo **vị trí cột** trên trang đã hiển thị, khác với cách đọc ở bước 1. Bảng doanh thu có hai cột cùng tên “% CHANGE EXCLUDING CURRENCY CHANGES” (cột thứ hai là của năm trước), nên đây là chỗ dễ lấy nhầm nhất. Lần đọc lại xác nhận đã lấy đúng cột đầu tiên. **7 số doanh số Digital** được đọc lại từ câu văn trên đúng trang đó.

**Bước 3: File Excel chuẩn.** Mọi số được ghi vào file Excel, mỗi dòng một chỉ tiêu của một năm, kèm link mở báo cáo trên SEC và số trang. Sheet 04 tính các chỉ số bằng **công thức Excel thật**, nên ai cũng xem được cách tính. Sheet 05 ghi thuật ngữ, các số liệu còn thiếu và các lần NIKE trình bày lại số liệu.

**Bước 4: Kiểm tra trước khi xuất.** Trước khi cập nhật dashboard, chương trình đọc lại file Excel và so với số đã trích từ 10-K. Nếu số hoặc số trang trong file khác số trích xuất, thiếu nguồn, sai đơn vị, ô trống không ghi lý do, hoặc phép cộng trừ vượt ngưỡng thì **dừng, không xuất**. Nhờ vậy, một chỗ sửa nhầm trong file Excel không thể lên dashboard. Bước này không thay cho việc kiểm tra các số trích xuất, việc đó đã làm ở bước 2. Ngoài ra còn 59 bài kiểm thử tự động chạy lại mỗi lần cập nhật. Các bài này kiểm tra số trên dashboard web có đúng bằng số trong file không, và dữ liệu của bản Excel, Power BI có khớp với bản web không.

**Bước 5: Tạo 3 bản dashboard.** Bản web được đưa lên GitHub Pages sau khi qua toàn bộ kiểm thử. Bản Excel và Power BI được tạo từ cùng file dữ liệu.

## 4. Kết quả: 203 số liệu

203 số = 29 chỉ tiêu × 7 năm.

- **202 số “Khớp nguồn chính thức”.** 167 số được xác nhận bằng 4 cách đầu ở bảng trên, còn 35 số được xác nhận bằng cách đọc lại trên trang 10-K có ảnh chụp.
- **1 số “Không công bố”:** tỷ trọng NIKE Direct do NIKE tự nêu cho FY2026. Báo cáo FY2026 không có câu này. Tỷ trọng nhóm tự tính vẫn có đủ 7 năm (FY2026: 39,2%).
- **“Đã xác minh (người)” = 0.** Trạng thái này chỉ dành cho khi một thành viên tự mở báo cáo và xác nhận. Nếu muốn, bạn mở trang bằng chứng, xem ảnh, rồi đánh dấu trong file Excel dữ liệu (phải điền tên và ngày).

## 5. Các chỉ số và cách tính

| Chỉ số | Công thức | Ví dụ FY2026 |
| --- | --- | --- |
| Tỷ trọng NIKE Direct | Doanh thu NIKE Direct / Doanh thu NIKE Brand | 17.720 / 45.222 = 39,2% |
| Tăng trưởng (theo báo cáo) | Số năm nay / số năm trước trong cùng báo cáo − 1 | Direct: 17.720 / 18.783 − 1 = −5,7% |
| Tăng trưởng loại trừ tỷ giá | NIKE tự tính theo tỷ giá năm trước, lấy nguyên văn (non-GAAP) | Direct: −8% |
| Biên lợi nhuận gộp | Lợi nhuận gộp / Doanh thu | 42,9% |
| Chi phí bán hàng và quản lý / doanh thu | (Tạo nhu cầu + Vận hành chung) / Doanh thu | 34,7% |
| EBIT | Lợi nhuận sau thuế + Chi phí lãi thuần + Thuế | 3.108 + (−50) + 792 = 3.850 triệu USD |
| Biên EBIT | EBIT / Doanh thu | 8,3% |
| Số ngày tồn kho | 365 × Tồn kho bình quân đầu và cuối kỳ / Giá vốn hàng bán | 103,3 ngày |
| Digital trong NIKE Direct | Doanh số Digital / Doanh thu NIKE Direct (xấp xỉ) | 8,6 tỷ / 17,72 tỷ ≈ 49% |

## 6. Những lựa chọn quan trọng và lý do

- **Năm tài chính của NIKE kết thúc ngày 31/5.** FY2026 là 01/06/2025 – 31/05/2026, không trùng năm dương lịch.
- **Tỷ trọng kênh dùng mẫu số NIKE Brand, không phải NIKE, Inc.** Converse có kênh trực tiếp riêng và không thuộc NIKE Direct. NIKE cũng tự nêu tỷ trọng trên cơ sở NIKE Brand.
- **Dùng EBIT thay cho “lợi nhuận hoạt động”.** Báo cáo kết quả kinh doanh của NIKE không có dòng tổng phụ operating income: sau chi phí bán hàng và quản lý là chi phí (thu nhập) lãi thuần và chi phí (thu nhập) khác thuần, rồi đến lợi nhuận trước thuế. EBIT là chỉ tiêu non-GAAP do NIKE tự định nghĩa. EBIT vẫn gồm khoản chi phí (thu nhập) khác thuần, nên không hoàn toàn giống “lợi nhuận hoạt động” theo cách hiểu thông thường. EBIT nhóm tự tính khớp đúng EBIT NIKE công bố.
- **Không có lợi nhuận theo kênh.** NIKE không công bố lợi nhuận hay chi phí riêng của NIKE Direct và bán buôn. Vì vậy dashboard không so sánh khả năng sinh lời giữa các kênh.
- **Hai loại tăng trưởng.** Tăng trưởng “theo báo cáo” do nhóm tự tính từ số liệu, còn tăng trưởng “loại trừ tỷ giá” do NIKE tính và nhóm không tự tính lại được.
- **Số năm trước lấy trong cùng báo cáo.** Mỗi 10-K đều in kèm số của năm trước để so sánh, ví dụ 10-K FY2026 in doanh thu FY2026 là 46.398 và FY2025 là 46.309 triệu USD. Tăng trưởng được tính với chính số năm trước này (46.398 / 46.309 − 1 = +0,2%), giống cách NIKE tính. Nếu báo cáo năm sau trình bày lại thì chỉ ghi chú, không thay số gốc. Ví dụ duy nhất: doanh số Digital FY2023 được trình bày lại từ 12,6 xuống 12,4 tỷ USD. Khi so FY2024 (12,1 tỷ USD) với năm trước, cần dùng số 12,4 để cùng cách phân loại.
- **Số liệu cho thấy các thay đổi xảy ra cùng giai đoạn, không chứng minh nhân quả.** Phần giải thích vì sao thuộc về phân tích của nhóm, có dẫn nguồn.
- **Định dạng số.** Bảng và câu tóm tắt ở cả 3 bản đều viết kiểu Việt Nam (12.382; 34,8%). Riêng nhãn trên biểu đồ của bản Excel hiển thị dấu thập phân theo cài đặt ngôn ngữ Excel của người xem (34,8% hoặc 34.8%).

## 7. Câu hỏi có thể gặp và gợi ý trả lời

**Số liệu lấy ở đâu?**
Từ 7 báo cáo 10-K của NIKE trên SEC EDGAR. Mỗi số trong file Excel có link mở đúng báo cáo và số trang để kiểm tra.

**Làm sao biết số đúng?**
Mỗi số được kiểm tra lại bằng ít nhất một cách khác (bảng ở mục 3). Riêng 35 số khó kiểm tra nhất có ảnh chụp đúng chỗ trên báo cáo, xem được ngay trên dashboard web. Trước khi cập nhật dashboard, chương trình còn chặn mọi chỗ trong file Excel lệch với số trích từ báo cáo.

**Sao không lấy từ Yahoo Finance, Statista…?**
Các trang đó thường tự tính lại hoặc gộp số theo cách riêng, ví dụ tự tạo chỉ tiêu “operating income” mà NIKE không công bố. Báo cáo gốc là nguồn chính xác nhất và trích dẫn được.

**Tỷ trọng NIKE Direct FY2026 là bao nhiêu, tính thế nào?**
39,2% = 17.720 / 45.222 triệu USD, trên doanh thu NIKE Brand.

**Sao có hai con số tăng trưởng khác nhau?**
Một con số tính theo USD như trong báo cáo, con số kia đã được NIKE loại bỏ ảnh hưởng của tỷ giá. Ví dụ NIKE Direct FY2026: −5,7% theo báo cáo, −8% loại trừ tỷ giá.

**Cách làm này liên quan gì đến chủ đề “đổi mới” của môn?**
XBRL là chuẩn báo cáo tài chính điện tử: số trong báo cáo được gắn nhãn để máy đọc được. Cách “máy lấy số, đối chiếu nhiều lớp, có bằng chứng” đang được dùng trong kiểm toán hiện nay. Mọi con số ở đây đều truy được về một trang cụ thể trong báo cáo gốc, ai cũng tự kiểm tra lại được.

**Mở các file bằng gì?**
Dashboard web: trình duyệt bất kỳ. File Excel dashboard: Excel trên máy hoặc Excel for the web, không dùng Google Sheets. File Power BI: Power BI Desktop bản mới nhất trên Windows.

**Khi có báo cáo FY2027 thì cập nhật thế nào?**
Thêm một dòng vào danh sách báo cáo, rồi chạy lại quy trình. Các bước kiểm tra sẽ báo nếu NIKE đổi cách trình bày.

**Vì sao có 3 phiên bản dashboard?**
Bản web để xem nhanh và dán biểu đồ vào slide. Bản Excel cho người quen Excel. Bản Power BI là công cụ phân tích phổ biến trong doanh nghiệp. Cả ba được tạo từ cùng một dữ liệu.
