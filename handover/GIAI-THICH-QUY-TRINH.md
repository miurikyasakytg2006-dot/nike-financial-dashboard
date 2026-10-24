# Giải thích quy trình: số liệu và dashboard NIKE FY2020–FY2026

Tài liệu này dành cho người trình bày: đọc để hiểu dữ liệu được làm ra thế nào và trả lời được khi bị hỏi. Phần kỹ thuật chi tiết (lệnh chạy, cấu trúc mã) nằm ở `README.md` của kho mã.

## 1. Tóm tắt trong một câu

Toàn bộ số liệu lấy từ **7 báo cáo thường niên Form 10-K** của NIKE, Inc. nộp Ủy ban Chứng khoán Mỹ (SEC), FY2020–FY2026. Một chương trình đọc số từ báo cáo, kiểm tra chéo nhiều cách, ghi vào **một file Excel chuẩn**. Từ file đó xuất ra **3 bản dashboard giống nhau**: web, Excel web và Power BI.

## 2. Các file và dùng để làm gì

| File | Dùng để làm gì |
| --- | --- |
| `NIKE_Financial_Dataset_FY2020-2026.xlsx` | Bộ dữ liệu gốc. Mỗi dòng là một chỉ tiêu của một năm, kèm link báo cáo, số trang, câu gốc và trạng thái kiểm tra. Dùng khi cần trích số cho báo cáo. |
| Dashboard web: https://miurikyasakytg2006-dot.github.io/nike-financial-dashboard/ | Xem nhanh trên trình duyệt, lọc năm, tải ảnh biểu đồ (PNG) để dán vào slide. |
| `NIKE_Dashboard_FY2020-2026.xlsx` | Cùng dashboard, dạng Excel, mở bằng Excel for the web trên OneDrive. Lọc năm bằng slicer. |
| `NIKE_Dashboard_FY2020-2026.pbix` | Cùng dashboard, dạng Power BI. Mở bằng Power BI Desktop (bản mới nhất, miễn phí trên Microsoft Store). |
| `bang-chung-doi-chieu/doi-chieu-35-so.html` | Bằng chứng đối chiếu 35 số chỉ có trong phần thuyết minh: bảng so sánh và ảnh chụp đúng ô trên trang 10-K. |

Cả 3 bản dashboard đọc **cùng một file dữ liệu**, nên số liệu giống hệt nhau.

## 3. Quy trình 5 bước

```
 7 báo cáo 10-K ──► 1. Lấy số ──► 2. Kiểm tra chéo ──► 3. File Excel chuẩn ──► 4. Kiểm tra trước khi xuất ──► 5. Ba bản dashboard
```

**Bước 1: Lấy số.** Báo cáo 10-K nộp SEC có gắn **thẻ XBRL**: mỗi con số trong báo cáo tài chính được gắn một "nhãn" mà máy đọc được (ví dụ nhãn "Doanh thu"). Chương trình đọc các nhãn này để lấy số của báo cáo kết quả kinh doanh, bảng cân đối và thuyết minh doanh thu theo kênh. Những số không có nhãn, nằm trong phần **MD&A** (phần ban điều hành giải thích kết quả), thì đọc từ bảng và câu văn. Với mỗi số, chương trình ghi lại số trang, mục và nguyên văn câu hoặc dòng bảng.

**Bước 2: Kiểm tra chéo.** Mỗi số được xác nhận bằng ít nhất một cách độc lập:

| Cách kiểm tra | Áp dụng cho | Kết quả |
| --- | --- | --- |
| So với dữ liệu XBRL do chính SEC công bố (cùng báo cáo) | Số hợp nhất của NIKE, Inc. | 248/248 khớp |
| So hai chỗ khác nhau trong cùng báo cáo (thuyết minh doanh thu và bảng trong MD&A) | Doanh thu theo kênh | Khớp toàn bộ |
| Cộng trừ: doanh thu − giá vốn = lợi nhuận gộp; Direct + bán buôn + GBD = NIKE Brand; … (7 phép) | Mọi năm | Đạt, chênh lệch tối đa 1–2 triệu USD do làm tròn |
| So tỷ lệ tự tính với % Nike tự công bố | Tăng trưởng, tỷ trọng | Đạt, chênh lệch tối đa 0,43 điểm % (Nike làm tròn số nguyên) |
| Đọc lại trực tiếp trên trang 10-K bằng cách thứ hai, chụp ảnh đúng ô | 35 số chỉ có trong MD&A (tăng trưởng loại trừ tỷ giá, doanh số Digital) | 35/35 khớp |

Với 35 số cuối, cách thứ hai đọc số theo **vị trí cột** trên trang đã hiển thị, khác với cách đọc ở bước 1. Bảng doanh thu có hai cột cùng tên "% CHANGE EXCLUDING CURRENCY CHANGES" (cột thứ hai là của năm trước), nên đây là chỗ dễ lấy nhầm nhất; lần đọc lại xác nhận đã lấy đúng cột đầu tiên.

**Bước 3: File Excel chuẩn.** Mọi số được ghi vào file Excel, mỗi dòng một chỉ tiêu của một năm, kèm link mở báo cáo trên SEC và số trang. Sheet 04 tính các chỉ số bằng **công thức Excel thật**, nên ai cũng xem được cách tính. Sheet 05 ghi thuật ngữ, khoảng trống dữ liệu và các lần Nike trình bày lại số liệu.

**Bước 4: Kiểm tra trước khi xuất.** Trước khi cập nhật dashboard, chương trình đọc lại file Excel. Nếu có bất kỳ lỗi nào thì **dừng, không xuất**: thiếu dòng, sai đơn vị, ô trống không ghi lý do, số khác báo cáo, thiếu nguồn, hoặc phép kiểm tra cộng trừ vượt ngưỡng. Nhờ vậy số sai không thể lên dashboard. Ngoài ra có 52 bài kiểm thử tự động, kể cả kiểm tra rằng số hiển thị trên dashboard đúng bằng số trong file.

**Bước 5: Xuất 3 bản dashboard.** Bản web đưa lên GitHub Pages và tự cập nhật mỗi khi mã được cập nhật. Bản Excel web và Power BI được tạo từ cùng file dữ liệu.

## 4. Kết quả: 203 số liệu

203 số = 29 chỉ tiêu × 7 năm.

- **202 số "Khớp nguồn chính thức".** 167 số được xác nhận bằng đối chiếu tự động (4 cách đầu ở bảng trên), 35 số bằng cách đọc lại trên trang 10-K có ảnh chụp.
- **1 số "Không công bố":** tỷ trọng NIKE Direct do Nike tự nêu cho FY2026. Báo cáo FY2026 không có câu này. Tỷ trọng tự tính vẫn có đủ 7 năm (FY2026: 39,2%).
- **"Đã xác minh (người)" = 0.** Trạng thái này chỉ dành cho khi một thành viên tự mở báo cáo và xác nhận. Nếu muốn, bạn mở trang bằng chứng, xem ảnh, rồi đánh dấu trong file Excel dữ liệu.

## 5. Các chỉ số và cách tính

| Chỉ số | Công thức | Ví dụ FY2026 |
| --- | --- | --- |
| Tỷ trọng NIKE Direct | Doanh thu NIKE Direct / Doanh thu NIKE Brand | 17.720 / 45.222 = 39,2% |
| Tăng trưởng (theo báo cáo) | Số năm nay / số năm trước trong cùng báo cáo − 1 | Direct: 17.720 / 18.783 − 1 = −5,7% |
| Tăng trưởng loại trừ tỷ giá | Nike tự tính theo tỷ giá năm trước, lấy nguyên văn (non-GAAP) | Direct: −8% |
| Biên lợi nhuận gộp | Lợi nhuận gộp / Doanh thu | 42,9% |
| Chi phí bán hàng và quản lý / doanh thu | (Tạo nhu cầu + Vận hành chung) / Doanh thu | 34,7% |
| EBIT | Lợi nhuận sau thuế + Chi phí lãi thuần + Thuế | 3.108 + (−50) + 792 = 3.850 triệu USD |
| Biên EBIT | EBIT / Doanh thu | 8,3% |
| Số ngày tồn kho | 365 / (Giá vốn / Tồn kho bình quân đầu và cuối kỳ) | 103,3 ngày |
| Digital trong NIKE Direct | Doanh số Digital / Doanh thu NIKE Direct (xấp xỉ) | 8,6 tỷ / 17,72 tỷ ≈ 49% |

## 6. Những lựa chọn quan trọng và lý do

- **Năm tài chính của Nike kết thúc ngày 31/5.** FY2026 là 01/06/2025 – 31/05/2026, không trùng năm dương lịch.
- **Tỷ trọng kênh dùng mẫu số NIKE Brand, không phải NIKE, Inc.** Converse có kênh trực tiếp riêng và không thuộc NIKE Direct. Nike cũng tự nêu tỷ trọng trên cơ sở NIKE Brand.
- **Dùng EBIT thay cho "lợi nhuận hoạt động".** Báo cáo của Nike không có dòng operating income. EBIT là thước đo Nike tự dùng và ghi rõ là non-GAAP. EBIT tự tính được đối chiếu với EBIT Nike công bố.
- **Không có lợi nhuận theo kênh.** Nike không công bố lợi nhuận hay chi phí riêng của NIKE Direct và bán buôn. Vì vậy bộ dữ liệu không so sánh khả năng sinh lời giữa các kênh.
- **Hai loại tăng trưởng.** "Theo báo cáo" tự tính từ số liệu; "loại trừ tỷ giá" do Nike tính và không tự tính lại được.
- **Số năm trước lấy trong cùng báo cáo.** Đây cũng là cách Nike tính tăng trưởng. Nếu báo cáo năm sau trình bày lại thì chỉ ghi chú, không thay số gốc. Ví dụ duy nhất: doanh số Digital FY2023 được trình bày lại từ 12,6 xuống 12,4 tỷ USD.
- **Số liệu cho thấy các thay đổi xảy ra cùng giai đoạn, không chứng minh nhân quả.** Phần giải thích vì sao thuộc về phân tích của nhóm, có dẫn nguồn.

## 7. Câu hỏi có thể gặp và gợi ý trả lời

**Số liệu lấy ở đâu?**
Từ 7 báo cáo 10-K của Nike trên SEC EDGAR. Mỗi số trong file Excel có link mở đúng báo cáo và số trang để kiểm tra.

**Làm sao biết số đúng?**
Mỗi số được xác nhận bằng ít nhất một cách độc lập (xem bảng ở mục 3). Ngoài ra, chương trình dừng nếu phát hiện lỗi. Với 35 số khó kiểm tra nhất có ảnh chụp đúng ô trên báo cáo.

**Sao không lấy từ Yahoo Finance, Statista…?**
Các trang đó thường tự tính lại hoặc gộp số theo cách riêng, ví dụ tự tạo chỉ tiêu "operating income" mà Nike không công bố. Báo cáo gốc là nguồn chính xác nhất và trích dẫn được.

**Tỷ trọng NIKE Direct FY2026 là bao nhiêu, tính thế nào?**
39,2% = 17.720 / 45.222 triệu USD, trên doanh thu NIKE Brand.

**Sao có hai con số tăng trưởng khác nhau?**
Một con số tính theo USD báo cáo. Con số kia Nike loại bỏ ảnh hưởng của tỷ giá. Ví dụ NIKE Direct FY2026: −5,7% theo báo cáo, −8% loại trừ tỷ giá.

**Cách làm này liên quan gì đến chủ đề "đổi mới" của môn?**
XBRL là chuẩn báo cáo tài chính điện tử: số trong báo cáo được gắn nhãn để máy đọc được. Cách "máy lấy số, đối chiếu nhiều lớp, có bằng chứng" đang được dùng trong kiểm toán hiện nay. Mọi con số ở đây đều truy được về một trang cụ thể trong báo cáo gốc, ai cũng tự kiểm tra lại được.

**Khi có báo cáo FY2027 thì cập nhật thế nào?**
Thêm một dòng vào danh sách báo cáo, rồi chạy lại quy trình. Các bước kiểm tra sẽ báo nếu Nike đổi cách trình bày.

**Vì sao có 3 phiên bản dashboard?**
Bản web để xem nhanh và dán biểu đồ vào slide. Bản Excel cho người quen Excel. Bản Power BI là công cụ phân tích phổ biến trong doanh nghiệp. Cả ba dùng chung một dữ liệu.
