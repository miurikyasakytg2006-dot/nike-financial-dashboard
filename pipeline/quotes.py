"""NIKE's own words on its direct-to-consumer strategy, quoted verbatim from the 10-Ks with the printed page.

Each entry names the report, the start of the passage and how many sentences to quote; the text and page are read
from the document itself (pipeline.extract writes data/extracted/quotes.csv), so a quote cannot drift from the source.
The Vietnamese line says what the passage says, without adding to it.
"""
QUOTES = [
    dict(id="shift", filing="10-K FY2019", start="In the current marketplace environment, we believe", sentences=2,
         topic="Vì sao",
         vi="NIKE cho rằng cách người tiêu dùng mua sắm vẫn đang thay đổi rõ rệt. Họ đòi hỏi sản phẩm mới liên tục, dịch "
            "vụ tốt, giao hàng nhanh, tất cả được thúc đẩy bởi xu hướng chuyển sang trải nghiệm Digital và trải nghiệm chỉ "
            "một thương hiệu (mono-brand) trong NIKE Direct."),
    dict(id="offense", filing="10-K FY2019", start="Through the Consumer Direct Offense", sentences=1, topic="Chiến lược",
         vi="Thông qua Consumer Direct Offense, NIKE tập trung vào chiến lược “Triple Double”, với mục tiêu nhân đôi tác "
            "động của đổi mới, đưa hàng ra thị trường nhanh hơn và tăng kết nối trực tiếp với người tiêu dùng."),
    dict(id="acceleration", filing="10-K FY2020", start="In June 2020, we announced", sentences=1, topic="Chiến lược",
         vi="Tháng 6/2020, NIKE công bố giai đoạn mới của chiến lược Consumer Direct Offense, dựa trên sức mạnh công nghệ "
            "số, gọi là Consumer Direct Acceleration."),
    dict(id="digital_first", filing="10-K FY2020", start="This next phase of our Consumer Direct Offense", sentences=1,
         topic="Chiến lược",
         vi="NIKE kỳ vọng giai đoạn này mang lại tăng trưởng và lợi nhuận bền vững, khi NIKE đẩy nhanh việc trở thành công "
            "ty ưu tiên kênh số (digital-first)."),
    dict(id="fueled", filing="10-K FY2021", start="Our NIKE Direct business fueled our growth", sentences=1, topic="Kết quả",
         vi="Theo NIKE, NIKE Direct là động lực tăng trưởng suốt FY2021 trong đại dịch, nhờ kết hợp nền tảng số với hệ "
            "thống cửa hàng để kết nối trực tiếp với người tiêu dùng."),
    dict(id="margin", filing="10-K FY2021", start="Additionally, the favorable impact of growth in our higher margin NIKE Direct",
         sentences=1, topic="Lợi nhuận",
         vi="NIKE Direct, dẫn đầu là Digital, có biên lợi nhuận cao hơn. Nhưng trong FY2021, phần lợi từ tăng trưởng của "
            "NIKE Direct bị lấn át bởi việc tăng khuyến mại ở cửa hàng outlet nửa đầu năm để giảm tồn kho dư do COVID-19."),
    dict(id="demand_creation", filing="10-K FY2021", start="Demand creation expense decreased 13% for fiscal 2021", sentences=1,
         topic="Chi phí",
         vi="Chi phí tạo nhu cầu FY2021 giảm 13%: chi marketing và quảng cáo cho sự kiện thương hiệu và hoạt động bán lẻ "
            "giảm, chi tài trợ thể thao cũng giảm vì các giải đấu bị hoãn do COVID-19."),
]
QUOTES_BY_ID = {q["id"]: q for q in QUOTES}
