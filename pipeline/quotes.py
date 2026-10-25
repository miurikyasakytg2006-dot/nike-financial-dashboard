"""NIKE's own words on its direct-to-consumer strategy, quoted verbatim from the 10-Ks with the printed page.

Each entry names the report, the start of the passage and how many sentences to quote; the text and page are read
from the document itself (pipeline.extract writes data/extracted/quotes.csv), so a quote cannot drift from the source.
The Vietnamese line says what the passage says, without adding to it.
"""
QUOTES = [
    dict(id="shift", filing="10-K FY2019", start="In the current marketplace environment, we believe", sentences=2,
         topic="Vì sao",
         vi="NIKE cho rằng cách người tiêu dùng mua sắm đang thay đổi rõ: họ muốn sản phẩm mới liên tục, dịch vụ tốt, giao "
            "hàng nhanh, và chuyển dần sang mua online và mua ở cửa hàng chỉ bán một thương hiệu như NIKE Direct."),
    dict(id="offense", filing="10-K FY2019", start="Through the Consumer Direct Offense", sentences=1, topic="Chiến lược",
         vi="Với chiến lược Consumer Direct Offense (“Triple Double”), NIKE đặt mục tiêu nhân đôi tác động của đổi mới sản "
            "phẩm, đưa hàng ra thị trường nhanh hơn và tăng kết nối trực tiếp với người tiêu dùng."),
    dict(id="acceleration", filing="10-K FY2020", start="In June 2020, we announced", sentences=1, topic="Chiến lược",
         vi="Tháng 6/2020, NIKE công bố giai đoạn mới của chiến lược, đặt tên Consumer Direct Acceleration, dựa nhiều hơn vào "
            "kênh số."),
    dict(id="digital_first", filing="10-K FY2020", start="This next phase of our Consumer Direct Offense", sentences=1,
         topic="Chiến lược",
         vi="NIKE kỳ vọng giai đoạn này mang lại tăng trưởng và lợi nhuận bền vững, khi đưa NIKE thành công ty lấy kênh số "
            "làm đầu (digital-first)."),
    dict(id="fueled", filing="10-K FY2021", start="Our NIKE Direct business fueled our growth", sentences=1, topic="Kết quả",
         vi="Theo NIKE, NIKE Direct là động lực tăng trưởng trong FY2021, nhờ kết hợp nền tảng số với hệ thống cửa hàng để "
            "bán trực tiếp cho người tiêu dùng."),
    dict(id="margin", filing="10-K FY2021", start="Additionally, the favorable impact of growth in our higher margin NIKE Direct",
         sentences=1, topic="Lợi nhuận",
         vi="NIKE nói NIKE Direct (dẫn đầu là Digital) có biên lợi nhuận cao hơn, nhưng trong FY2021 tác động này bị bù trừ "
            "bởi khuyến mại ở cửa hàng outlet nửa đầu năm để giảm tồn kho dư do COVID-19."),
    dict(id="demand_creation", filing="10-K FY2021", start="Demand creation expense decreased 13% for fiscal 2021", sentences=1,
         topic="Chi phí",
         vi="Chi phí tạo nhu cầu FY2021 giảm 13% vì chi cho sự kiện thương hiệu, quảng cáo và tài trợ thể thao giảm khi các "
            "giải đấu bị hoãn do COVID-19."),
]
QUOTES_BY_ID = {q["id"]: q for q in QUOTES}
