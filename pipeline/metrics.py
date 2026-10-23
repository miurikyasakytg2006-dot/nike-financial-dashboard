"""Catalogue of reported metrics, calculated KPIs and reconciliation checks.

This module is the single definition shared by the workbook builder, the validator and
the dashboard export. KPI and check formulas are written once, as arithmetic expressions
over metric ids; the same expression is evaluated in Python and translated into Excel
cell references, so the workbook formulas and the exported JSON cannot drift apart.

Identifiers ending in "__prior" mean "the prior fiscal year's value as presented in the
same 10-K" (the comparative column), which is the basis NIKE itself uses for % change.
"""

# Data status vocabulary (Vietnamese label, English label).
STATUS_MATCHED = "Khớp nguồn chính thức"
STATUS_PENDING = "Chờ người kiểm tra"
STATUS_HUMAN = "Đã xác minh (người)"
STATUS_MISSING = "Thiếu / không công bố"
STATUS_CONFLICT = "Bằng chứng mâu thuẫn"
STATUSES = {
    STATUS_MATCHED: "Matched to official source",
    STATUS_PENDING: "Pending human review",
    STATUS_HUMAN: "Human verified",
    STATUS_MISSING: "Missing or undisclosed",
    STATUS_CONFLICT: "Conflicting evidence",
}

USD_M = "USD triệu"
USD_B = "USD tỷ"
PCT = "%"

KIND_GAAP = "GAAP – BCTC đã kiểm toán"
KIND_NONGAAP = "Non-GAAP – NIKE định nghĩa"
KIND_MDNA = "MD&A – thuyết minh"
KIND_MDNA_GAAP = "MD&A – % tính từ số GAAP"

_DUR = "duration"
_INST = "instant"
_NB = {"srt:ConsolidationItemsAxis": "us-gaap:OperatingSegmentsMember",
       "us-gaap:StatementBusinessSegmentsAxis": "nke:NIKEBrandMember"}
_CHANNEL_AXIS = "us-gaap:ContractWithCustomerSalesChannelAxis"


def _xbrl(concepts, dims=None, period=_DUR, sign=1, section=r"CONSOLIDATED STATEMENTS OF INCOME"):
    return {"concepts": concepts, "dims": dims or {}, "period": period, "sign": sign, "section": section}


# Reported metrics. "sheet" is the workbook sheet the observation lives on.
REPORTED = [
    # ---- NIKE, Inc. consolidated (sheet 02) ----
    dict(id="revenues", sheet="financial", scope="NIKE, Inc.", channel="", unit=USD_M, kind=KIND_GAAP,
         vi="Doanh thu", en="Revenues",
         location="Item 8 – Consolidated Statements of Income, dòng “Revenues”",
         definition="Tổng doanh thu hợp nhất của NIKE, Inc. (NIKE Brand + Converse + Corporate).",
         xbrl=_xbrl(["us-gaap:RevenueFromContractWithCustomerExcludingAssessedTax"])),
    dict(id="cost_of_sales", sheet="financial", scope="NIKE, Inc.", channel="", unit=USD_M, kind=KIND_GAAP,
         vi="Giá vốn hàng bán", en="Cost of sales",
         location="Item 8 – Consolidated Statements of Income, dòng “Cost of sales”",
         definition="Giá vốn hàng bán hợp nhất.",
         xbrl=_xbrl(["us-gaap:CostOfGoodsAndServicesSold"])),
    dict(id="gross_profit", sheet="financial", scope="NIKE, Inc.", channel="", unit=USD_M, kind=KIND_GAAP,
         vi="Lợi nhuận gộp", en="Gross profit",
         location="Item 8 – Consolidated Statements of Income, dòng “Gross profit”",
         definition="Doanh thu trừ giá vốn hàng bán, như trình bày trên báo cáo kết quả kinh doanh.",
         xbrl=_xbrl(["us-gaap:GrossProfit"])),
    dict(id="demand_creation_expense", sheet="financial", scope="NIKE, Inc.", channel="", unit=USD_M, kind=KIND_GAAP,
         vi="Chi phí tạo nhu cầu", en="Demand creation expense",
         location="Item 8 – Consolidated Statements of Income, dòng “Demand creation expense”",
         definition="Cấu phần marketing của tổng chi phí bán hàng và quản lý. 10-K FY2020: chi phí quảng cáo, khuyến mại, hợp đồng tài trợ vận động viên, truyền thông, sự kiện thương hiệu; 10-K FY2026: “brand marketing expense and sports marketing expense”.",
         xbrl=_xbrl(["us-gaap:MarketingAndAdvertisingExpense"])),
    dict(id="operating_overhead_expense", sheet="financial", scope="NIKE, Inc.", channel="", unit=USD_M, kind=KIND_GAAP,
         vi="Chi phí vận hành chung", en="Operating overhead expense",
         location="Item 8 – Consolidated Statements of Income, dòng “Operating overhead expense”",
         definition="Cấu phần còn lại của tổng chi phí bán hàng và quản lý: chủ yếu lương và phúc lợi, nghiên cứu và phát triển, nợ khó đòi, tiền thuê, khấu hao, dịch vụ chuyên môn, đầu tư công nghệ, đi lại (theo 10-K FY2026).",
         xbrl=_xbrl(["us-gaap:GeneralAndAdministrativeExpense"])),
    dict(id="total_selling_admin_expense", sheet="financial", scope="NIKE, Inc.", channel="", unit=USD_M, kind=KIND_GAAP,
         vi="Tổng chi phí bán hàng và quản lý (SG&A)", en="Total selling and administrative expense",
         location="Item 8 – Consolidated Statements of Income, dòng “Total selling and administrative expense”",
         definition="Tổng chi phí tạo nhu cầu và chi phí vận hành chung. NIKE không phân bổ khoản này theo kênh phân phối.",
         xbrl=_xbrl(["us-gaap:SellingGeneralAndAdministrativeExpense"])),
    dict(id="interest_expense_net", sheet="financial", scope="NIKE, Inc.", channel="", unit=USD_M, kind=KIND_GAAP,
         vi="Chi phí (thu nhập) lãi thuần", en="Interest expense (income), net",
         location="Item 8 – Consolidated Statements of Income, dòng “Interest expense (income), net” (từ FY2024 ghi là “Interest (income) expense, net”)",
         definition="Số dương = chi phí lãi thuần; số âm = thu nhập lãi thuần. Giữ đúng quy ước dấu của NIKE.",
         xbrl=_xbrl(["us-gaap:InterestIncomeExpenseNonoperatingNet"], sign=-1)),
    dict(id="other_expense_net", sheet="financial", scope="NIKE, Inc.", channel="", unit=USD_M, kind=KIND_GAAP,
         vi="Chi phí (thu nhập) khác thuần", en="Other (income) expense, net",
         location="Item 8 – Consolidated Statements of Income, dòng “Other (income) expense, net”",
         definition="Số dương = chi phí khác thuần; số âm = thu nhập khác thuần. Giữ đúng quy ước dấu của NIKE.",
         xbrl=_xbrl(["us-gaap:OtherNonoperatingIncomeExpense"], sign=-1)),
    dict(id="income_before_taxes", sheet="financial", scope="NIKE, Inc.", channel="", unit=USD_M, kind=KIND_GAAP,
         vi="Lợi nhuận trước thuế", en="Income before income taxes",
         location="Item 8 – Consolidated Statements of Income, dòng “Income before income taxes”",
         definition="Lợi nhuận trước thuế thu nhập doanh nghiệp.",
         xbrl=_xbrl(["us-gaap:IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest",
                     "us-gaap:IncomeLossFromContinuingOperationsBeforeIncomeTaxesMinorityInterestAndIncomeLossFromEquityMethodInvestments"])),
    dict(id="income_tax_expense", sheet="financial", scope="NIKE, Inc.", channel="", unit=USD_M, kind=KIND_GAAP,
         vi="Chi phí thuế thu nhập", en="Income tax expense",
         location="Item 8 – Consolidated Statements of Income, dòng “Income tax expense”",
         definition="Chi phí thuế thu nhập doanh nghiệp.",
         xbrl=_xbrl(["us-gaap:IncomeTaxExpenseBenefit"])),
    dict(id="net_income", sheet="financial", scope="NIKE, Inc.", channel="", unit=USD_M, kind=KIND_GAAP,
         vi="Lợi nhuận sau thuế", en="Net income",
         location="Item 8 – Consolidated Statements of Income, dòng “Net income”",
         definition="Lợi nhuận thuần của NIKE, Inc.",
         xbrl=_xbrl(["us-gaap:NetIncomeLoss"])),
    dict(id="inventories", sheet="financial", scope="NIKE, Inc.", channel="", unit=USD_M, kind=KIND_GAAP,
         vi="Hàng tồn kho cuối kỳ", en="Inventories",
         location="Item 8 – Consolidated Balance Sheets, dòng “Inventories” (số dư tại 31/5)",
         definition="Số dư hàng tồn kho tại ngày kết thúc năm tài chính (31/5). Giá trị năm trước là số dư đầu kỳ.",
         xbrl=_xbrl(["us-gaap:InventoryFinishedGoodsNetOfReserves"], period=_INST, section=r"CONSOLIDATED BALANCE SHEETS")),
    dict(id="ebit_reported", sheet="financial", scope="NIKE, Inc.", channel="", unit=USD_M, kind=KIND_NONGAAP,
         vi="EBIT do NIKE công bố", en="Total NIKE, Inc. Earnings Before Interest and Taxes (EBIT)",
         location="Item 7 – MD&A, bảng “Total NIKE, Inc. EBIT” (mục Use of Non-GAAP Financial Measures / Operating Segments)",
         definition="Lợi nhuận sau thuế cộng chi phí (thu nhập) lãi thuần và chi phí thuế. NIKE ghi rõ đây là chỉ tiêu non-GAAP. Dùng để đối chiếu với EBIT tự tính ở sheet 04.",
         mdna="ebit"),
    dict(id="revenues_pct_change_reported", sheet="financial", scope="NIKE, Inc.", channel="", unit=PCT, kind=KIND_MDNA_GAAP,
         vi="% thay đổi doanh thu (NIKE công bố, đã làm tròn)", en="TOTAL NIKE, INC. REVENUES – % change",
         location="Item 7 – MD&A, bảng doanh thu NIKE, Inc., cột “% CHANGE”",
         definition="Tỷ lệ tăng trưởng theo báo cáo, NIKE làm tròn đến số nguyên. Dùng để đối chiếu với tăng trưởng tự tính.",
         mdna="revenue_table:TOTAL NIKE, INC. REVENUES:pct"),
    dict(id="revenues_pct_change_cn", sheet="financial", scope="NIKE, Inc.", channel="", unit=PCT, kind=KIND_NONGAAP,
         vi="% thay đổi doanh thu loại trừ biến động tỷ giá", en="TOTAL NIKE, INC. REVENUES – % change excluding currency changes",
         location="Item 7 – MD&A, bảng doanh thu NIKE, Inc., cột “% CHANGE EXCLUDING CURRENCY CHANGES”",
         definition="Tăng trưởng trung lập tỷ giá (currency-neutral), chỉ tiêu non-GAAP do NIKE tính theo tỷ giá năm trước. Không thể tự tính lại từ BCTC.",
         mdna="revenue_table:TOTAL NIKE, INC. REVENUES:pct_cn"),

    # ---- Distribution channels (sheet 03) ----
    dict(id="nike_direct_revenue", sheet="channels", scope="NIKE Brand", channel="NIKE Direct", unit=USD_M, kind=KIND_GAAP,
         vi="Doanh thu NIKE Direct", en="Sales through Direct to Consumer (NIKE Brand)",
         location="Item 8 – Note “Revenues”, bảng phân tách doanh thu, cột “Total NIKE Brand”, dòng “Sales through Direct to Consumer”",
         definition="Doanh thu NIKE Brand bán trực tiếp cho người tiêu dùng qua cửa hàng do NIKE sở hữu và nền tảng số (“NIKE Brand Digital”) – theo định nghĩa trong Item 1 của 10-K. Không gồm kênh trực tiếp của Converse. Thuyết minh BCTC gọi dòng này là “Sales through Direct to Consumer”, MD&A gọi là “Sales through NIKE Direct”.",
         xbrl=_xbrl(["us-gaap:RevenueFromContractWithCustomerExcludingAssessedTax"],
                    dims={**_NB, _CHANNEL_AXIS: "us-gaap:SalesChannelDirectlyToConsumerMember"}, section=r"NOTE \d+ — REVENUES")),
    dict(id="nike_brand_wholesale_revenue", sheet="channels", scope="NIKE Brand", channel="Bán buôn (Wholesale)", unit=USD_M, kind=KIND_GAAP,
         vi="Doanh thu bán buôn NIKE Brand", en="Sales to Wholesale Customers (NIKE Brand)",
         location="Item 8 – Note “Revenues”, bảng phân tách doanh thu, cột “Total NIKE Brand”, dòng “Sales to Wholesale Customers”",
         definition="Doanh thu NIKE Brand bán cho khách hàng bán buôn (nhà bán lẻ, nhà phân phối, đối tác). Không gồm bán buôn của Converse.",
         xbrl=_xbrl(["us-gaap:RevenueFromContractWithCustomerExcludingAssessedTax"],
                    dims={**_NB, _CHANNEL_AXIS: "nke:SalesChannelWholesaleMember"}, section=r"NOTE \d+ — REVENUES")),
    dict(id="global_brand_divisions_revenue", sheet="channels", scope="NIKE Brand", channel="Global Brand Divisions", unit=USD_M, kind=KIND_GAAP,
         vi="Doanh thu Global Brand Divisions", en="Other / Global Brand Divisions (NIKE Brand)",
         location="Item 8 – Note “Revenues”, bảng phân tách doanh thu, cột “Total NIKE Brand”, dòng “Other”",
         definition="“NIKE Brand licensing and other miscellaneous revenues” không thuộc khu vực địa lý nào; không thuộc kênh Direct hay Wholesale. Quy mô rất nhỏ.",
         xbrl=_xbrl(["us-gaap:RevenueFromContractWithCustomerExcludingAssessedTax"],
                    dims={**_NB, _CHANNEL_AXIS: "nke:SalesChannelOtherMember"}, section=r"NOTE \d+ — REVENUES")),
    dict(id="nike_brand_revenue", sheet="channels", scope="NIKE Brand", channel="Tổng NIKE Brand", unit=USD_M, kind=KIND_GAAP,
         vi="Tổng doanh thu NIKE Brand", en="Total NIKE Brand revenues",
         location="Item 8 – Note “Revenues”, bảng phân tách doanh thu, cột “Total NIKE Brand”, dòng “Total revenues”",
         definition="NIKE Direct + Bán buôn NIKE Brand + Global Brand Divisions. Không gồm Converse và Corporate.",
         xbrl=_xbrl(["us-gaap:RevenueFromContractWithCustomerExcludingAssessedTax"], dims=_NB, section=r"NOTE \d+ — REVENUES")),
    dict(id="converse_revenue", sheet="channels", scope="Converse", channel="", unit=USD_M, kind=KIND_GAAP,
         vi="Doanh thu Converse", en="Converse revenues",
         location="Item 8 – Note “Revenues”, bảng phân tách doanh thu, cột “Converse”, dòng “Total revenues”",
         definition="Doanh thu thương hiệu Converse (gồm cả kênh trực tiếp và bán buôn của Converse). Dùng để đối chiếu NIKE Brand với NIKE, Inc.",
         xbrl=_xbrl(["us-gaap:RevenueFromContractWithCustomerExcludingAssessedTax"],
                    dims={"srt:ConsolidationItemsAxis": "us-gaap:OperatingSegmentsMember",
                          "us-gaap:StatementBusinessSegmentsAxis": "nke:ConverseSegmentMember"}, section=r"NOTE \d+ — REVENUES")),
    dict(id="corporate_revenue", sheet="channels", scope="Corporate", channel="", unit=USD_M, kind=KIND_GAAP,
         vi="Doanh thu Corporate", en="Corporate revenues",
         location="Item 8 – Note “Revenues”, bảng phân tách doanh thu, cột “Corporate”, dòng “Total revenues”",
         definition="Chủ yếu là lãi/lỗ phòng ngừa rủi ro tỷ giá liên quan đến doanh thu; có thể âm. Dùng để đối chiếu.",
         xbrl=_xbrl(["us-gaap:RevenueFromContractWithCustomerExcludingAssessedTax"],
                    dims={"srt:ConsolidationItemsAxis": "us-gaap:CorporateNonSegmentMember"}, section=r"NOTE \d+ — REVENUES")),
    dict(id="nike_brand_digital_sales", sheet="channels", scope="NIKE Brand", channel="NIKE Direct – Digital", unit=USD_B, kind=KIND_MDNA,
         vi="Doanh số NIKE Brand Digital", en="NIKE Brand Digital sales (FY2020–21: “Digital commerce sales”)",
         location="Item 7 – MD&A, đoạn văn về NIKE Direct (câu trích ở cột “Số hoặc câu hiển thị trong tài liệu”)",
         definition="Doanh số qua nền tảng số của NIKE Brand – một phần của NIKE Direct. NIKE chỉ công bố dạng làm tròn đến 0,1 tỷ USD trong phần thuyết minh.",
         mdna="digital"),
    dict(id="nike_direct_share_stated", sheet="channels", scope="NIKE Brand", channel="NIKE Direct", unit=PCT, kind=KIND_MDNA,
         vi="Tỷ trọng NIKE Direct do NIKE nêu (xấp xỉ)", en="NIKE Direct as approx. % of total NIKE Brand revenues (as stated)",
         location="Item 7 – MD&A, câu “NIKE Direct revenues represented approximately …% of total NIKE Brand revenues”",
         definition="Tỷ trọng làm tròn do NIKE tự nêu trong MD&A. Dùng để đối chiếu với tỷ trọng tự tính ở sheet 04.",
         mdna="share"),
    dict(id="nike_direct_pct_change_reported", sheet="channels", scope="NIKE Brand", channel="NIKE Direct", unit=PCT, kind=KIND_MDNA_GAAP,
         vi="% thay đổi NIKE Direct (NIKE công bố, đã làm tròn)", en="Sales through NIKE Direct – % change",
         location="Item 7 – MD&A, bảng doanh thu, mục Supplemental NIKE Brand Revenues Details, cột “% CHANGE”",
         definition="Tăng trưởng theo báo cáo do NIKE làm tròn đến số nguyên. Dùng để đối chiếu.",
         mdna="revenue_table:Sales through NIKE Direct:pct"),
    dict(id="nike_direct_pct_change_cn", sheet="channels", scope="NIKE Brand", channel="NIKE Direct", unit=PCT, kind=KIND_NONGAAP,
         vi="% thay đổi NIKE Direct loại trừ tỷ giá", en="Sales through NIKE Direct – % change excluding currency changes",
         location="Item 7 – MD&A, bảng doanh thu, mục Supplemental NIKE Brand Revenues Details, cột “% CHANGE EXCLUDING CURRENCY CHANGES”",
         definition="Tăng trưởng trung lập tỷ giá (non-GAAP). Không tự tính lại được từ BCTC.",
         mdna="revenue_table:Sales through NIKE Direct:pct_cn"),
    dict(id="nike_brand_wholesale_pct_change_reported", sheet="channels", scope="NIKE Brand", channel="Bán buôn (Wholesale)", unit=PCT, kind=KIND_MDNA_GAAP,
         vi="% thay đổi bán buôn (NIKE công bố, đã làm tròn)", en="Sales to Wholesale Customers – % change",
         location="Item 7 – MD&A, bảng doanh thu, mục Supplemental NIKE Brand Revenues Details, cột “% CHANGE”",
         definition="Tăng trưởng theo báo cáo do NIKE làm tròn đến số nguyên. Dùng để đối chiếu.",
         mdna="revenue_table:Sales to Wholesale Customers:pct"),
    dict(id="nike_brand_wholesale_pct_change_cn", sheet="channels", scope="NIKE Brand", channel="Bán buôn (Wholesale)", unit=PCT, kind=KIND_NONGAAP,
         vi="% thay đổi bán buôn loại trừ tỷ giá", en="Sales to Wholesale Customers – % change excluding currency changes",
         location="Item 7 – MD&A, bảng doanh thu, mục Supplemental NIKE Brand Revenues Details, cột “% CHANGE EXCLUDING CURRENCY CHANGES”",
         definition="Tăng trưởng trung lập tỷ giá (non-GAAP). Không tự tính lại được từ BCTC.",
         mdna="revenue_table:Sales to Wholesale Customers:pct_cn"),
    dict(id="nike_brand_pct_change_reported", sheet="channels", scope="NIKE Brand", channel="Tổng NIKE Brand", unit=PCT, kind=KIND_MDNA_GAAP,
         vi="% thay đổi doanh thu NIKE Brand (NIKE công bố, đã làm tròn)", en="TOTAL NIKE BRAND REVENUES – % change",
         location="Item 7 – MD&A, bảng doanh thu, mục Supplemental NIKE Brand Revenues Details, cột “% CHANGE”",
         definition="Tăng trưởng theo báo cáo do NIKE làm tròn đến số nguyên. Dùng để đối chiếu.",
         mdna="revenue_table:TOTAL NIKE BRAND REVENUES:pct"),
    dict(id="nike_brand_pct_change_cn", sheet="channels", scope="NIKE Brand", channel="Tổng NIKE Brand", unit=PCT, kind=KIND_NONGAAP,
         vi="% thay đổi NIKE Brand loại trừ tỷ giá", en="TOTAL NIKE BRAND REVENUES – % change excluding currency changes",
         location="Item 7 – MD&A, bảng doanh thu, mục Supplemental NIKE Brand Revenues Details, cột “% CHANGE EXCLUDING CURRENCY CHANGES”",
         definition="Tăng trưởng trung lập tỷ giá (non-GAAP). Không tự tính lại được từ BCTC.",
         mdna="revenue_table:TOTAL NIKE BRAND REVENUES:pct_cn"),
]
REPORTED_BY_ID = {m["id"]: m for m in REPORTED}

# Calculated KPIs, in evaluation order (a KPI may use KPIs defined above it).
KPIS = [
    # Distribution strategy – NIKE Brand scope
    dict(id="direct_share", group="channels", unit=PCT, vi="Tỷ trọng NIKE Direct trong doanh thu NIKE Brand",
         expr="nike_direct_revenue / nike_brand_revenue",
         note="Mẫu số là tổng doanh thu NIKE Brand (cùng cơ sở NIKE dùng khi nêu tỷ trọng), không phải NIKE, Inc."),
    dict(id="wholesale_share", group="channels", unit=PCT, vi="Tỷ trọng bán buôn trong doanh thu NIKE Brand",
         expr="nike_brand_wholesale_revenue / nike_brand_revenue",
         note="Direct + Wholesale + Global Brand Divisions = 100% doanh thu NIKE Brand."),
    dict(id="gbd_share", group="channels", unit=PCT, vi="Tỷ trọng Global Brand Divisions trong doanh thu NIKE Brand",
         expr="global_brand_divisions_revenue / nike_brand_revenue",
         note="Phần còn lại, chủ yếu là doanh thu cấp phép; dưới 0,3%."),
    dict(id="direct_growth", group="channels", unit=PCT, vi="Tăng trưởng doanh thu NIKE Direct (theo báo cáo)",
         expr="nike_direct_revenue / nike_direct_revenue__prior - 1",
         note="So với số năm trước trình bày trong cùng 10-K. Không loại trừ tỷ giá."),
    dict(id="wholesale_growth", group="channels", unit=PCT, vi="Tăng trưởng doanh thu bán buôn NIKE Brand (theo báo cáo)",
         expr="nike_brand_wholesale_revenue / nike_brand_wholesale_revenue__prior - 1",
         note="So với số năm trước trình bày trong cùng 10-K. Không loại trừ tỷ giá."),
    dict(id="nike_brand_growth", group="channels", unit=PCT, vi="Tăng trưởng doanh thu NIKE Brand (theo báo cáo)",
         expr="nike_brand_revenue / nike_brand_revenue__prior - 1",
         note="So với số năm trước trình bày trong cùng 10-K."),
    dict(id="digital_share_of_direct", group="channels", unit=PCT, vi="Tỷ trọng Digital trong NIKE Direct (xấp xỉ)",
         expr="nike_brand_digital_sales * 1000 / nike_direct_revenue",
         note="Xấp xỉ: doanh số Digital chỉ được công bố làm tròn 0,1 tỷ USD (sai số khoảng ±0,3 điểm %). Nhân 1.000 để đổi tỷ USD sang triệu USD."),
    # Financial performance – NIKE, Inc. scope
    dict(id="revenue_growth", group="financial", unit=PCT, vi="Tăng trưởng doanh thu NIKE, Inc. (theo báo cáo)",
         expr="revenues / revenues__prior - 1", note="So với số năm trước trình bày trong cùng 10-K."),
    dict(id="gross_margin", group="financial", unit=PCT, vi="Biên lợi nhuận gộp",
         expr="gross_profit / revenues", note="Lợi nhuận gộp / Doanh thu, cùng phạm vi NIKE, Inc."),
    dict(id="sga_ratio", group="financial", unit=PCT, vi="Tỷ lệ chi phí bán hàng và quản lý / doanh thu",
         expr="total_selling_admin_expense / revenues", note="Chi phí hợp nhất; không phân bổ được theo kênh."),
    dict(id="demand_creation_ratio", group="financial", unit=PCT, vi="Tỷ lệ chi phí tạo nhu cầu / doanh thu",
         expr="demand_creation_expense / revenues", note="Cấu phần marketing của SG&A."),
    dict(id="operating_overhead_ratio", group="financial", unit=PCT, vi="Tỷ lệ chi phí vận hành chung / doanh thu",
         expr="operating_overhead_expense / revenues", note="Cấu phần vận hành của SG&A."),
    dict(id="ebit", group="financial", unit=USD_M, vi="EBIT (tự tính theo định nghĩa của NIKE)",
         expr="net_income + income_tax_expense + interest_expense_net",
         note="Non-GAAP. NIKE không trình bày dòng “Operating income”; EBIT là thước đo NIKE dùng. Đối chiếu với EBIT do NIKE công bố."),
    dict(id="ebit_margin", group="financial", unit=PCT, vi="Biên EBIT",
         expr="ebit / revenues", note="Non-GAAP. EBIT / Doanh thu NIKE, Inc."),
    dict(id="net_income_margin", group="financial", unit=PCT, vi="Biên lợi nhuận sau thuế",
         expr="net_income / revenues", note="GAAP."),
    dict(id="inventory_growth", group="financial", unit=PCT, vi="Thay đổi hàng tồn kho so với đầu kỳ",
         expr="inventories / inventories__prior - 1", note="Số dư 31/5 năm nay so với 31/5 năm trước."),
    dict(id="average_inventory", group="financial", unit=USD_M, vi="Hàng tồn kho bình quân",
         expr="(inventories + inventories__prior) / 2", note="Bình quân số dư đầu kỳ và cuối kỳ."),
    dict(id="inventory_turnover", group="financial", unit="lần", vi="Vòng quay hàng tồn kho",
         expr="cost_of_sales / average_inventory", note="Giá vốn hàng bán / Hàng tồn kho bình quân."),
    dict(id="days_inventory", group="financial", unit="ngày", vi="Số ngày tồn kho bình quân",
         expr="365 / inventory_turnover", note="365 / Vòng quay hàng tồn kho."),
]
KPIS_BY_ID = {k["id"]: k for k in KPIS}

# Reconciliation checks: expression should be 0 within the tolerance (USD million).
CHECKS = [
    dict(id="chk_gross_profit", vi="Doanh thu − Giá vốn − Lợi nhuận gộp", tol=1,
         expr="revenues - cost_of_sales - gross_profit"),
    dict(id="chk_sga", vi="Tạo nhu cầu + Vận hành chung − Tổng SG&A", tol=1,
         expr="demand_creation_expense + operating_overhead_expense - total_selling_admin_expense"),
    dict(id="chk_pretax", vi="Lợi nhuận gộp − SG&A − Lãi thuần − Khác thuần − LNTT", tol=2,
         expr="gross_profit - total_selling_admin_expense - interest_expense_net - other_expense_net - income_before_taxes"),
    dict(id="chk_net_income", vi="LNTT − Thuế − LNST", tol=1,
         expr="income_before_taxes - income_tax_expense - net_income"),
    dict(id="chk_channels", vi="Direct + Wholesale + GBD − Tổng NIKE Brand", tol=1,
         expr="nike_direct_revenue + nike_brand_wholesale_revenue + global_brand_divisions_revenue - nike_brand_revenue"),
    dict(id="chk_segments", vi="NIKE Brand + Converse + Corporate − Doanh thu NIKE, Inc.", tol=1,
         expr="nike_brand_revenue + converse_revenue + corporate_revenue - revenues"),
    dict(id="chk_ebit", vi="EBIT tự tính − EBIT NIKE công bố", tol=1,
         expr="ebit - ebit_reported"),
]

# Calculated ratios compared with NIKE's own rounded figures (difference as a fraction;
# tolerance 0.01 = 1 percentage point, because NIKE rounds to whole percent).
COMPARISONS = [
    dict(id="cmp_revenue_growth", vi="Tăng trưởng doanh thu tự tính − % NIKE công bố", tol=0.01,
         expr="revenue_growth - revenues_pct_change_reported", confirms="revenues_pct_change_reported"),
    dict(id="cmp_direct_growth", vi="Tăng trưởng NIKE Direct tự tính − % NIKE công bố", tol=0.01,
         expr="direct_growth - nike_direct_pct_change_reported", confirms="nike_direct_pct_change_reported"),
    dict(id="cmp_wholesale_growth", vi="Tăng trưởng bán buôn tự tính − % NIKE công bố", tol=0.01,
         expr="wholesale_growth - nike_brand_wholesale_pct_change_reported", confirms="nike_brand_wholesale_pct_change_reported"),
    dict(id="cmp_brand_growth", vi="Tăng trưởng NIKE Brand tự tính − % NIKE công bố", tol=0.01,
         expr="nike_brand_growth - nike_brand_pct_change_reported", confirms="nike_brand_pct_change_reported"),
    dict(id="cmp_direct_share", vi="Tỷ trọng NIKE Direct tự tính − tỷ trọng NIKE nêu", tol=0.01,
         expr="direct_share - nike_direct_share_stated", confirms="nike_direct_share_stated"),
]
CHECK_CONFIRMS = {"chk_ebit": "ebit_reported"}
