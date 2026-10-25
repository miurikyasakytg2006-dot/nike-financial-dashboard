"""Catalogue of reported metrics, calculated KPIs and reconciliation checks.

This module is the single definition shared by the workbook builder, the validator and
the dashboard export. KPI and check formulas are written once, as arithmetic expressions
over metric ids; the same expression is evaluated in Python and translated into Excel
cell references, so the workbook formulas and the exported JSON cannot drift apart.

Every metric is observed by fiscal year ("A": FY2019–FY2021) and, where NIKE reports it
quarterly, by fiscal quarter ("Q": Q1 FY2019 – Q4 FY2021). Identifiers ending in "__prior"
mean "the same period one year earlier, as presented in the same report" (the comparative
column), which is the basis NIKE itself uses for % change; a quarter-end inventory balance,
which the 10-Q compares with 31 May instead, takes it from the previous year's 10-Q.
"__begin" (inventories only) is the balance at the start of the period.
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

USD_M = "triệu USD"
USD_B = "tỷ USD"
PCT = "%"

KIND_GAAP = "Báo cáo tài chính"   # the workbook shows KIND_GAAP_A or KIND_GAAP_Q, by the report the row comes from
KIND_GAAP_A = "Báo cáo tài chính năm, đã kiểm toán (10-K)"
KIND_GAAP_Q = "Báo cáo tài chính quý, không kiểm toán (10-Q)"
KIND_NONGAAP = "Chỉ tiêu NIKE tự định nghĩa (non-GAAP)"
KIND_MDNA = "Ban điều hành nêu trong MD&A, không kiểm toán"
KIND_MDNA_GAAP = "Tỷ lệ % NIKE in trong MD&A"
KIND_Q4 = "Nhóm tự tính: cả năm trừ 9 tháng đầu"
KIND_RELEASE = "Thông cáo kết quả quý 4 của NIKE (8-K), không kiểm toán"

A, Q = "A", "Q"
_DUR = "duration"
_INST = "instant"
# Dimension members, compared by local name (the prefix of an axis changed between taxonomies: us-gaap → srt).
_SEG = {"ConsolidationItemsAxis": "OperatingSegmentsMember"}
_NB = {**_SEG, "StatementBusinessSegmentsAxis": "NIKEBrandMember"}
_CH = "ContractWithCustomerSalesChannelAxis"
_DTC, _WHS, _OTH = "SalesChannelDirectlyToConsumerMember", "SalesChannelWholesaleMember", "SalesChannelOtherMember"
REVENUE_CONCEPTS = ["RevenueFromContractWithCustomerExcludingAssessedTax",
                    "RevenueFromContractWithCustomerIncludingAssessedTax", "SalesRevenueNet"]
STATEMENT = r"STATEMENTS OF INCOME|^ITEM 1\. FINANCIAL STATEMENTS"
BALANCE = r"BALANCE SHEETS|^ITEM 1\. FINANCIAL STATEMENTS"
NOTE_REVENUES = r"NOTE \d+ ?[—–-] ?REVENUES"
NOTE_SEGMENTS = r"NOTE \d+ ?[—–-] ?OPERATING SEGMENTS"

REGIONS = [
    # id, Vietnamese name, short label, name in the 10-K, XBRL member
    ("na", "Bắc Mỹ", "Bắc Mỹ", "North America", "NorthAmericaSegmentMember"),
    ("emea", "Châu Âu, Trung Đông và châu Phi", "EMEA", "Europe, Middle East & Africa", "EuropeMiddleEastAndAfricaSegmentMember"),
    ("gc", "Đại Trung Hoa", "Đại Trung Hoa", "Greater China", "GreaterChinaSegmentMember"),
    ("apla", "Châu Á – Thái Bình Dương và Mỹ Latinh", "APLA", "Asia Pacific & Latin America", "AsiaPacificAndLatinAmericaSegmentMember"),
]


def _xbrl(concepts, dims=None, period=_DUR, sign=1, section=STATEMENT, row=None, alt_dims=()):
    """XBRL source of a metric. `row` is the printed row label, used to find the value in reports filed before
    inline XBRL; `alt_dims` lists other dimension sets that tag the same amount in some years."""
    return {"concepts": concepts, "dims": [dims or {}] + list(alt_dims), "period": period, "sign": sign,
            "section": section, "row": row}


def _m(**kw):
    kw.setdefault("freq", {A, Q})
    kw.setdefault("region", "")
    kw.setdefault("channel", "")
    return dict(kw)


# Reported metrics. "sheet" groups them in the workbook: financial (NIKE, Inc.), channels (NIKE Brand), regions.
REPORTED = [
    # ---- NIKE, Inc. consolidated ----
    _m(id="revenues", sheet="financial", scope="NIKE, Inc.", unit=USD_M, kind=KIND_GAAP,
       vi="Doanh thu", en="Revenues", line="Revenues",
       definition="Tổng doanh thu hợp nhất của NIKE, Inc. (NIKE Brand + Converse + Corporate).",
       xbrl=_xbrl(REVENUE_CONCEPTS, row=r"Revenues")),
    _m(id="cost_of_sales", sheet="financial", scope="NIKE, Inc.", unit=USD_M, kind=KIND_GAAP,
       vi="Giá vốn hàng bán", en="Cost of sales", line="Cost of sales",
       definition="Giá vốn hàng bán hợp nhất.",
       xbrl=_xbrl(["CostOfGoodsAndServicesSold", "CostOfGoodsSold"], row=r"Cost of sales")),
    _m(id="gross_profit", sheet="financial", scope="NIKE, Inc.", unit=USD_M, kind=KIND_GAAP,
       vi="Lợi nhuận gộp", en="Gross profit", line="Gross profit",
       definition="Doanh thu trừ giá vốn hàng bán, như trình bày trên báo cáo kết quả kinh doanh.",
       xbrl=_xbrl(["GrossProfit"], row=r"Gross profit")),
    _m(id="demand_creation_expense", sheet="financial", scope="NIKE, Inc.", unit=USD_M, kind=KIND_GAAP,
       vi="Chi phí tạo nhu cầu", en="Demand creation expense", line="Demand creation expense",
       definition="Chi phí quảng cáo và quảng bá, một phần của tổng chi phí bán hàng và quản lý: hợp đồng tài trợ "
                  "(endorsement), sản phẩm tặng, quảng cáo truyền hình, kỹ thuật số và báo in, chi phí truyền thông, sự kiện "
                  "thương hiệu và trưng bày thương hiệu ở cửa hàng (theo 10-K FY2020).",
       xbrl=_xbrl(["MarketingAndAdvertisingExpense"], row=r"Demand creation expense")),
    _m(id="operating_overhead_expense", sheet="financial", scope="NIKE, Inc.", unit=USD_M, kind=KIND_GAAP,
       vi="Chi phí vận hành chung", en="Operating overhead expense", line="Operating overhead expense",
       definition="Phần còn lại của tổng chi phí bán hàng và quản lý, chủ yếu là lương và phúc lợi, chi phí nghiên cứu và "
                  "phát triển, nợ khó đòi, và chi phí hành chính khác như tiền thuê, khấu hao, dịch vụ chuyên môn, hội họp và "
                  "đi lại (theo 10-K FY2020).",
       xbrl=_xbrl(["GeneralAndAdministrativeExpense"], row=r"Operating overhead expense")),
    _m(id="total_selling_admin_expense", sheet="financial", scope="NIKE, Inc.", unit=USD_M, kind=KIND_GAAP,
       vi="Tổng chi phí bán hàng và quản lý (SG&A)", en="Total selling and administrative expense",
       line="Total selling and administrative expense",
       definition="Tổng chi phí tạo nhu cầu và chi phí vận hành chung. NIKE không công bố khoản này theo kênh phân phối.",
       xbrl=_xbrl(["SellingGeneralAndAdministrativeExpense"], row=r"Total selling and administrative expense")),
    _m(id="interest_expense_net", sheet="financial", scope="NIKE, Inc.", unit=USD_M, kind=KIND_GAAP,
       vi="Chi phí (thu nhập) lãi thuần", en="Interest expense (income), net", line="Interest expense (income), net",
       definition="Số dương là chi phí lãi thuần, số âm là thu nhập lãi thuần, giữ đúng cách ghi dấu của NIKE.",
       xbrl=_xbrl(["InterestIncomeExpenseNonoperatingNet"], sign=-1, row=r"Interest expense \(income\), net")),
    _m(id="other_expense_net", sheet="financial", scope="NIKE, Inc.", unit=USD_M, kind=KIND_GAAP,
       vi="Chi phí (thu nhập) khác thuần", en="Other (income) expense, net", line="Other (income) expense, net",
       definition="Số dương là chi phí khác thuần, số âm là thu nhập khác thuần, giữ đúng cách ghi dấu của NIKE. "
                  "10-Q quý 1 FY2019 ghi dòng này là “Other expense (income), net”.",
       xbrl=_xbrl(["OtherNonoperatingIncomeExpense"], sign=-1, row=r"Other (\(income\) expense|expense \(income\)), net")),
    _m(id="income_before_taxes", sheet="financial", scope="NIKE, Inc.", unit=USD_M, kind=KIND_GAAP,
       vi="Lợi nhuận trước thuế", en="Income before income taxes", line="Income before income taxes",
       definition="Lợi nhuận trước thuế thu nhập doanh nghiệp.",
       xbrl=_xbrl(["IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest",
                   "IncomeLossFromContinuingOperationsBeforeIncomeTaxesMinorityInterestAndIncomeLossFromEquityMethodInvestments"],
                  row=r"Income before income taxes")),
    _m(id="income_tax_expense", sheet="financial", scope="NIKE, Inc.", unit=USD_M, kind=KIND_GAAP,
       vi="Chi phí thuế thu nhập", en="Income tax expense", line="Income tax expense",
       definition="Chi phí thuế thu nhập doanh nghiệp.",
       xbrl=_xbrl(["IncomeTaxExpenseBenefit"], row=r"Income tax expense( \(benefit\))?")),
    _m(id="net_income", sheet="financial", scope="NIKE, Inc.", unit=USD_M, kind=KIND_GAAP,
       vi="Lợi nhuận sau thuế", en="Net income", line="Net income",
       definition="Lợi nhuận sau thuế của NIKE, Inc. Quý 4 FY2020 lỗ.",
       xbrl=_xbrl(["NetIncomeLoss"], row=r"Net income( \(loss\))?")),
    _m(id="inventories", sheet="financial", scope="NIKE, Inc.", unit=USD_M, kind=KIND_GAAP,
       vi="Hàng tồn kho cuối kỳ", en="Inventories", line="Inventories",
       definition="Số dư hàng tồn kho tại ngày cuối kỳ: 31/5 với năm, ngày cuối quý với quý. Số dư cuối quý 4 là số dư 31/5 "
                  "trong 10-K.",
       xbrl=_xbrl(["InventoryFinishedGoodsNetOfReserves"], period=_INST, section=BALANCE, row=r"Inventories")),
    _m(id="ebit_reported", sheet="financial", scope="NIKE, Inc.", unit=USD_M, kind=KIND_NONGAAP,
       vi="EBIT do NIKE công bố", en="Total NIKE, Inc. Earnings Before Interest and Taxes (EBIT)",
       line="TOTAL NIKE, INC. EARNINGS BEFORE INTEREST AND TAXES",
       definition="Lợi nhuận sau thuế cộng chi phí (thu nhập) lãi thuần và chi phí thuế. Từ FY2020, NIKE ghi rõ đây là "
                  "chỉ tiêu non-GAAP. Dùng để đối chiếu với EBIT tự tính.",
       mdna="ebit"),
    _m(id="revenues_pct_change_reported", sheet="financial", scope="NIKE, Inc.", unit=PCT, kind=KIND_MDNA_GAAP,
       vi="% thay đổi doanh thu (NIKE công bố, đã làm tròn)", en="TOTAL NIKE, INC. REVENUES – % change",
       line="TOTAL NIKE, INC. REVENUES, cột “% CHANGE”",
       definition="Tăng trưởng so với cùng kỳ năm trước, tính theo USD; NIKE làm tròn đến số nguyên. Dùng để đối chiếu.",
       mdna="revenue_table:TOTAL NIKE, INC. REVENUES:pct"),
    _m(id="revenues_pct_change_cn", sheet="financial", scope="NIKE, Inc.", unit=PCT, kind=KIND_NONGAAP,
       vi="% thay đổi doanh thu loại trừ tỷ giá", en="TOTAL NIKE, INC. REVENUES – % change excluding currency changes",
       line="TOTAL NIKE, INC. REVENUES, cột “% CHANGE EXCLUDING CURRENCY CHANGES”",
       definition="Tăng trưởng loại trừ tỷ giá (NIKE gọi là currency-neutral): chỉ tiêu non-GAAP, NIKE tính lại theo tỷ "
                  "giá của cùng kỳ năm trước để bỏ phần chênh lệch do quy đổi ngoại tệ. Không tự tính lại được từ báo cáo "
                  "tài chính.",
       mdna="revenue_table:TOTAL NIKE, INC. REVENUES:pct_cn"),

    # ---- Distribution channels, NIKE Brand ----
    _m(id="nike_direct_revenue", sheet="channels", scope="NIKE Brand", channel="NIKE Direct", unit=USD_M, kind=KIND_GAAP,
       vi="Doanh thu NIKE Direct", en="Sales through Direct to Consumer (NIKE Brand)",
       line="Sales through Direct to Consumer, cột Total NIKE Brand",
       definition="Doanh thu NIKE Brand bán trực tiếp cho người tiêu dùng, qua cửa hàng do NIKE sở hữu và Digital (ứng "
                  "dụng, website của NIKE). Không gồm kênh trực tiếp của Converse. Thuyết minh gọi dòng này là “Sales through "
                  "Direct to Consumer”, MD&A gọi là “Sales through NIKE Direct”.",
       xbrl=_xbrl(REVENUE_CONCEPTS, dims={**_NB, _CH: _DTC}, section=NOTE_REVENUES,
                  row=r"Sales through Direct to Consumer")),
    _m(id="nike_brand_wholesale_revenue", sheet="channels", scope="NIKE Brand", channel="Bán buôn (Wholesale)", unit=USD_M,
       kind=KIND_GAAP, vi="Doanh thu bán buôn NIKE Brand", en="Sales to Wholesale Customers (NIKE Brand)",
       line="Sales to Wholesale Customers, cột Total NIKE Brand",
       definition="Doanh thu NIKE Brand bán sản phẩm cho nhà bán lẻ và nhà phân phối. Không gồm Converse.",
       xbrl=_xbrl(REVENUE_CONCEPTS, dims={**_NB, _CH: _WHS}, section=NOTE_REVENUES, row=r"Sales to Wholesale Customers")),
    _m(id="global_brand_divisions_revenue", sheet="channels", scope="NIKE Brand", channel="Global Brand Divisions",
       unit=USD_M, kind=KIND_GAAP, vi="Doanh thu Global Brand Divisions", en="Other / Global Brand Divisions (NIKE Brand)",
       line="Other, cột Total NIKE Brand",
       definition="Doanh thu từ cấp phép thương hiệu NIKE (licensing) và các khoản nhỏ khác, không thuộc khu vực địa lý "
                  "nào, cũng không thuộc NIKE Direct hay bán buôn.",
       xbrl=_xbrl(REVENUE_CONCEPTS, dims={**_NB, _CH: _OTH}, section=NOTE_REVENUES, row=r"Other")),
    _m(id="nike_brand_revenue", sheet="channels", scope="NIKE Brand", channel="Tổng NIKE Brand", unit=USD_M, kind=KIND_GAAP,
       vi="Tổng doanh thu NIKE Brand", en="Total NIKE Brand revenues", line="TOTAL REVENUES, cột Total NIKE Brand",
       definition="NIKE Direct + Bán buôn NIKE Brand + Global Brand Divisions. Không gồm Converse và Corporate.",
       xbrl=_xbrl(REVENUE_CONCEPTS, dims=_NB, section=NOTE_REVENUES + "|" + NOTE_SEGMENTS,
                  row=r"TOTAL REVENUES|Total NIKE Brand")),
    _m(id="converse_revenue", sheet="channels", scope="Converse", unit=USD_M, kind=KIND_GAAP,
       vi="Doanh thu Converse", en="Converse revenues", line="TOTAL REVENUES, cột Converse",
       definition="Doanh thu thương hiệu Converse (gồm bán trực tiếp, bán buôn và doanh thu cấp phép của Converse). Dùng "
                  "để đối chiếu NIKE Brand với NIKE, Inc.",
       xbrl=_xbrl(REVENUE_CONCEPTS, dims={**_SEG, "StatementBusinessSegmentsAxis": "ConverseSegmentMember"},
                  section=NOTE_REVENUES + "|" + NOTE_SEGMENTS, row=r"TOTAL REVENUES|Converse")),
    _m(id="corporate_revenue", sheet="channels", scope="Corporate", unit=USD_M, kind=KIND_GAAP,
       vi="Doanh thu Corporate", en="Corporate revenues", line="TOTAL REVENUES, cột Corporate",
       definition="Chủ yếu là lãi/lỗ phòng ngừa rủi ro tỷ giá liên quan đến doanh thu, có thể âm. Chỉ dùng để đối chiếu.",
       xbrl=_xbrl(REVENUE_CONCEPTS, dims={"ConsolidationItemsAxis": "CorporateNonSegmentMember"},
                  section=NOTE_REVENUES + "|" + NOTE_SEGMENTS, row=r"TOTAL REVENUES|Corporate")),
    _m(id="nike_brand_digital_sales", sheet="channels", scope="NIKE Brand", channel="NIKE Direct – Digital", unit=USD_B,
       kind=KIND_MDNA, vi="Doanh số NIKE Brand Digital", en="Digital commerce sales (Q2–Q3 FY2021: “Digital sales”)",
       line="câu “Digital commerce sales were …” (quý 2–3 FY2021: “Digital sales were …”)",
       definition="Doanh số qua ứng dụng và website của NIKE, là một phần của NIKE Direct. NIKE chỉ nêu số này trong MD&A: "
                  "quý 1–3 FY2019 tính đến triệu USD; số cả năm và các quý từ FY2020 làm tròn đến 0,1 tỷ USD. Quý 4 không "
                  "có 10-Q, nên số quý 4 là ước tính: cả năm trừ 9 tháng đầu.",
       mdna="digital"),
    _m(id="nike_direct_share_stated", sheet="channels", scope="NIKE Brand", channel="NIKE Direct", unit=PCT, kind=KIND_MDNA,
       vi="Tỷ trọng NIKE Direct do NIKE nêu (xấp xỉ)", en="NIKE Direct as approx. % of total NIKE Brand revenues (as stated)",
       line="câu “NIKE Direct revenues represented approximately …%”",
       definition="Tỷ trọng làm tròn do NIKE tự nêu trong MD&A. Dùng để đối chiếu với tỷ trọng tự tính. Quý 4 không có.",
       mdna="share"),
    _m(id="nike_direct_pct_change_reported", sheet="channels", scope="NIKE Brand", channel="NIKE Direct", unit=PCT,
       kind=KIND_MDNA_GAAP, vi="% thay đổi NIKE Direct (NIKE công bố, đã làm tròn)", en="Sales through NIKE Direct – % change",
       line="Sales through NIKE Direct, cột “% CHANGE”",
       definition="Tăng trưởng so với cùng kỳ năm trước, tính theo USD; NIKE làm tròn đến số nguyên. Dùng để đối chiếu.",
       mdna="revenue_table:SALES THROUGH NIKE DIRECT:pct"),
    _m(id="nike_direct_pct_change_cn", sheet="channels", scope="NIKE Brand", channel="NIKE Direct", unit=PCT,
       kind=KIND_NONGAAP, vi="% thay đổi NIKE Direct loại trừ tỷ giá",
       en="Sales through NIKE Direct – % change excluding currency changes",
       line="Sales through NIKE Direct, cột “% CHANGE EXCLUDING CURRENCY CHANGES”",
       definition="Tăng trưởng loại trừ tỷ giá (non-GAAP). Không tự tính lại được từ báo cáo tài chính.",
       mdna="revenue_table:SALES THROUGH NIKE DIRECT:pct_cn"),
    _m(id="nike_brand_wholesale_pct_change_reported", sheet="channels", scope="NIKE Brand", channel="Bán buôn (Wholesale)",
       unit=PCT, kind=KIND_MDNA_GAAP, vi="% thay đổi bán buôn (NIKE công bố, đã làm tròn)",
       en="Sales to Wholesale Customers – % change", line="Sales to Wholesale Customers, cột “% CHANGE”",
       definition="Tăng trưởng so với cùng kỳ năm trước, tính theo USD; NIKE làm tròn đến số nguyên. Dùng để đối chiếu.",
       mdna="revenue_table:SALES TO WHOLESALE CUSTOMERS:pct"),
    _m(id="nike_brand_wholesale_pct_change_cn", sheet="channels", scope="NIKE Brand", channel="Bán buôn (Wholesale)",
       unit=PCT, kind=KIND_NONGAAP, vi="% thay đổi bán buôn loại trừ tỷ giá",
       en="Sales to Wholesale Customers – % change excluding currency changes",
       line="Sales to Wholesale Customers, cột “% CHANGE EXCLUDING CURRENCY CHANGES”",
       definition="Tăng trưởng loại trừ tỷ giá (non-GAAP). Không tự tính lại được từ báo cáo tài chính.",
       mdna="revenue_table:SALES TO WHOLESALE CUSTOMERS:pct_cn"),
    _m(id="nike_brand_pct_change_reported", sheet="channels", scope="NIKE Brand", channel="Tổng NIKE Brand", unit=PCT,
       kind=KIND_MDNA_GAAP, vi="% thay đổi doanh thu NIKE Brand (NIKE công bố, đã làm tròn)",
       en="TOTAL NIKE BRAND REVENUES – % change", line="TOTAL NIKE BRAND REVENUES, cột “% CHANGE”",
       definition="Tăng trưởng so với cùng kỳ năm trước, tính theo USD; NIKE làm tròn đến số nguyên. Dùng để đối chiếu.",
       mdna="revenue_table:TOTAL NIKE BRAND REVENUES:pct"),
    _m(id="nike_brand_pct_change_cn", sheet="channels", scope="NIKE Brand", channel="Tổng NIKE Brand", unit=PCT,
       kind=KIND_NONGAAP, vi="% thay đổi NIKE Brand loại trừ tỷ giá",
       en="TOTAL NIKE BRAND REVENUES – % change excluding currency changes",
       line="TOTAL NIKE BRAND REVENUES, cột “% CHANGE EXCLUDING CURRENCY CHANGES”",
       definition="Tăng trưởng loại trừ tỷ giá (non-GAAP). Không tự tính lại được từ báo cáo tài chính.",
       mdna="revenue_table:TOTAL NIKE BRAND REVENUES:pct_cn"),
]

def mid_sentence(name):
    """A region's name inside a sentence: 'Châu Âu, …' becomes 'châu Âu, …'."""
    return name[0].lower() + name[1:] if name.startswith("Châu") else name


# ---- Geographic operating segments of NIKE Brand ----
for rid, vi, short, en, member in REGIONS:
    dims = {**_NB, "SubsegmentsAxis": member}
    vi = mid_sentence(vi)
    REPORTED += [
        _m(id=f"{rid}_revenue", sheet="regions", scope="NIKE Brand", region=rid, channel="Tổng khu vực", unit=USD_M,
           kind=KIND_GAAP, vi=f"Doanh thu {short}", en=f"{en} revenues", line=f"TOTAL REVENUES, cột {en}",
           definition=f"Doanh thu NIKE Brand tại khu vực {vi} ({en}). Bốn khu vực cộng Global Brand Divisions bằng tổng "
                      f"doanh thu NIKE Brand.",
           xbrl=_xbrl(REVENUE_CONCEPTS, dims=dims, section=NOTE_REVENUES + "|" + NOTE_SEGMENTS,
                      row=rf"TOTAL REVENUES|{en}")),
        _m(id=f"{rid}_direct_revenue", sheet="regions", scope="NIKE Brand", region=rid, channel="NIKE Direct", unit=USD_M,
           kind=KIND_GAAP, vi=f"Doanh thu NIKE Direct – {short}", en=f"{en} – Sales through Direct to Consumer",
           line=f"Sales through Direct to Consumer, cột {en}",
           definition=f"Doanh thu NIKE Direct tại khu vực {vi}.",
           xbrl=_xbrl(REVENUE_CONCEPTS, dims={**dims, _CH: _DTC}, section=NOTE_REVENUES,
                      row=r"Sales through Direct to Consumer")),
        _m(id=f"{rid}_wholesale_revenue", sheet="regions", scope="NIKE Brand", region=rid, channel="Bán buôn (Wholesale)",
           unit=USD_M, kind=KIND_GAAP, vi=f"Doanh thu bán buôn – {short}", en=f"{en} – Sales to Wholesale Customers",
           line=f"Sales to Wholesale Customers, cột {en}",
           definition=f"Doanh thu bán buôn NIKE Brand tại khu vực {vi}.",
           xbrl=_xbrl(REVENUE_CONCEPTS, dims={**dims, _CH: _WHS}, section=NOTE_REVENUES,
                      row=r"Sales to Wholesale Customers")),
        _m(id=f"{rid}_ebit", sheet="regions", scope="NIKE Brand", region=rid, channel="Tổng khu vực", unit=USD_M,
           kind=KIND_GAAP, freq={A}, vi=f"EBIT {short}", en=f"{en} – Earnings before interest and taxes",
           line=f"EARNINGS BEFORE INTEREST AND TAXES, dòng {en}",
           definition=f"Lợi nhuận trước lãi vay và thuế của khu vực {vi}, theo thuyết minh bộ phận. Chưa trừ chi phí NIKE "
                      f"Brand quản lý tập trung (Global Brand Divisions, gồm cả vận hành Digital toàn cầu) và chi phí "
                      f"Corporate, vì NIKE không chia các khoản này về khu vực.",
           xbrl=_xbrl(["EarningsBeforeInterestAndTaxes"], dims=dims, section=NOTE_SEGMENTS, row=en,
                      alt_dims=[{**dims, "ProductOrServiceAxis": "NIKEBrandMember"}])),
    ]

REPORTED_BY_ID = {m["id"]: m for m in REPORTED}
REGION_BY_ID = {r[0]: r for r in REGIONS}


def location(m, form):
    """Where a reader finds the metric in a 10-K or a 10-Q."""
    if m.get("xbrl"):
        sec = m["xbrl"]["section"]
        if sec in (STATEMENT, BALANCE):
            stmt = "Statements of Income" if sec == STATEMENT else "Balance Sheets"
            where = (f"Item 8 – Consolidated {stmt}" if form == "10-K"
                     else f"Item 1 – Unaudited Condensed Consolidated {stmt}")
        elif sec == NOTE_SEGMENTS:
            where = f"{'Item 8' if form == '10-K' else 'Item 1'} – thuyết minh “Operating Segments”"
        else:
            where = f"{'Item 8' if form == '10-K' else 'Item 1'} – thuyết minh “Revenues”, bảng phân tách doanh thu"
        return f"{where}, dòng “{m['line']}”" if sec in (STATEMENT, BALANCE) else f"{where}, {m['line']}"
    item = "Item 7" if form == "10-K" else "Item 2"
    if m["mdna"] == "ebit":
        return f"{item} – MD&A, bảng EBIT theo bộ phận, dòng “{m['line']}”"
    if m["mdna"] in ("digital", "share"):
        return f"{item} – MD&A, đoạn về NIKE Direct, {m['line']}"
    return f"{item} – MD&A, bảng doanh thu, dòng {m['line']}"


def _k(**kw):
    kw.setdefault("freq", {A, Q})
    return dict(kw)


# Calculated KPIs, in evaluation order (a KPI may use KPIs defined above it).
KPIS = [
    # Distribution strategy – NIKE Brand scope
    _k(id="direct_share", group="channels", unit=PCT, vi="Tỷ trọng NIKE Direct trong doanh thu NIKE Brand",
       expr="nike_direct_revenue / nike_brand_revenue",
       note="Mẫu số là tổng doanh thu NIKE Brand, không phải NIKE, Inc.; NIKE cũng chia cho số này khi nêu tỷ trọng."),
    _k(id="wholesale_share", group="channels", unit=PCT, vi="Tỷ trọng bán buôn trong doanh thu NIKE Brand",
       expr="nike_brand_wholesale_revenue / nike_brand_revenue",
       note="NIKE Direct + bán buôn + Global Brand Divisions = 100% doanh thu NIKE Brand."),
    _k(id="gbd_share", group="channels", unit=PCT, vi="Tỷ trọng Global Brand Divisions trong doanh thu NIKE Brand",
       expr="global_brand_divisions_revenue / nike_brand_revenue", note="Phần còn lại, dưới 0,2%."),
    _k(id="direct_growth", group="channels", unit=PCT, vi="Tăng trưởng doanh thu NIKE Direct (theo USD)",
       expr="nike_direct_revenue / nike_direct_revenue__prior - 1",
       note="So với cùng kỳ năm trước (số năm trước in ngay trong báo cáo đó). Chưa loại trừ ảnh hưởng tỷ giá."),
    _k(id="wholesale_growth", group="channels", unit=PCT, vi="Tăng trưởng doanh thu bán buôn NIKE Brand (theo USD)",
       expr="nike_brand_wholesale_revenue / nike_brand_wholesale_revenue__prior - 1",
       note="So với cùng kỳ năm trước (số năm trước in ngay trong báo cáo đó). Chưa loại trừ ảnh hưởng tỷ giá."),
    _k(id="nike_brand_growth", group="channels", unit=PCT, vi="Tăng trưởng doanh thu NIKE Brand (theo USD)",
       expr="nike_brand_revenue / nike_brand_revenue__prior - 1", note="So với cùng kỳ năm trước (số năm trước in ngay trong báo cáo đó)."),
    # The two contributions are shares of the year-earlier NIKE Brand revenue; the workbook shows them in points.
    _k(id="direct_contribution", group="channels", unit=PCT, vi="Đóng góp của NIKE Direct vào tăng trưởng NIKE Brand",
       unit_label="điểm %", scale=100, num_format="+0.0;−0.0;0.0",
       expr="(nike_direct_revenue - nike_direct_revenue__prior) / nike_brand_revenue__prior",
       note="Phần thay đổi doanh thu NIKE Direct chia cho doanh thu NIKE Brand cùng kỳ năm trước. Ba phần đóng góp (NIKE "
            "Direct, bán buôn, Global Brand Divisions) cộng lại bằng tăng trưởng NIKE Brand."),
    _k(id="wholesale_contribution", group="channels", unit=PCT, vi="Đóng góp của bán buôn vào tăng trưởng NIKE Brand",
       unit_label="điểm %", scale=100, num_format="+0.0;−0.0;0.0",
       expr="(nike_brand_wholesale_revenue - nike_brand_wholesale_revenue__prior) / nike_brand_revenue__prior",
       note="Phần thay đổi doanh thu bán buôn chia cho doanh thu NIKE Brand cùng kỳ năm trước."),
    _k(id="digital_share_of_direct", group="channels", unit=PCT, vi="Tỷ trọng Digital trong NIKE Direct (xấp xỉ)",
       num_format="0%", expr="nike_brand_digital_sales * 1000 / nike_direct_revenue",
       note="Xấp xỉ khi số Digital đã làm tròn (số cả năm và các quý từ FY2020). Nhân 1.000 để đổi tỷ USD sang triệu USD."),
    _k(id="direct_ex_digital", group="channels", unit=USD_M, vi="NIKE Direct ngoài Digital (ước tính, chủ yếu là cửa hàng)",
       num_format="#,##0", expr="nike_direct_revenue - nike_brand_digital_sales * 1000",
       note="NIKE Direct trừ doanh số Digital, gần với doanh thu cửa hàng do NIKE sở hữu. Sai số do làm tròn số Digital: "
            "không có ở quý 1–3 FY2019 (NIKE nêu đến triệu USD); tới ±50 triệu USD ở các kỳ khác và quý 4 FY2019; tới "
            "±100 triệu USD ở quý 4 FY2020 và FY2021."),
    # Financial performance – NIKE, Inc. scope
    _k(id="revenue_growth", group="financial", unit=PCT, vi="Tăng trưởng doanh thu NIKE, Inc. (theo USD)",
       expr="revenues / revenues__prior - 1", note="So với cùng kỳ năm trước (số năm trước in ngay trong báo cáo đó)."),
    _k(id="gross_margin", group="financial", unit=PCT, vi="Biên lợi nhuận gộp",
       expr="gross_profit / revenues", note="Lợi nhuận gộp / doanh thu, đều của NIKE, Inc."),
    _k(id="sga_ratio", group="financial", unit=PCT, vi="Tỷ lệ chi phí bán hàng và quản lý / doanh thu",
       expr="total_selling_admin_expense / revenues", note="NIKE không công bố khoản chi phí này theo kênh."),
    _k(id="demand_creation_ratio", group="financial", unit=PCT, vi="Tỷ lệ chi phí tạo nhu cầu / doanh thu",
       expr="demand_creation_expense / revenues", note="Chi phí quảng cáo và quảng bá (một phần của SG&A) so với doanh thu."),
    _k(id="operating_overhead_ratio", group="financial", unit=PCT, vi="Tỷ lệ chi phí vận hành chung / doanh thu",
       expr="operating_overhead_expense / revenues",
       note="Lương, thuê mặt bằng và chi phí hành chính khác (phần còn lại của SG&A) so với doanh thu."),
    _k(id="other_expense_ratio", group="financial", unit=PCT, vi="Tỷ lệ chi phí (thu nhập) khác thuần / doanh thu",
       expr="other_expense_net / revenues",
       note="Biên EBIT = biên lợi nhuận gộp − chi phí tạo nhu cầu − chi phí vận hành chung − khoản này (đều tính theo % "
            "doanh thu)."),
    _k(id="ebit", group="financial", unit=USD_M, vi="EBIT (tự tính theo định nghĩa của NIKE)",
       expr="net_income + income_tax_expense + interest_expense_net",
       note="Chỉ tiêu non-GAAP. NIKE không trình bày dòng “Operating income” mà dùng EBIT. Số tự tính được đối chiếu với "
            "EBIT do NIKE công bố."),
    _k(id="ebit_margin", group="financial", unit=PCT, vi="Biên EBIT",
       expr="ebit / revenues", note="Chỉ tiêu non-GAAP: EBIT / doanh thu NIKE, Inc."),
    _k(id="net_income_margin", group="financial", unit=PCT, vi="Biên lợi nhuận sau thuế",
       expr="net_income / revenues", note="Theo chuẩn mực kế toán Mỹ (GAAP)."),
    _k(id="inventory_growth", group="financial", unit=PCT, vi="Hàng tồn kho so với cùng thời điểm năm trước",
       expr="inventories / inventories__prior - 1",
       note="Năm: số dư 31/5 so với 31/5 năm trước (cùng 10-K). Quý: số dư cuối quý so với cuối cùng quý năm trước, lấy "
            "từ 10-Q cùng quý năm trước, vì 10-Q chỉ in kèm số dư ngày 31/5."),
    _k(id="average_inventory", group="financial", unit=USD_M, vi="Hàng tồn kho bình quân",
       expr="(inventories + inventories__begin) / 2", note="Bình quân số dư đầu kỳ và cuối kỳ."),
    _k(id="days_inventory", group="financial", unit="ngày", vi="Số ngày tồn kho bình quân",
       expr="days_in_period * average_inventory / cost_of_sales",
       note="Số ngày của kỳ × hàng tồn kho bình quân / giá vốn hàng bán của kỳ. Năm tính 365 ngày, quý tính 91,25 ngày "
            "(365/4) để các quý so sánh được với nhau."),
]
for rid, vi, short, en, member in REGIONS:
    KPIS += [
        _k(id=f"{rid}_direct_share", group="regions", unit=PCT, vi=f"Tỷ trọng NIKE Direct – {short}",
           expr=f"{rid}_direct_revenue / {rid}_revenue", note=f"NIKE Direct / doanh thu khu vực {mid_sentence(vi)}."),
        _k(id=f"{rid}_growth", group="regions", unit=PCT, vi=f"Tăng trưởng doanh thu – {short} (theo USD)",
           expr=f"{rid}_revenue / {rid}_revenue__prior - 1", note="So với cùng kỳ năm trước (số năm trước in ngay trong báo cáo đó)."),
        _k(id=f"{rid}_direct_growth", group="regions", unit=PCT, vi=f"Tăng trưởng NIKE Direct – {short} (theo USD)",
           expr=f"{rid}_direct_revenue / {rid}_direct_revenue__prior - 1",
           note="So với cùng kỳ năm trước. Số năm trước của FY2019 lấy từ bảng khu vực trong MD&A."),
        _k(id=f"{rid}_ebit_margin", group="regions", unit=PCT, freq={A}, vi=f"Biên EBIT – {short}",
           expr=f"{rid}_ebit / {rid}_revenue",
           note="EBIT khu vực / doanh thu khu vực. Chưa trừ chi phí Global Brand Divisions và Corporate, vì NIKE không chia "
                "các khoản này về khu vực."),
    ]
KPIS_BY_ID = {k["id"]: k for k in KPIS}

# Reconciliation checks: expression should be 0 within the tolerance (USD million).
CHECKS = [
    dict(id="chk_gross_profit", vi="Doanh thu − Giá vốn − Lợi nhuận gộp", tol=1,
         expr="revenues - cost_of_sales - gross_profit"),
    dict(id="chk_sga", vi="Chi phí tạo nhu cầu + chi phí vận hành chung − tổng SG&A", tol=1,
         expr="demand_creation_expense + operating_overhead_expense - total_selling_admin_expense"),
    dict(id="chk_pretax", vi="Lợi nhuận gộp − SG&A − lãi thuần − chi phí khác thuần − lợi nhuận trước thuế", tol=2,
         expr="gross_profit - total_selling_admin_expense - interest_expense_net - other_expense_net - income_before_taxes"),
    dict(id="chk_net_income", vi="Lợi nhuận trước thuế − thuế − lợi nhuận sau thuế", tol=1,
         expr="income_before_taxes - income_tax_expense - net_income"),
    dict(id="chk_channels", vi="NIKE Direct + bán buôn + Global Brand Divisions − tổng NIKE Brand", tol=1,
         expr="nike_direct_revenue + nike_brand_wholesale_revenue + global_brand_divisions_revenue - nike_brand_revenue"),
    dict(id="chk_segments", vi="NIKE Brand + Converse + Corporate − doanh thu NIKE, Inc.", tol=1,
         expr="nike_brand_revenue + converse_revenue + corporate_revenue - revenues"),
    dict(id="chk_ebit", vi="EBIT tự tính − EBIT NIKE công bố", tol=1, expr="ebit - ebit_reported"),
    dict(id="chk_regions", vi="Bốn khu vực + Global Brand Divisions − tổng NIKE Brand", tol=2,
         expr="na_revenue + emea_revenue + gc_revenue + apla_revenue + global_brand_divisions_revenue - nike_brand_revenue"),
    dict(id="chk_regions_direct", vi="NIKE Direct bốn khu vực − NIKE Direct toàn NIKE Brand", tol=2,
         expr="na_direct_revenue + emea_direct_revenue + gc_direct_revenue + apla_direct_revenue - nike_direct_revenue"),
    dict(id="chk_regions_wholesale", vi="Bán buôn bốn khu vực − bán buôn toàn NIKE Brand", tol=2,
         expr="na_wholesale_revenue + emea_wholesale_revenue + gc_wholesale_revenue + apla_wholesale_revenue"
              " - nike_brand_wholesale_revenue"),
] + [dict(id=f"chk_{rid}_channels", vi=f"NIKE Direct + bán buôn − doanh thu {short}", tol=1,
          expr=f"{rid}_direct_revenue + {rid}_wholesale_revenue - {rid}_revenue") for rid, _, short, _, _ in REGIONS]

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


def applies(item, freq):
    return freq in item.get("freq", {A, Q})
