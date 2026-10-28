# Carvana（CVNA）個股模型 v0.1（命題＋營收線＋收支平衡＋信用風險）
# 規格：docs/plan/個股模型原則_命題加營收分部.md（命題模組 → 官方營收線 → 損益表 → 現金流與融資）。
# 用法：python3 carvana/build_carvana.py raw.xlsx ；再以 LibreOffice headless 重算：
#   soffice --headless --calc --convert-to xlsx --outdir <dir> raw.xlsx
# 金額：百萬美元（US$M）；每台金額：美元；台數：輛。會計年度＝曆年。
# 工作表（順序固定）：README → In → Q → Thesis → PL → CF → Credit → Val → Recon → Sens → Sources
# 環境變數 SENS_OVR＝JSON {In 鍵: 值}：只替換基準值（敏感度建置用，scripts/sens.py）；成品不設。
# 環境變數 CVNA_ROWMAP＝路徑：輸出列號對照（產生 html.json 用）。
import sys, os, json
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.utils import get_column_letter

HERE = os.path.dirname(os.path.abspath(__file__))
OVR = json.loads(os.environ.get("SENS_OVR", "{}"))
F = json.load(open(os.path.join(HERE, "data", "carvana_facts_20261029.json"), encoding="utf-8"))
SENS_F = os.path.join(HERE, "data", "sens_v02.json")
SENS = json.load(open(SENS_F, encoding="utf-8")) if (os.path.exists(SENS_F) and not OVR) else None
VERSION = "0.2"
BUILD_DATE = "2026-10-29"

BLUE = Font(color="0000FF"); BOLD = Font(bold=True); HDR = PatternFill("solid", fgColor="DDEBF7")
KEY = PatternFill("solid", fgColor="FFF2CC"); GREY = Font(color="808080"); RED = Font(color="C00000", bold=True)
wb = openpyxl.Workbook()
SC = ("A", "B", "C")
SCN = {"A": "情境 A（樂觀：沿公司 300 萬台路徑、利得率維持）",
       "B": "情境 B（基準：成長放緩、利得率小降）",
       "C": "情境 C（信用壓力：2027 年次級車貸循環＋二手車價下跌）"}
YRS = list(range(2025, 2036))                      # C＝2025A、D＝2026E（上半年實績＋下半年模型）、E＝2027 … M＝2035
yc = lambda y: get_column_letter(3 + y - 2025)
FC = [y for y in YRS if y >= 2027]
ROWMAP = {}


def hdr(ws, row, labels, c0=1):
    for j, h in enumerate(labels, c0):
        c = ws.cell(row, j, h); c.font = BOLD; c.fill = HDR


def dn(name, ref):
    wb.defined_names[name] = DefinedName(name, attr_text=ref)


QROWS = [
    ("units", "零售台數", "輛"), ("whs_units", "批發台數", "輛"),
    ("rev_ret", "零售車輛銷售", "US$M"), ("rev_whs", "批發銷售與營收", "US$M"), ("rev_oth", "其他銷售與營收", "US$M"), ("rev_tot", "合併營收", "US$M"),
    ("gp_ret", "零售毛利", "US$M"), ("gp_whs", "批發毛利", "US$M"), ("sga", "SG&A", "US$M"), ("oi", "營業利益", "US$M"),
    ("da", "折舊攤銷（銷貨成本＋SG&A）", "US$M"), ("sbc", "股份基礎給付", "US$M"), ("orig", "貸款發放", "US$M"), ("gos", "貸款出售利得", "US$M"),
    ("warrant", "權證營收", "US$M"), ("adj_ebitda", "調整後 EBITDA", "US$M"), ("int", "利息費用（淨）", "US$M"), ("oth_exp", "其他費用（收益）與清償損失（推得）", "US$M"),
    ("tax", "所得稅（利益）", "US$M"), ("ni", "淨利（合併）", "US$M"), ("rp_oth", "其他營收中關係人部分（DriveTime 等）", "US$M"),
    ("capex", "資本支出", "US$M"), ("cfo", "營業活動現金流", "US$M"),
    ("cash", "現金（期末）", "US$M"), ("inv", "存貨（期末）", "US$M"), ("frhfs", "待售貸款（期末）", "US$M"), ("bi", "證券化受益權（期末）", "US$M"),
    ("debt_bs", "債務合計（含融資租賃，期末）", "US$M"), ("equity", "股東權益（含非控制權益，期末）", "US$M"), ("liq_total", "公司口徑總流動性資源（期末）", "US$M"),
]
QR = {}
_rr = 4
for _k, _l, _u in QROWS:
    QR[_k] = _rr; _rr += 1
q = lambda k, col: f"Q!{col}{QR[k]}"
# =====================================================================================
# In
# =====================================================================================
IN = wb.active; IN.title = "In"
IN["A1"] = "In — 輸入（藍字＝輸入；黑字＝公式；基準欄 E，低 D、高 F）。金額：百萬美元（US$M）；每台金額：美元；台數：輛"
IN["A2"] = "標記：Verified／Interested-party／Analogy／Assumed／Derived；Analogy、Assumed 給低／高區間。公司財報數字依專案慣例標 Interested-party（公司自述、經會計師核閱或查核）"
hdr(IN, 3, ["鍵", "項目", "單位", "低", "基準", "高", "標記", "來源1", "來源2／說明"])
rows = []; sec = lambda t: rows.append(("sec", t)); inp = lambda *a: rows.append(a)
PR2 = "Carvana Q2 2026 股東信（8-K 2026-07-29 附件 99.1）"
PR3 = "Carvana Q3 2026 股東信（8-K 2026-10-28 附件 99.1）"
Q103 = "Carvana 10-Q Q3 2026（2026-10-28）"
Q10 = "Carvana 10-Q Q2 2026（2026-07-29）"
ABS = "Carvana Auto Receivables Trust 2026-P3 公開說明書 424B5（2026-08-20）Historical Performance（Bridgecrest 服務的管理組合）"
FACTS = "carvana/data/carvana_facts_20261029.json"

sec("A. 股價、股數")
inp("px", "現價（2026-10-28 收盤，推得）", "US$", None, 59.03, None, "Derived", "StockStory（2026-10-28）：財報公布後上漲 8.4% 至 63.99 → 收盤約 63.99 ÷ 1.084＝59.03", "盤後 63.99；2026-10-26 收盤 56.61（stockanalysis）")
inp("sh_fx", "完全轉換股數（LLC 單位全數交換＋其他稀釋）", "百萬股", None, 1127, None, "Interested-party", PR2 + " 第 1 頁：Assuming full conversion … 1.127 billion shares", "A 類流通 719.0 百萬股＋B 類 380.5 百萬股（B 類為表決權股；經濟權益以 LLC 單位交換計）")

sec("B. 命題：零售台數路徑")
inp("u_q4", "2026Q4 零售台數", "輛", None, "=lq_units*(1+q4_seq)", None, "Derived", "Q3 2026 實績 217,385 ×（1＋季增）", "公司指引 Q4 較 Q3 增加（Q3 2026 股東信）")
inp("q4_adj", "2026Q4 每台總毛利季節性調整（相對 Q3）", "US$", -935, -600, -300, "Derived", "Q4 2025 總 GPU 6,427 較 Q3 2025 7,362 低 935（零售 −466、其他 −233、批發 −236；Q 頁）", "季節性；以零售毛利承載")
inp("q4_seq", "2026Q4 台數相對 Q3 變化", "%", 0.0, 0.02, 0.05, "Assumed", "公司指引：sequential increase（Q3 2026 股東信）；2025 Q4／Q3＝+4.9%", "—")
for s, g27, g28, g30, g35 in (("A", 0.32, 0.28, 0.20, 0.08), ("B", 0.25, 0.20, 0.13, 0.05), ("C", -0.10, 0.05, 0.10, 0.04)):
    note = {"A": "公司目標 2030–2035 年 300 萬台（股東信）；此路徑約 2033 年觸及用地上限",
            "B": "2025 年 +43%、2026 上半年 +39%、Q3 共識 +35%，逐年放緩；2035 年約為二手零售市場 6%",
            "C": "2027 年次級車貸循環：放款收緊、二手車價下跌，台數 −10%；2028 年起緩步恢復"}[s]
    inp(f"g27_{s}", f"情境 {s}：2027 年零售台數年增", "%", g27 - 0.08, g27, g27 + 0.08, "Assumed", note, "—")
    inp(f"g28_{s}", f"情境 {s}：2028 年零售台數年增", "%", g28 - 0.08, g28, g28 + 0.08, "Assumed", "同上", "—")
    inp(f"g30_{s}", f"情境 {s}：2030 年零售台數年增（2029 年取 2028 與 2030 中點）", "%", g30 - 0.05, g30, g30 + 0.05, "Assumed", "同上", "—")
    inp(f"g35_{s}", f"情境 {s}：2035 年零售台數年增（2031–2034 線性）", "%", g35 - 0.03, g35, g35 + 0.03, "Assumed", "同上", "—")
inp("unit_cap", "生產用地可支援的年零售台數上限", "輛", None, 3000000, None, "Interested-party", PR2 + " 第 3 頁：real estate to support annual retail production of 3 million", "台數路徑以此為上限")
inp("cap_0", "已完成建置的年產能（2026 年中）", "輛", None, 1500000, None, "Interested-party", PR2 + " 第 3 頁：fully built out annual capacity for approximately 1.5 million", "—")
inp("util", "產能利用率上限（超過即需擴建）", "%", 0.75, 0.85, 0.95, "Assumed", "公司未揭露（未找到）", "—")
inp("capcost", "每台年產能建置成本", "US$／台", 400, 600, 1000, "Assumed", "ADESA 站點全面改建成本：未找到（公司只揭露首個全面改建 2027 年初投產）", "成長性資本支出＝新增產能 × 每台成本")

sec("C. 命題：單位經濟（每零售台）")
inp("asp_q4", "2026Q4 零售每台營收", "US$", 27500, 28263, 28800, "Interested-party", PR3 + "：Q3 2026 零售每台營收 28,263", "—")
inp("asp_g", "零售每台售價年增（2027 起）", "%", 0.0, 0.02, 0.04, "Assumed", "二手車價（Manheim 指數 2026-09 為 205.9，年減 0.6%，Cox Automotive）", "貸款發放額、批發營收同幅")
inp("whs_pu", "批發營收 ÷ 零售台數（2026Q4 起）", "US$", None, f"={q('rev_whs','K')}*1000000/{q('units','K')}", None, "Derived", "2026 前三季批發營收 3,723 ÷ 零售 602,103 台（Q 頁）", "Q3 單季 5,989")
inp("gpr_q4", "2026Q4 零售毛利／台（季節調整前）", "US$", 3165, 3556, 3600, "Interested-party", PR3 + "：Q3 2026 零售 GPU 3,556（共識 3,404）", "低＝Q1 2026 實績")
for s, v27, vlt in (("A", 3600, 3700), ("B", 3350, 3300), ("C", 2700, 3100)):
    inp(f"gpr27_{s}", f"情境 {s}：2027 年零售 GPU", "US$", v27 - 300, v27, v27 + 300, "Assumed", "2024 年 3,312、2025 年 3,315、2026 上半年 3,361（Q 頁）" + ("；情境 C：二手車價下跌時零售毛利受壓（2022 年循環曾大幅下滑）" if s == "C" else ""), "2028–2029 線性收斂至長期值")
    inp(f"gprLT_{s}", f"情境 {s}：2030 年起零售 GPU", "US$", vlt - 300, vlt, vlt + 300, "Assumed", "同上", "—")
inp("gpw_h2", "2026Q4 起批發毛利／零售台", "US$", 650, 708, 806, "Interested-party", PR3 + "：Q3 2026 批發 GPU 708（共識 812）", "高＝2026 上半年 806")
inp("gpw27_C", "情境 C：2027 年批發毛利／零售台", "US$", 550, 650, 750, "Assumed", "批發車價下跌時批發毛利下降", "情境 A、B 沿用 gpw_h2")
inp("opu", "每零售台貸款發放金額（2026Q4）", "US$", None, f"={q('orig','J')}*1000000/{q('units','J')}", None, "Derived", "Q3 2026 發放 5,182（前三季 14,187 − 上半年 9,005）÷ 217,385 台（Q 頁）", "2027 起隨售價年增")
inp("gos_h2", "2026Q4 貸款出售利得率（利得 ÷ 發放額）", "%", 0.065, f"={q('gos','J')}/{q('orig','J')}", 0.08, "Derived", "Q3 2026 利得 381（前三季 1,084 − 上半年 703）÷ 發放 5,182（Q 頁）", "Q2 2026 7.35%；公司稱 Q3 受基準利率快速上升與油價影響每台約 −200")
for s, v27, vlt in (("A", 0.08, 0.08), ("B", 0.072, 0.07)):
    inp(f"gos27_{s}", f"情境 {s}：2027 年貸款出售利得率", "%", v27 - 0.01, v27, v27 + 0.01, "Assumed", "2025 年 9.3%、2026 上半年 7.8%、Q2 7.3%（Q 頁）", "2028–2029 線性收斂至長期值")
    inp(f"gosLT_{s}", f"情境 {s}：2030 年起貸款出售利得率", "%", vlt - 0.01, vlt, vlt + 0.01, "Assumed", "同上", "—")
inp("gos27_C", "情境 C：2027 年貸款出售利得率", "%", None, "=gos27_B-loss_pt*dcnl_C-dur*dspr_C", None, "Derived", "情境 B 2027 年 − 損失傳導 × 終身損失增幅 − 存續期間 × 利差擴大（Credit 頁 S2）", "—")
inp("gosLT_C", "情境 C：2030 年起貸款出售利得率", "%", 0.05, 0.06, 0.07, "Assumed", "循環後部分回復", "—")
inp("oth_pu", "其他毛利中非貸款部分／零售台（VSC、GAP、權證等）", "US$", None, f"=({q('rev_oth','J')}-{q('gos','J')})*1000000/{q('units','J')}", None, "Derived", "Q3 2026（其他毛利 594 − 出售利得 381）÷ 217,385 台（Q 頁）", "VSC 佣金來自 DriveTime（關係人）")
inp("oth_g", "非貸款其他毛利／台年增", "%", -0.03, 0.0, 0.03, "Assumed", "—", "—")

sec("D. 費用、折舊、資本支出")
inp("sga_fix", "SG&A 固定部分（每季）", "US$M", 140, 170, 200, "Derived", "季度 SG&A 對零售台數迴歸（Q 頁）", "Q3 2026 驗算：170＋217,385 × 2,900＝800（實績 805）")
inp("sga_var", "SG&A 變動部分／台（2026 年）", "US$", 2700, 2900, 3100, "Derived", "固定 170／季時，Q3 2026 SG&A 805 ÷ 217,385 台反推約 2,920；Q3 廣告 152、物流 68（油價）偏高（Q3 股東信 SG&A 拆分）", "v0.1 為 2,700")
inp("infl", "固定費用、折舊、SBC 年增", "%", 0.02, 0.03, 0.04, "Assumed", "—", "—")
for s, e in (("A", 0.03), ("B", 0.015), ("C", 0.0)):
    inp(f"eff_{s}", f"情境 {s}：變動 SG&A／台每年下降", "%", max(e - 0.015, 0), e, e + 0.015, "Assumed", "公司稱已承擔可支援更大量的固定成本、規模可再擴 EBITDA 率（Q1 2025 股東信，Interested-party）", "—")
inp("da_q4", "2026Q4 折舊攤銷", "US$M", None, f"={q('da','J')}", None, "Derived", "沿用 Q3 2026 的 64（Q 頁）", "2027 起隨 infl 成長＋成長性資本支出 ÷ 年限")
inp("life", "成長性資本支出折舊年限", "年", 10, 15, 20, "Assumed", "公司未揭露（建物與改裝設備混合）", "—")
inp("sbc_q4", "2026Q4 SBC", "US$M", None, f"={q('sbc','J')}", None, "Derived", "沿用 Q3 2026 的 25（Q 頁）", "—")
inp("warr", "權證營收（Root，自調整後 EBITDA 扣除）", "US$M／年", None, 21, None, "Interested-party", "2025 全年 21（Q4 2025 股東信非 GAAP 對帳）", "—")
inp("capex_q4", "2026Q4 資本支出", "US$M", 60, 70, 100, "Assumed", "Q3 2026 為 64；三個 ADESA 全面改建施工中（Q3 股東信）", "—")
inp("mcx", "維護性資本支出 ÷ 營收", "%", 0.003, 0.005, 0.008, "Assumed", "2025 年總資本支出 147 ÷ 營收 20,322＝0.72%（含成長性）", "—")

sec("E. 營運資金與風險自留")
inp("inv_d", "存貨天數（÷ 零售銷貨成本）", "天", 55, 60, 65, "Derived", "2025-12-31：2,408 ÷（Q4 零售成本 3,667 × 4）× 365＝60；2026-06-30＝62（Q 頁）", "—")
inp("inv_d_C", "情境 C：2027 年存貨天數", "天", 65, 75, 85, "Assumed", "需求下滑時週轉變慢", "2028 起回到 inv_d")
inp("fr_d", "待售貸款天數（÷ 年發放額）", "天", 14, 16, 20, "Derived", "2026-09-30：896 ÷（Q3 發放 5,182 × 4）× 365＝15.8（Q 頁）", "v0.1 為 18")
inp("ap_pct", "應付與應計 ÷（銷貨成本＋SG&A）年額", "%", None, 0.054, None, "Derived", "2026-09-30：1,574 ÷（Q3 銷貨成本 6,519＋SG&A 805）× 4＝5.4%", "—")
inp("abs_sh", "證券化占貸款出售比例", "%", None, "=3.7/(5.7+3.7+4.5)", None, "Derived", Q103 + " 附註 7：2026 前三季證券化 3.7B ÷（Ally 5.7＋證券化 3.7＋整批出售 4.5）", "—")
inp("bi_ret", "證券化風險自留比例", "%", None, 0.05, None, "Verified", "Regulation RR（10-Q 附註 8：at least 5%）", "—")
inp("bi_adv", "自留受益權的融資成數", "%", None, "=427/557", None, "Derived", Q103 + " 附註 8、9：受益權融資 427 ÷ 受益權 557", "公允價值下跌時須補擔保")

sec("F. 稅與 TRA（應收稅款協議）")
inp("etr_26", "2026Q4 有效稅率", "%", None, f"={q('tax','K')}/({q('oi','K')}-{q('int','K')}-{q('oth_exp','K')})", None, "Derived", "2026 前三季所得稅 171 ÷ 稅前 1,572", "—")
inp("etr_27", "2027 年有效稅率", "%", 0.08, 0.12, 0.16, "Assumed", "非控制權益（LLC 單位）部分不在 Carvana Co. 課稅；交換越多稅率越高", "2028–2029 線性")
inp("etr_30", "2030 年起有效稅率", "%", 0.12, 0.18, 0.25, "Assumed", "同上", "—")
inp("ctr", "現金稅率（2026–2028，遞延所得稅資產抵用期）", "%", 0.0, 0.03, 0.08, "Assumed", "遞延所得稅資產 2,905（2026-09-30）；舊註 2,968（2026-06-30）；2026 上半年付稅 8（XBRL IncomeTaxesPaid）", "2029 起現金稅＝帳面稅")
inp("tra_0", "TRA 負債（2026-09-30）", "US$M", None, 2130, None, "Interested-party", Q103 + " 資產負債表：2,130，其中關係人 1,645", "TRA：公司須支付已實現節稅的 85%")
inp("tra_q4", "2026Q4 TRA 付款", "US$M", 0, 0, 40, "Derived", "2026 前三季付款 37 全在上半年（現金流量表）", "—")
inp("tra_yrs", "TRA 平均償付年數（2027 起平均攤付）", "年", 10, 15, 20, "Assumed", "公司未揭露償付時程（未找到）；實際付款隨課稅所得與交換而變", "之後新增交換會再增加 TRA 負債（本版未建模）")
inp("ncitax", "代非控制權益繳稅", "US$M／年", None, 8, None, "Derived", "2026 上半年 4（現金流量表）", "—")

sec("G. 債務與利率（2026-09-30 結構；2030 年擔保票據已於 2026-08 以定期貸款 B 贖回）")
inp("sofr", "SOFR（2026-10-26）", "%", None, 0.0393, None, "Verified", "sofrrate.com 2026-10 每日表：2026-10-26 為 3.93%", "各年持平（判斷）")
inp("tlb", "定期貸款 B（2033-08-14 到期）", "US$M", None, 1660, None, "Interested-party", Q103 + " 附註 9：1,660（OID 4、發行成本 17）", "每季攤還原始本金 0.25%（2027 起）")
inp("swap_fix", "利率交換固定利率（付固定、收 1 個月 Term SOFR；名目 1,660，至 2028-07-31）", "%", None, 0.04055, None, "Interested-party", Q103 + " 附註 17、第 3 項：4.055%", "交換期間定期貸款 B 實質固定在 6.305%；之後隨 SOFR 浮動")
inp("tlb_sp", "定期貸款 B 利差（SOFR＋）", "%", None, 0.0225, None, "Interested-party", "同上：Term SOFR＋2.25%", "公司稱每年省利息約 45（2026-08-12 新聞稿）")
inp("ssn_31", "2031 年到期擔保票據", "US$M", None, 2269, None, "Interested-party", Q10 + " 附註 9", "2028-08-15 前僅能以補償溢價贖回")
inp("ssn_r", "擔保票據現金利率", "%", None, 0.09, None, "Interested-party", Q10 + " 附註 9：9%（PIK 已於 2025-08 結束）", "—")
inp("refi_yr", "情境 A、B：2031 年票據再融資年度", "年", None, 2028, None, "Assumed", "2028-08-15 起可依贖回價格贖回（10-Q 附註 9）", "情境 C 於 2031 年到期時再融資")
inp("refi_sp", "情境 A、B：再融資利差（SOFR＋）", "%", 0.0225, 0.0275, 0.035, "Assumed", "比照 2026 年定期貸款 B 2.25%", "—")
inp("refi_prem", "2031 年票據提前贖回溢價", "%", 0.0225, 0.045, 0.045, "Assumed", "贖回價格表：未找到（契約未在本次取得的文件中）；高收益債慣例首個贖回日為票面＋半年息", "—")
inp("stress_sp", "情境 C：2031 年到期再融資利差（SOFR＋）", "%", 0.04, 0.05, 0.07, "Assumed", "信用循環中的高收益擔保債利差（判斷）", "—")
inp("uns", "無擔保票據（2027–2030 到期）", "US$M", None, 107, None, "Interested-party", Q10 + " 附註 9：2027 年 32、2028 年 22、2029 年 26、2030 年 27", "—")
inp("uns_r", "無擔保票據加權利率", "%", None, "=(32*0.055+22*0.05875+26*0.04875+27*0.1025)/107", None, "Derived", Q10 + " 附註 9 各期利率", "—")
inp("re_fin", "不動產售後租回融資", "US$M", None, 485, None, "Interested-party", Q10 + " 附註 9", "期限 20–25 年；本版不攤還")
inp("re_r", "不動產融資利率", "%", 0.07, 0.08, 0.09, "Assumed", "隱含利率：未找到", "—")
inp("bi_r", "受益權融資利率", "%", 0.06, 0.07, 0.08, "Assumed", "未找到", "—")
inp("fleet", "運輸車隊融資", "US$M", None, 142, None, "Interested-party", Q103 + " 附註 9（Citizens Bank，額度 250）", "本版持平")
inp("fleet_r", "運輸車隊融資利率", "%", None, 0.065, None, "Derived", "SOFR 交換利率約 3.6%＋2.80–2.95%（10-Q 附註 9）", "—")
inp("finls", "融資租賃負債", "US$M", None, 118, None, "Interested-party", Q10 + "（XBRL FinanceLeaseLiability）", "本版持平")
inp("finls_r", "融資租賃折現率", "%", None, 0.064, None, "Interested-party", Q10 + "：加權平均折現率 6.4%", "—")
inp("fp_ratio", "存貨融資（Ally floor plan）÷ 存貨", "%", None, "=71/3456", None, "Derived", Q103 + "：floor plan 71 ÷ 存貨 3,456", "額度 1.5B 至 2027-04-30")
inp("fp_r", "存貨融資利率", "%", None, 0.0444, None, "Interested-party", Q10 + " 附註 9：Q2 2026 有效利率 4.44%", "—")
inp("cash_y", "現金利息收益率", "%", 0.03, 0.035, 0.04, "Assumed", "SOFR 3.93% 減貨幣基金費用（判斷）", "利息費用為淨額口徑")
inp("int_lq", "Q3 2026 利息費用（淨）", "US$M", None, 98, None, "Interested-party", PR3, "Q3 含 2030 年票據 1.5 個月（9%）；校準時已扣除此部分")
inp("refi_save", "定期貸款 B 每年利息節省（公司說法）", "US$M", None, 45, None, "Interested-party", "Carvana 新聞稿 2026-08-12", "僅供對照")
inp("redeem_cost", "2030 年票據贖回溢價（2026 年已付）", "US$M", None, 50, None, "Interested-party", Q103 + " 附註 9：以 1,711 贖回本金 1,660（溢價 50、利息 1）；清償損失 53", "v0.1 假設 50")
inp("cash_cap", "現金上限（超過部分視為回饋股東或提前還債）", "US$M", 3000, 5000, 10000, "Assumed", "公司未揭露資本配置政策（未找到）；避免現金無限累積造成利息收入膨脹", "不影響收支關卡")
inp("min_cash", "最低營運現金", "US$M", 750, 1000, 1500, "Assumed", "—", "低於此即視為需外部融資")
inp("new_r_sp", "新增外部融資利差（SOFR＋）", "%", None, 0.05, None, "Assumed", "收支缺口以高收益擔保債補足（判斷）", "—")

sec("H. 貸款去化（承接）能力")
inp("ally_cap", "Ally 遠期承購承諾（2026-07-27 至 2027-07-26）", "US$M／年", None, 8000, None, "Interested-party", Q103 + " 附註 7：8.0B；2026-09-30 未用 6.0B", "每年續約一次")
inp("fp_unused", "整批出售協議未用額度（至 2027 年第四季）", "US$M", None, 8300, None, "Interested-party", Q103 + " 附註 7：未用 8.3B（2026-06-30 為 8.7B）", "換算每 12 個月：÷ fp_yrs")
inp("fp_yrs", "整批出售協議剩餘年數（2026-09-30 至 2027 年底）", "年", None, 1.25, None, "Derived", "15 個月", "—")
inp("abs_rr", "證券化年化量", "US$M／年", None, "=3700*4/3", None, "Derived", Q103 + " 附註 7：2026 前三季 3.7B ×（4／3）", "非承諾；依 ABS 市場")
inp("wh_unused", "倉儲（Finance Receivable Facilities）未用額度", "US$M", None, 4100, None, "Interested-party", Q103 + " 附註 9：未用 4.1B；到期日 2026-10 至 2027-11", "—")
inp("wh_adv", "倉儲額度墊款成數", "%", 0.75, 0.85, 0.9, "Assumed", "未找到", "—")

sec("I. 消費信用壓力（情境 C 的利得率推導）")
inp("dcnl_C", "情境 C：新放款終身累計淨損失增幅", "pt", 0.01, 0.02, 0.03, "Assumed", "對照：管理組合淨沖銷率 2021 年 1.46% → 2024 年 4.49%（三年 +3.0 pt，Credit 頁 S1）", "—")
inp("loss_pt", "損失增幅傳導到出售價格的比例", "x", 0.7, 0.9, 1.0, "Assumed", "買方以預期損失扣價；回收已在淨額中", "—")
inp("dspr_C", "情境 C：ABS／整批買方要求利差擴大", "pt", 0.005, 0.01, 0.02, "Assumed", "判斷", "—")
inp("dur", "出售貸款的價格存續期間", "年", 1.8, 2.0, 2.2, "Derived", "2026-P3 貸款池（APR 10.77%、剩餘 74 個月、發行人定價速度 ABS 1.30%）模擬：WAL 2.25–2.31 年、修正存續期間 1.98–2.02 年；ABS 1.1–1.5% 時 1.84–2.16", "執行長稱每 1bp 每台約 4–5 美元（Q2 2026 法說）；模型隱含約 4.8")

sec("J. 估值")
inp("rf", "無風險利率（美國 10 年期公債，2026-10-26）", "%", None, 0.05, None, "Verified", "沿用 company-models 的 Oracle v0.4、Amazon v0.3 同一輸入（美國財政部 2026-10-26：5.00%）", "跨模型一致")
inp("beta_raw", "原始 Beta（5 年）", "x", None, 3.55, None, "Interested-party", "stockanalysis.com 統計頁（2026-10-27 擷取）", "含 2022 年股價崩跌期")
inp("beta", "Beta（Blume 調整：0.67 × 原始＋0.33）", "x", 1.8, "=0.67*beta_raw+0.33", 3.0, "Derived", "比照 Oracle v0.4 修正一的方法", "低值 1.8 為判斷（成長型汽車零售）")
inp("erp", "股權風險溢酬", "%", None, 0.05, None, "Assumed", "沿用 company-models 慣例", "—")
inp("kd", "稅前債務成本", "%", None, "=(tlb*(sofr+tlb_sp)+ssn_31*ssn_r+uns*uns_r+re_fin*re_r)/(tlb+ssn_31+uns+re_fin)", None, "Derived", "主要長期債務加權（2026-08 再融資後）", "—")
inp("t_val", "估值用稅率（利息稅盾）", "%", None, 0.21, None, "Assumed", "美國聯邦稅率", "—")
inp("g_tv", "永續成長率", "%", 0.02, 0.03, 0.04, "Assumed", "—", "—")
inp("peer_lad", "同業 EV/EBITDA：Lithia（LAD）", "x", None, 11.79, None, "Verified", "stockanalysis.com（2026-10-25 擷取，近四季）", "—")
inp("peer_an", "同業 EV/EBITDA：AutoNation（AN）", "x", None, 10.97, None, "Verified", "stockanalysis.com（2026-10-23 擷取，近四季）", "CarMax 22.51x 不採用：其 EV 含自營車貸（CAF）無追索權債務，口徑不同")
inp("prem", "成長溢價倍數（相對同業中位數）", "x", 1.0, 1.5, 2.0, "Assumed", "Carvana 台數成長約為同業數倍（判斷）；市場現值 EV／近四季調整後 EBITDA 約 25x", "Analogy＋Assumed")
inp("w_dcf", "DCF 權重（其餘為倍數法）", "%", None, 0.5, None, "Assumed", "沿用個股模型原則第 4 條預設", "—")
inp("p_A", "情境 A 機率", "%", None, 0.25, None, "Assumed", "沿用 company-models 慣例 25／50／25", "—")
inp("p_B", "情境 B 機率", "%", None, 0.5, None, "Assumed", "同上", "—")
inp("p_C", "情境 C 機率", "%", None, 0.25, None, "Assumed", "同上", "—")

sec("K. 對帳用外部數字")
inp("guide_lo", "公司 2026 年調整後 EBITDA 指引下限", "US$M", None, 2700, None, "Interested-party", PR3 + "：2.7–3.0B 的 higher end", "公司自 Q4 2026 起不再提供季度與年度指引")
inp("guide_hi", "公司 2026 年調整後 EBITDA 指引上限", "US$M", None, 3000, None, "Interested-party", PR3, "—")
inp("om_tgt", "公司中期目標：營業利益率（300 萬台，2030–2035）", "%", None, 0.125, None, "Interested-party", PR3 + "：2027 起改以 GAAP 目標 12.5% 營業利益率（等同原 13.5% 調整後 EBITDA 率）", "—")
inp("cons_q3_units", "Q3 2026 零售台數共識", "輛", None, 210299, None, "Interested-party", "Zacks（Yahoo Finance 2026-10-23）", "—")
inp("cons_q3_ebitda", "Q3 2026 調整後 EBITDA 共識", "US$M", None, 791.3, None, "Interested-party", "StockStory（2026-10-28）", "—")
inp("cons_q3_gpr", "Q3 2026 零售 GPU 共識", "US$", None, 3404, None, "Interested-party", "Zacks", "—")
inp("cons_q3_gpo", "Q3 2026 其他 GPU 共識", "US$", None, 2722, None, "Interested-party", "Zacks", "—")
inp("cons_q3_gpw", "Q3 2026 批發 GPU 共識", "US$", None, 812, None, "Interested-party", "Zacks", "—")
inp("cons_q3_rev", "Q3 2026 營收共識", "US$M", None, 7670, None, "Interested-party", "Zacks（Yahoo Finance 2026-10-23）", "—")
inp("an_tgt", "分析師平均目標價（21 位）", "US$", None, 82, None, "Interested-party", "TIKR（2026-10-13）：平均約 82、最低 60", "賣方分析師")
inp("moody_lev", "Moody's 升評門檻：債務／EBITDA 低於", "x", None, 4.0, None, "Interested-party", "Moody's 2025-10-02 升 B2、展望正向（Investing.com 報導）", "2026 年評等動作：未找到")
inp("moody_cov", "Moody's 升評門檻：（EBITDA − 資本支出）／利息高於", "x", None, 3.0, None, "Interested-party", "同上", "—")
inp("sp_lev", "S&P 正向門檻：債務／EBITDA 低於", "x", None, 5.0, None, "Interested-party", "S&P 2025-03 升 B、展望正向（Investing.com 報導）", "—")
inp("sp_focf", "S&P 正向門檻：FOCF／債務高於", "%", None, 0.05, None, "Interested-party", "同上", "—")

r = 4
for it in rows:
    if it[0] == "sec":
        IN.cell(r, 1, it[1]).font = BOLD; r += 1; continue
    k, lab, unit, lo, base, hi, tag, s1, s2 = it
    if k in OVR:
        base = OVR[k]
    IN.cell(r, 1, k); IN.cell(r, 2, lab); IN.cell(r, 3, unit)
    lo = round(lo, 6) if isinstance(lo, float) else lo
    hi = round(hi, 6) if isinstance(hi, float) else hi
    for col, v in ((4, lo), (5, base), (6, hi)):
        if v is None:
            continue
        c = IN.cell(r, col, v)
        if not (isinstance(v, str) and v.startswith("=")):
            c.font = BLUE
    IN.cell(r, 7, tag); IN.cell(r, 8, s1); IN.cell(r, 9, s2)
    dn(k, f"In!$E${r}")
    if lo is not None:
        dn(k + "_lo", f"In!$D${r}")
    if hi is not None:
        dn(k + "_hi", f"In!$F${r}")
    ROWMAP.setdefault("In", {})[k] = r
    r += 1
IN.column_dimensions["A"].width = 12; IN.column_dimensions["B"].width = 46; IN.column_dimensions["H"].width = 60; IN.column_dimensions["I"].width = 50

# =====================================================================================
# Q — 事實總帳（季度＋年度）
# =====================================================================================
Q = wb.create_sheet("Q")
Q["A1"] = "Q — 事實總帳（US$M；台數為輛）。來源：各季股東信（8-K 附件 99.1）、10-Q、10-K、SEC XBRL；" + FACTS
Q["A2"] = "欄：C–J＝Q4'24、Q1'25、Q2'25、Q3'25、Q4'25、Q1'26、Q2'26、Q3'26；K＝2026 前三季；L＝FY2024；M＝FY2025。核對列應為 0"
hdr(Q, 3, ["項目", "單位"] + F["quarters"] + ["2026 前三季", "FY2024", "FY2025"])
qd = F["q"]; fy = F["fy"]; ytd = F["ytd_2026"]
# 列號固定（Thesis／In 以列號引用）：4 units … 由清單順序決定；以下 QR 記錄
rr = 4
for key, lab, unit in QROWS:
    Q.cell(rr, 1, lab); Q.cell(rr, 2, unit); rr += 1
# 另加計算列
for key, lab, unit in (("gp_oth", "其他毛利（＝其他營收）", "US$M"), ("gp_tot", "毛利合計", "US$M"), ("chk_rev", "核對：營收線合計 − 合併營收", "US$M"),
                       ("chk_ebitda", "核對：營業利益＋其他營業費用＋折舊＋SBC − 權證 − 調整後 EBITDA", "US$M"), ("oth_opx", "其他營業費用（推得）", "US$M"),
                       ("gos_m", "貸款出售利得率", "%"), ("gpu", "總毛利／零售台", "US$"), ("gpu_ret", "零售毛利／台", "US$"), ("gpu_oth", "其他毛利／台", "US$"),
                       ("sga_pu", "SG&A／台", "US$"), ("ebitda_m", "調整後 EBITDA 率", "%"), ("gos_sh", "出售利得 ÷ 調整後 EBITDA", "%"),
                       ("gosrp_sh", "（出售利得＋關係人其他營收）÷ 調整後 EBITDA", "%"), ("ebitda_xg", "調整後 EBITDA 扣除出售利得", "US$M"), ("ebitda_xg_m", "扣除出售利得的 EBITDA 率", "%")):
    QR[key] = rr; Q.cell(rr, 1, lab); Q.cell(rr, 2, unit); rr += 1
qcols = [get_column_letter(3 + i) for i in range(8)]
for i, col in enumerate(qcols):
    for key, _, _ in QROWS:
        if key == "oth_exp":
            Q[f"{col}{QR[key]}"] = f"={col}{QR['oi']}-{col}{QR['int']}-{col}{QR['tax']}-{col}{QR['ni']}"; continue
        Q[f"{col}{QR[key]}"] = qd[key][i]; Q[f"{col}{QR[key]}"].font = BLUE
# 2026H1（J）、FY2024（K）、FY2025（L）
H1MAP = dict(ytd)
H1MAP.update({"cash": 3049, "inv": 3456, "frhfs": 896, "bi": 557, "debt_bs": 5228, "equity": 5653, "liq_total": 7503})
for key, _, _ in QROWS:
    if key == "oth_exp":
        for col in "KLM":
            Q[f"{col}{QR[key]}"] = f"={col}{QR['oi']}-{col}{QR['int']}-{col}{QR['tax']}-{col}{QR['ni']}"
        continue
    if key in H1MAP:
        Q[f"K{QR[key]}"] = H1MAP[key]; Q[f"K{QR[key]}"].font = BLUE
    if key in fy:
        for j, col in enumerate("LM"):
            Q[f"{col}{QR[key]}"] = fy[key][j]; Q[f"{col}{QR[key]}"].font = BLUE
for col in qcols + ["K", "L", "M"]:
    g = lambda k: f"{col}{QR[k]}"
    Q[g("gp_oth")] = f"={g('rev_oth')}"
    Q[g("gp_tot")] = f"={g('gp_ret')}+{g('gp_whs')}+{g('gp_oth')}"
    Q[g("chk_rev")] = f"={g('rev_ret')}+{g('rev_whs')}+{g('rev_oth')}-{g('rev_tot')}"
    Q[g("oth_opx")] = f"={g('gp_tot')}-{g('sga')}-{g('oi')}"
    Q[g("chk_ebitda")] = f"={g('oi')}+{g('oth_opx')}+{g('da')}+{g('sbc')}-{g('warrant')}-{g('adj_ebitda')}"
    Q[g("gos_m")] = f"=IF({g('orig')}>0,{g('gos')}/{g('orig')},\"\")"
    Q[g("gpu")] = f"={g('gp_tot')}*1000000/{g('units')}"
    Q[g("gpu_ret")] = f"={g('gp_ret')}*1000000/{g('units')}"
    Q[g("gpu_oth")] = f"={g('gp_oth')}*1000000/{g('units')}"
    Q[g("sga_pu")] = f"={g('sga')}*1000000/{g('units')}"
    Q[g("ebitda_m")] = f"={g('adj_ebitda')}/{g('rev_tot')}"
    Q[g("gos_sh")] = f"={g('gos')}/{g('adj_ebitda')}"
    Q[g("gosrp_sh")] = f"=IF({g('rp_oth')}=\"\",\"\",({g('gos')}+{g('rp_oth')})/{g('adj_ebitda')})"
    Q[g("ebitda_xg")] = f"={g('adj_ebitda')}-{g('gos')}"
    Q[g("ebitda_xg_m")] = f"={g('ebitda_xg')}/{g('rev_tot')}"
# FY2024 關係人其他營收 200、FY2025 347（股東信）
Q[f"L{QR['rp_oth']}"] = 200; Q[f"M{QR['rp_oth']}"] = 347
for k in ("cash", "inv", "frhfs", "bi", "debt_bs", "equity", "liq_total"):
    Q[f"L{QR[k]}"] = qd[k][0]; Q[f"M{QR[k]}"] = qd[k][4]
for c in ("L", "M"):
    Q[f"{c}{QR['whs_units']}"] = (199780 if c == "L" else 297643)
# 近四季（TTM，Q3'25–Q2'26）
TT = rr + 1
Q.cell(TT, 1, "近四季（TTM，Q4 2025–Q3 2026）").font = BOLD
ttm_items = [("ttm_ebitda", "調整後 EBITDA（TTM）", "adj_ebitda"), ("ttm_gos", "貸款出售利得（TTM）", "gos"), ("ttm_int", "利息費用（TTM）", "int"),
             ("ttm_capex", "資本支出（TTM）", "capex"), ("ttm_cfo", "營業現金流（TTM）", "cfo"), ("ttm_rp", "關係人其他營收（TTM）", "rp_oth")]
for i, (k, lab, src) in enumerate(ttm_items):
    rw = TT + 1 + i; QR[k] = rw
    Q.cell(rw, 1, lab); Q.cell(rw, 2, "US$M"); Q.cell(rw, 3, f"=SUM(G{QR[src]}:J{QR[src]})")
    dn(k, f"Q!$C${rw}")
rw = TT + 1 + len(ttm_items)
QR["ttm_lev"] = rw; Q.cell(rw, 1, "債務／調整後 EBITDA（TTM，2026-09-30）"); Q.cell(rw, 2, "x"); Q.cell(rw, 3, f"=J{QR['debt_bs']}/C{QR['ttm_ebitda']}")
QR["ttm_nlev"] = rw + 1; Q.cell(rw + 1, 1, "淨債務／調整後 EBITDA（TTM）"); Q.cell(rw + 1, 2, "x"); Q.cell(rw + 1, 3, f"=(J{QR['debt_bs']}-J{QR['cash']})/C{QR['ttm_ebitda']}")
QR["ttm_lev_xg"] = rw + 2; Q.cell(rw + 2, 1, "債務／扣除出售利得的 EBITDA（TTM）"); Q.cell(rw + 2, 2, "x"); Q.cell(rw + 2, 3, f"=J{QR['debt_bs']}/(C{QR['ttm_ebitda']}-C{QR['ttm_gos']})")
QR["ttm_cov"] = rw + 3; Q.cell(rw + 3, 1, "（EBITDA − 資本支出）／利息（TTM）"); Q.cell(rw + 3, 2, "x"); Q.cell(rw + 3, 3, f"=(C{QR['ttm_ebitda']}-C{QR['ttm_capex']})/C{QR['ttm_int']}")
QR["ttm_gosrp"] = rw + 4; Q.cell(rw + 4, 1, "（出售利得＋關係人其他營收）÷ 調整後 EBITDA（TTM）"); Q.cell(rw + 4, 2, "%"); Q.cell(rw + 4, 3, f"=(C{QR['ttm_gos']}+C{QR['ttm_rp']})/C{QR['ttm_ebitda']}")
for k in ("ttm_lev", "ttm_nlev", "ttm_lev_xg", "ttm_cov", "ttm_gosrp"):
    dn(k, f"Q!$C${QR[k]}")
Q.column_dimensions["A"].width = 52
ROWMAP["Q"] = QR
# 名稱：2026 上半年與 FY2025 事實
for k in ("units", "rev_ret", "rev_whs", "rev_oth", "gp_ret", "gp_whs", "sga", "oi", "da", "sbc", "orig", "gos", "warrant", "int", "tax", "ni", "capex", "cfo", "adj_ebitda", "oth_exp"):
    dn("ytd_" + k, f"Q!$K${QR[k]}")
    dn("fy25_" + k, f"Q!$M${QR[k]}")
for k in ("cash", "inv", "frhfs", "bi", "debt_bs"):
    dn("lq_" + k, f"Q!$J${QR[k]}")
    dn("d25_" + k, f"Q!$G${QR[k]}")
dn("lq_orig", f"Q!$J${QR['orig']}"); dn("lq_gos", f"Q!$J${QR['gos']}"); dn("lq_ebitda", f"Q!$J${QR['adj_ebitda']}"); dn("lq_int", f"Q!$J${QR['int']}"); dn("lq_capex", f"Q!$J${QR['capex']}")
dn("lq_units", f"Q!$J${QR['units']}")

# =====================================================================================
# Thesis — 命題：零售台數 × 每台單位經濟（毛利三層：零售／批發／其他＝貸款出售利得＋非貸款）
# =====================================================================================
TH = wb.create_sheet("Thesis")
TH["A1"] = "Thesis — 命題模組：零售台數 × 每台單位經濟。其他毛利拆成「貸款出售利得＝每台發放額 × 利得率」與「非貸款部分（VSC、GAP 等）」；2026 年＝上半年實績＋下半年模型"
TH["A2"] = "讀法：Carvana 的利潤約三到四成來自把車貸賣掉的利得；賣得掉、賣得好，取決於消費信用表現與 ABS 市場。情境 C 的利得率由 Credit 頁 S2 的損失與利差假設推導"
R = {s: {} for s in SC}


def put_row(ws, R_s, row, key, label, unit, cells, fmt=None):
    ws.cell(row, 1, label); ws.cell(row, 2, unit)
    for col, f in cells.items():
        ws[f"{col}{row}"] = f
        if fmt:
            ws[f"{col}{row}"].number_format = fmt
    R_s[key] = row


def interp3(y, v27, vLT):
    """2027＝v27；2028、2029 線性；2030 起＝vLT（字串為名稱）"""
    if y == 2027: return f"={v27}"
    if y == 2028: return f"={v27}+({vLT}-{v27})/3"
    if y == 2029: return f"={v27}+({vLT}-{v27})*2/3"
    return f"={vLT}"


row = 4
TBLK = {}
for s in SC:
    Rs = R[s]; r0 = row
    TH.cell(row, 1, SCN[s]).font = BOLD; hdr(TH, row + 1, ["項目", "單位"] + [("FY%dA" % y if y == 2025 else "FY%dE" % y) for y in YRS]); Rs["hdr"] = row + 1
    row += 2
    # 台數年增
    cells = {}
    for y in YRS:
        c = yc(y)
        if y == 2025: cells[c] = "=L4/K4-1"
        elif y == 2026: cells[c] = f"={c}{row+1}/{yc(2025)}{row+1}-1"
        elif y == 2027: cells[c] = f"=g27_{s}"
        elif y == 2028: cells[c] = f"=g28_{s}"
        elif y == 2029: cells[c] = f"=(g28_{s}+g30_{s})/2"
        elif y == 2030: cells[c] = f"=g30_{s}"
        else: cells[c] = f"=g30_{s}+(g35_{s}-g30_{s})*{y-2030}/5"
    cells[yc(2025)] = f"={q('units','M')}/{q('units','L')}-1"
    put_row(TH, Rs, row, "g", "零售台數年增（2027 起為輸入；觸及上限後實際年增較低）", "%", cells, "0.0%"); row += 1
    cells = {yc(2025): "=fy25_units", yc(2026): "=ytd_units+u_q4"}
    for y in FC:
        cells[yc(y)] = f"=MIN({yc(y-1)}{row}*(1+{yc(y)}{Rs['g']}),unit_cap)"
    put_row(TH, Rs, row, "u", "零售台數", "輛", cells, "#,##0"); row += 1
    U = lambda y: f"{yc(y)}{Rs['u']}"
    H2U = f"({U(2026)}-ytd_units)"
    # 售價、零售營收
    cells = {yc(2025): f"=fy25_rev_ret*1000000/{U(2025)}", yc(2026): f"=(ytd_rev_ret*1000000+{H2U}*asp_q4)/{U(2026)}"}
    for y in FC: cells[yc(y)] = f"={yc(y-1)}{row}*(1+asp_g)"
    put_row(TH, Rs, row, "asp", "零售每台營收", "US$", cells, "#,##0"); row += 1
    put_row(TH, Rs, row, "rev_ret", "零售車輛銷售", "US$M", {yc(y): f"={U(y)}*{yc(y)}{Rs['asp']}/1000000" for y in YRS}, "#,##0"); row += 1
    cells = {yc(2025): f"=fy25_rev_whs*1000000/{U(2025)}", yc(2026): f"=(ytd_rev_whs*1000000+{H2U}*whs_pu)/{U(2026)}"}
    for y in FC: cells[yc(y)] = f"={yc(y-1)}{row}*(1+asp_g)"
    put_row(TH, Rs, row, "whs_pu", "批發營收／零售台", "US$", cells, "#,##0"); row += 1
    put_row(TH, Rs, row, "rev_whs", "批發銷售與營收", "US$M", {yc(y): f"={U(y)}*{yc(y)}{Rs['whs_pu']}/1000000" for y in YRS}, "#,##0"); row += 1
    # 零售 GPU
    cells = {yc(2025): f"=fy25_gp_ret*1000000/{U(2025)}", yc(2026): f"=(ytd_gp_ret*1000000+{H2U}*(gpr_q4+q4_adj))/{U(2026)}"}
    for y in FC: cells[yc(y)] = interp3(y, f"gpr27_{s}", f"gprLT_{s}")
    put_row(TH, Rs, row, "gpr", "零售毛利／台", "US$", cells, "#,##0"); row += 1
    cells = {yc(2025): f"=fy25_gp_whs*1000000/{U(2025)}", yc(2026): f"=(ytd_gp_whs*1000000+{H2U}*gpw_h2)/{U(2026)}"}
    for y in FC: cells[yc(y)] = interp3(y, "gpw27_C", "gpw_h2") if s == "C" else "=gpw_h2"
    put_row(TH, Rs, row, "gpw", "批發毛利／零售台", "US$", cells, "#,##0"); row += 1
    # 貸款
    cells = {yc(2025): f"=fy25_orig*1000000/{U(2025)}", yc(2026): f"=(ytd_orig*1000000+{H2U}*opu)/{U(2026)}"}
    for y in FC: cells[yc(y)] = f"={yc(y-1)}{row}*(1+asp_g)"
    put_row(TH, Rs, row, "opu", "每零售台貸款發放額", "US$", cells, "#,##0"); row += 1
    put_row(TH, Rs, row, "orig", "貸款發放", "US$M", {yc(y): f"={U(y)}*{yc(y)}{Rs['opu']}/1000000" for y in YRS}, "#,##0"); row += 1
    cells = {yc(2025): "=fy25_gos/fy25_orig", yc(2026): f"=(ytd_gos+({yc(2026)}{Rs['orig']}-ytd_orig)*gos_h2)/{yc(2026)}{Rs['orig']}"}
    for y in FC: cells[yc(y)] = interp3(y, f"gos27_{s}", f"gosLT_{s}")
    put_row(TH, Rs, row, "gos_m", "貸款出售利得率", "%", cells, "0.00%"); row += 1
    put_row(TH, Rs, row, "gos", "貸款出售利得（其他毛利之一）", "US$M", {yc(y): f"={yc(y)}{Rs['orig']}*{yc(y)}{Rs['gos_m']}" for y in YRS}, "#,##0"); row += 1
    cells = {yc(2025): f"=({q('rev_oth','M')}-fy25_gos)*1000000/{U(2025)}", yc(2026): f"=(({q('rev_oth','K')}-ytd_gos)*1000000+{H2U}*oth_pu)/{U(2026)}"}
    for y in FC: cells[yc(y)] = f"={yc(y-1)}{row}*(1+oth_g)"
    put_row(TH, Rs, row, "othpu", "非貸款其他毛利／台", "US$", cells, "#,##0"); row += 1
    put_row(TH, Rs, row, "oth_nl", "非貸款其他毛利（VSC、GAP 等）", "US$M", {yc(y): f"={U(y)}*{yc(y)}{Rs['othpu']}/1000000" for y in YRS}, "#,##0"); row += 1
    put_row(TH, Rs, row, "gp_ret", "零售毛利", "US$M", {yc(y): f"={U(y)}*{yc(y)}{Rs['gpr']}/1000000" for y in YRS}, "#,##0"); row += 1
    put_row(TH, Rs, row, "gp_whs", "批發毛利", "US$M", {yc(y): f"={U(y)}*{yc(y)}{Rs['gpw']}/1000000" for y in YRS}, "#,##0"); row += 1
    put_row(TH, Rs, row, "gp_oth", "其他毛利（＝其他營收）", "US$M", {yc(y): f"={yc(y)}{Rs['gos']}+{yc(y)}{Rs['oth_nl']}" for y in YRS}, "#,##0"); row += 1
    put_row(TH, Rs, row, "gpu", "總毛利／台", "US$", {yc(y): f"=({yc(y)}{Rs['gp_ret']}+{yc(y)}{Rs['gp_whs']}+{yc(y)}{Rs['gp_oth']})*1000000/{U(y)}" for y in YRS}, "#,##0"); row += 1
    # SG&A
    cells = {yc(2025): f"={q('sga','M')}", yc(2026): f"=ytd_sga+sga_fix+{H2U}*sga_var/1000000"}
    for y in FC: cells[yc(y)] = f"=4*sga_fix*(1+infl)^{y-2026}+{U(y)}*sga_var*(1-eff_{s})^{y-2026}/1000000"
    put_row(TH, Rs, row, "sga", "SG&A", "US$M", cells, "#,##0"); row += 1
    put_row(TH, Rs, row, "sga_pu", "SG&A／台", "US$", {yc(y): f"={yc(y)}{Rs['sga']}*1000000/{U(y)}" for y in YRS}, "#,##0"); row += 1
    put_row(TH, Rs, row, "sga_var_pu", "變動 SG&A／台", "US$", {**{yc(2026): "=sga_var"}, **{yc(y): f"=sga_var*(1-eff_{s})^{y-2026}" for y in FC}}, "#,##0"); row += 1
    # 產能與資本支出
    cells = {yc(2026): f"=MAX(cap_0,{yc(2027)}{Rs['u']}/util)"}
    for y in FC:
        nxt = f"{yc(y+1)}{Rs['u']}" if y < 2035 else f"{yc(y)}{Rs['u']}*(1+{yc(y)}{Rs['g']})"
        cells[yc(y)] = f"=MAX({yc(y-1)}{row},MIN({nxt},unit_cap)/util)"
    put_row(TH, Rs, row, "cap", "年產能（期末，足以支應次年台數）", "輛", cells, "#,##0"); row += 1
    put_row(TH, Rs, row, "gcx", "成長性資本支出（新增產能 × 每台成本）", "US$M", {**{yc(2026): "=0"}, **{yc(y): f"=({yc(y)}{Rs['cap']}-{yc(y-1)}{Rs['cap']})*capcost/1000000" for y in FC}}, "#,##0"); row += 1
    TBLK[s] = (r0, row)
    row += 2

# =====================================================================================
# PL — 營收線 → 合併 → 損益（官方：單一應報導部門；營收線三條為公司固定揭露口徑）
# =====================================================================================
PL = wb.create_sheet("PL")
PL["A1"] = "PL — 損益表。Carvana 只有一個應報導部門（10-Q 附註 2：CODM 看合併淨利）；營收列採公司固定揭露的三條營收線，「其他銷售與營收」下拆命題子列（貸款出售利得／非貸款）"
PL["A2"] = "FY2025A 為實績（含第四季釋放評價備抵的稅務利益 2,785 與 TRA 重估等其他費用 2,250）；FY2026E＝上半年實績＋下半年模型"
hdr(PL, 3, ["項目", "單位"] + [("FY%dA" % y if y == 2025 else "FY%dE" % y) for y in YRS])
P = {s: {} for s in SC}
row = 5
PL.cell(4, 1, "核對列（三情境合計絕對差，應為 0）").font = BOLD
CHK = {"seg": 5, "sub": 6}
PL.cell(5, 1, "核對：營收線合計 − 合併營收"); PL.cell(6, 1, "核對：其他營收子列合計 − 其他營收")
row = 8
for s in SC:
    Ps = P[s]; T = R[s]
    PL.cell(row, 1, SCN[s]).font = BOLD; row += 1
    t = lambda k, y: f"Thesis!{yc(y)}{T[k]}"
    put_row(PL, Ps, row, "rev_ret", "零售車輛銷售", "US$M", {yc(y): f"={t('rev_ret', y)}" for y in YRS}, "#,##0"); row += 1
    put_row(PL, Ps, row, "rev_whs", "批發銷售與營收", "US$M", {yc(y): f"={t('rev_whs', y)}" for y in YRS}, "#,##0"); row += 1
    put_row(PL, Ps, row, "rev_oth", "其他銷售與營收", "US$M", {**{yc(2025): f"={q('rev_oth','M')}"}, **{yc(y): f"={t('gp_oth', y)}" for y in YRS if y >= 2026}}, "#,##0"); row += 1
    put_row(PL, Ps, row, "sub_gos", "　其中：貸款出售利得（命題子列）", "US$M", {yc(y): f"={t('gos', y)}" for y in YRS}, "#,##0"); row += 1
    put_row(PL, Ps, row, "sub_nl", "　其中：VSC、GAP 等非貸款（命題子列）", "US$M", {yc(y): f"={yc(y)}{Ps['rev_oth']}-{yc(y)}{Ps['sub_gos']}" for y in YRS}, "#,##0"); row += 1
    put_row(PL, Ps, row, "rev", "合併營收", "US$M", {**{yc(2025): f"={q('rev_tot','M')}"}, **{yc(y): f"={yc(y)}{Ps['rev_ret']}+{yc(y)}{Ps['rev_whs']}+{yc(y)}{Ps['rev_oth']}" for y in YRS if y >= 2026}}, "#,##0"); row += 1
    put_row(PL, Ps, row, "rev_g", "營收年增", "%", {yc(y): f"={yc(y)}{Ps['rev']}/{yc(y-1)}{Ps['rev']}-1" for y in YRS if y >= 2026}, "0.0%"); row += 1
    put_row(PL, Ps, row, "gp_ret", "零售毛利", "US$M", {yc(y): f"={t('gp_ret', y)}" for y in YRS}, "#,##0"); row += 1
    put_row(PL, Ps, row, "gp_whs", "批發毛利", "US$M", {yc(y): f"={t('gp_whs', y)}" for y in YRS}, "#,##0"); row += 1
    put_row(PL, Ps, row, "gp_oth", "其他毛利", "US$M", {yc(y): f"={yc(y)}{Ps['rev_oth']}" for y in YRS}, "#,##0"); row += 1
    put_row(PL, Ps, row, "gp", "毛利合計", "US$M", {yc(y): f"={yc(y)}{Ps['gp_ret']}+{yc(y)}{Ps['gp_whs']}+{yc(y)}{Ps['gp_oth']}" for y in YRS}, "#,##0"); row += 1
    put_row(PL, Ps, row, "sga", "SG&A", "US$M", {yc(y): f"={t('sga', y)}" for y in YRS}, "#,##0"); row += 1
    put_row(PL, Ps, row, "oi", "營業利益", "US$M", {**{yc(2025): "=fy25_oi"}, **{yc(y): f"={yc(y)}{Ps['gp']}-{yc(y)}{Ps['sga']}" for y in YRS if y >= 2026}}, "#,##0"); row += 1
    put_row(PL, Ps, row, "oi_m", "營業利益率", "%", {yc(y): f"={yc(y)}{Ps['oi']}/{yc(y)}{Ps['rev']}" for y in YRS}, "0.0%"); row += 1
    Ps["int"] = row; PL.cell(row, 1, "利息費用（淨；CF 頁債務表）"); PL.cell(row, 2, "US$M"); row += 1
    put_row(PL, Ps, row, "oth", "其他費用（收益）", "US$M", {**{yc(2025): f"={q('oth_exp','M')}"}, **{yc(2026): "=ytd_oth_exp"}, **{yc(y): "=0" for y in FC}}, "#,##0"); row += 1
    put_row(PL, Ps, row, "pti", "稅前淨利", "US$M", {yc(y): f"={yc(y)}{Ps['oi']}-{yc(y)}{Ps['int']}-{yc(y)}{Ps['oth']}" for y in YRS}, "#,##0"); row += 1
    cells = {yc(2025): "=fy25_tax", yc(2026): f"=ytd_tax+({yc(2026)}{row-1}-(ytd_oi-ytd_int-ytd_oth_exp))*etr_26"}
    for y in FC:
        er = {2027: "etr_27", 2028: "(etr_27+(etr_30-etr_27)/3)", 2029: "(etr_27+(etr_30-etr_27)*2/3)"}.get(y, "etr_30")
        cells[yc(y)] = f"={yc(y)}{row-1}*{er}"
    put_row(PL, Ps, row, "tax", "所得稅", "US$M", cells, "#,##0"); row += 1
    put_row(PL, Ps, row, "ni", "淨利（合併，含非控制權益）", "US$M", {yc(y): f"={yc(y)}{Ps['pti']}-{yc(y)}{Ps['tax']}" for y in YRS}, "#,##0"); row += 1
    put_row(PL, Ps, row, "ni_m", "淨利率", "%", {yc(y): f"={yc(y)}{Ps['ni']}/{yc(y)}{Ps['rev']}" for y in YRS}, "0.0%"); row += 1
    put_row(PL, Ps, row, "eps", "每股盈餘（完全轉換股數）", "US$", {yc(y): f"={yc(y)}{Ps['ni']}/sh_fx" for y in YRS if y >= 2026}, "0.00"); row += 1
    put_row(PL, Ps, row, "da", "折舊攤銷", "US$M", {**{yc(2025): "=fy25_da", yc(2026): "=ytd_da+da_q4"}, **{yc(y): f"=(ytd_da+da_q4)*(1+infl)^{y-2026}+SUM(Thesis!{yc(2027)}{T['gcx']}:{yc(y-1)}{T['gcx']})/life" if y > 2027 else f"=(ytd_da+da_q4)*(1+infl)" for y in FC}}, "#,##0"); row += 1
    put_row(PL, Ps, row, "sbc", "股份基礎給付", "US$M", {**{yc(2025): "=fy25_sbc", yc(2026): "=ytd_sbc+sbc_q4"}, **{yc(y): f"=(ytd_sbc+sbc_q4)*(1+infl)^{y-2026}" for y in FC}}, "#,##0"); row += 1
    put_row(PL, Ps, row, "warr", "權證營收（扣除）", "US$M", {**{yc(2025): "=fy25_warrant", yc(2026): "=ytd_warrant+warr/4"}, **{yc(y): "=warr" for y in FC}}, "#,##0"); row += 1
    put_row(PL, Ps, row, "ebitda", "調整後 EBITDA", "US$M", {**{yc(2025): f"={q('adj_ebitda','M')}"}, **{yc(y): f"={yc(y)}{Ps['oi']}+{yc(y)}{Ps['da']}+{yc(y)}{Ps['sbc']}-{yc(y)}{Ps['warr']}" for y in YRS if y >= 2026}}, "#,##0"); row += 1
    put_row(PL, Ps, row, "ebitda_m", "調整後 EBITDA 率", "%", {yc(y): f"={yc(y)}{Ps['ebitda']}/{yc(y)}{Ps['rev']}" for y in YRS}, "0.0%"); row += 1
    put_row(PL, Ps, row, "gos_sh", "貸款出售利得 ÷ 調整後 EBITDA", "%", {yc(y): f"={yc(y)}{Ps['sub_gos']}/{yc(y)}{Ps['ebitda']}" for y in YRS}, "0.0%"); row += 1
    row += 1
PL["A4"].font = BOLD
for col in [yc(y) for y in YRS]:
    PL[f"{col}5"] = "=" + "+".join(f"ABS({col}{P[s]['rev_ret']}+{col}{P[s]['rev_whs']}+{col}{P[s]['rev_oth']}-{col}{P[s]['rev']})" for s in SC)
    PL[f"{col}6"] = "=" + "+".join(f"ABS({col}{P[s]['sub_gos']}+{col}{P[s]['sub_nl']}-{col}{P[s]['rev_oth']})" for s in SC)

# =====================================================================================
# CF — 收支模組：現金流、債務表、利息、收支關卡
# =====================================================================================
CF = wb.create_sheet("CF")
CF["A1"] = "CF — 收支模組。收支關卡＝自由現金流 − TRA 付款 − 代 NCI 繳稅 − 排定到期本金 − 贖回溢價；期末現金低於最低營運現金時以新增外部融資補足（利息自次年計）。FY2026＝全年（上半年實績＋下半年模型）"
CF["A2"] = "利息以期初餘額 × 各工具利率計（避免循環參照）；「其他利息與費用」以 Q2 2026 實績校準，並隨貸款發放量縮放"
hdr(CF, 3, ["收支摘要", "單位", "情境 A", "情境 B", "情境 C"])
CFS = {"gate5": 4, "need5": 5, "cash_min": 6}
CF.cell(4, 1, "收支關卡：FY2027–FY2031 合計（A／B／C 於 C、D、E 欄）"); CF.cell(5, 1, "新增外部融資需求：FY2027–FY2031 合計"); CF.cell(6, 1, "FY2027–FY2031 期末現金最低值")
CF.cell(7, 1, "FY2027–FY2031 超額現金（高於現金上限，可回饋股東或提前還債）合計")
CF.cell(9, 1, "其他利息與費用校準（年化，以 2026-09-30 結構；扣除 Q3 中 2030 年票據 1.5 個月的高利率）"); CF.cell(9, 2, "US$M")
CF["C9"] = "=4*(int_lq-tlb*(ssn_r-swap_fix-tlb_sp)*1.5/12)-tlb*(swap_fix+tlb_sp)-ssn_31*ssn_r-uns*uns_r-re_fin*re_r-bi_adv*lq_bi*bi_r-fleet*fleet_r-finls*finls_r-fp_ratio*lq_inv*fp_r+lq_cash*cash_y"
dn("tlb_old", "In!$E$" + str(ROWMAP["In"]["tlb"]))  # 2030 年票據金額＝定期貸款 B 金額（1,660）
dn("int_cal", "CF!$C$9")
C = {s: {} for s in SC}
row = 12
for s in SC:
    Cs = C[s]; Ps = P[s]; T = R[s]
    CF.cell(row, 1, SCN[s]).font = BOLD; row += 1
    hdr(CF, row, ["項目", "單位"] + [("FY%dA" % y if y == 2025 else "FY%dE" % y) for y in YRS]); Cs["hdr"] = row; row += 1
    p = lambda k, y: f"PL!{yc(y)}{Ps[k]}"
    t = lambda k, y: f"Thesis!{yc(y)}{T[k]}"
    # 營運資金餘額
    cells = {yc(2025): "=d25_inv", yc(2026): f"=inv_d*({p('rev_ret', 2026)}-{p('gp_ret', 2026)})/365"}
    for y in FC:
        d = "inv_d_C" if (s == "C" and y == 2027) else "inv_d"
        cells[yc(y)] = f"={d}*({p('rev_ret', y)}-{p('gp_ret', y)})/365"
    put_row(CF, Cs, row, "inv", "存貨（期末）", "US$M", cells, "#,##0"); row += 1
    put_row(CF, Cs, row, "fr", "待售貸款（期末）", "US$M", {**{yc(2025): "=d25_frhfs"}, **{yc(y): f"=fr_d*{t('orig', y)}/365" for y in YRS if y >= 2026}}, "#,##0"); row += 1
    put_row(CF, Cs, row, "ap", "應付與應計（期末）", "US$M", {**{yc(y): f"=ap_pct*({p('rev', y)}-{p('gp', y)}+{p('sga', y)})" for y in YRS if y >= 2026}}, "#,##0"); row += 1
    put_row(CF, Cs, row, "bi", "證券化受益權（期末，隨發放量縮放）", "US$M", {**{yc(2025): "=d25_bi", yc(2026): f"=lq_bi*{t('orig', 2026)}/{t('orig', 2026)}"}, **{yc(y): f"=lq_bi*{t('orig', y)}/{t('orig', 2026)}" for y in FC}}, "#,##0"); row += 1
    # 現金流
    put_row(CF, Cs, row, "ni", "淨利", "US$M", {yc(y): f"={p('ni', y)}" for y in YRS if y >= 2026}, "#,##0"); row += 1
    put_row(CF, Cs, row, "da", "＋折舊攤銷＋SBC", "US$M", {yc(y): f"={p('da', y)}+{p('sbc', y)}" for y in YRS if y >= 2026}, "#,##0"); row += 1
    put_row(CF, Cs, row, "dtax", "＋帳面稅與現金稅差額（遞延所得稅資產抵用）", "US$M", {yc(y): (f"={p('tax', y)}-MAX(0,{p('pti', y)})*ctr" if y <= 2028 else "=0") for y in YRS if y >= 2026}, "#,##0"); row += 1
    put_row(CF, Cs, row, "dwc", "−營運資金增加（存貨＋待售貸款 − 應付）", "US$M", {yc(y): f"=-(({yc(y)}{Cs['inv']}-{yc(y-1)}{Cs['inv']})+({yc(y)}{Cs['fr']}-{yc(y-1)}{Cs['fr']})-({yc(y)}{Cs['ap']}-{yc(y-1)}{Cs['ap']}))" for y in YRS if y >= 2026}, "#,##0"); row += 1
    put_row(CF, Cs, row, "rr", "−風險自留淨現金（受益權增加 ×（1 − 融資成數））", "US$M", {yc(y): f"=-({yc(y)}{Cs['bi']}-{yc(y-1)}{Cs['bi']})*(1-bi_adv)" for y in YRS if y >= 2026}, "#,##0"); row += 1
    put_row(CF, Cs, row, "cfo", "營業現金流（模型口徑）", "US$M", {yc(y): f"={yc(y)}{Cs['ni']}+{yc(y)}{Cs['da']}+{yc(y)}{Cs['dtax']}+{yc(y)}{Cs['dwc']}+{yc(y)}{Cs['rr']}" for y in YRS if y >= 2026}, "#,##0"); row += 1
    put_row(CF, Cs, row, "capex", "資本支出（維護＋成長）", "US$M", {**{yc(2026): "=ytd_capex+capex_q4"}, **{yc(y): f"={p('rev', y)}*mcx+{t('gcx', y)}" for y in FC}}, "#,##0"); row += 1
    put_row(CF, Cs, row, "fcf", "自由現金流", "US$M", {yc(y): f"={yc(y)}{Cs['cfo']}-{yc(y)}{Cs['capex']}" for y in YRS if y >= 2026}, "#,##0"); row += 1
    put_row(CF, Cs, row, "tra", "TRA 付款", "US$M", {**{yc(2026): "=37+tra_q4"}, **{yc(y): f"=MIN(tra_0-tra_q4-SUM({yc(2027)}{row}:{yc(y-1)}{row}),tra_0/tra_yrs)" if y > 2027 else "=MIN(tra_0-tra_q4,tra_0/tra_yrs)" for y in FC}}, "#,##0"); row += 1
    put_row(CF, Cs, row, "ncit", "代非控制權益繳稅", "US$M", {yc(y): "=ncitax" for y in YRS if y >= 2026}, "#,##0"); row += 1
    # 排定本金：定期貸款 1%／年（2027 起）、無擔保票據到期
    sched = {2026: "0", 2027: "tlb*0.01+32", 2028: "tlb*0.01+22", 2029: "tlb*0.01+26", 2030: "tlb*0.01+27", 2031: "tlb*0.01", 2032: "tlb*0.01", 2033: f"tlb-tlb*0.01*6"}
    put_row(CF, Cs, row, "prin", "排定到期本金（定期貸款攤還、無擔保票據）", "US$M", {yc(y): "=" + sched.get(y, "0") for y in YRS if y >= 2026}, "#,##0"); row += 1
    if s in ("A", "B"):
        cells = {yc(y): f"=IF({y}=refi_yr,ssn_31*refi_prem,0)" for y in YRS if y >= 2027}
    else:
        cells = {yc(y): "=0" for y in YRS if y >= 2027}
    cells[yc(2026)] = "=redeem_cost"
    put_row(CF, Cs, row, "prem", "贖回溢價與費用", "US$M", cells, "#,##0"); row += 1
    put_row(CF, Cs, row, "gate", "收支關卡：自由現金流扣 TRA 與到期本金", "US$M", {yc(y): f"={yc(y)}{Cs['fcf']}-{yc(y)}{Cs['tra']}-{yc(y)}{Cs['ncit']}-{yc(y)}{Cs['prin']}-{yc(y)}{Cs['prem']}" for y in YRS if y >= 2026}, "#,##0"); row += 1
    # 現金與新增融資
    Cs["cash_pre"] = row; Cs["need"] = row + 1; Cs["dist"] = row + 2; Cs["cash"] = row + 3
    CF.cell(row, 1, "期末現金（融資與分配前）"); CF.cell(row, 2, "US$M")
    CF.cell(row + 1, 1, "新增外部融資（補足至最低營運現金）"); CF.cell(row + 1, 2, "US$M")
    CF.cell(row + 2, 1, "超額現金（高於現金上限；視為回饋股東或提前還債，不再計息）"); CF.cell(row + 2, 2, "US$M")
    CF.cell(row + 3, 1, "期末現金"); CF.cell(row + 3, 2, "US$M")
    CF[f"{yc(2025)}{row+3}"] = "=d25_cash"
    for y in YRS:
        if y < 2026: continue
        c = yc(y); pc_ = yc(y - 1)
        CF[f"{c}{row}"] = f"={pc_}{row+3}+{c}{Cs['gate']}"
        CF[f"{c}{row+1}"] = f"=MAX(0,min_cash-{c}{row})"
        CF[f"{c}{row+2}"] = f"=MAX(0,{c}{row}-cash_cap)"
        CF[f"{c}{row+3}"] = f"={c}{row}+{c}{row+1}-{c}{row+2}"
        for rr_ in (row, row + 1, row + 2, row + 3): CF[f"{c}{rr_}"].number_format = "#,##0"
    row += 4
    # 債務表
    CF.cell(row, 1, "債務（期末）").font = BOLD; row += 1
    tlbcells = {yc(2025): "=0", yc(2026): "=tlb"}
    for y in FC:
        tlbcells[yc(y)] = "=0" if y >= 2033 else f"=tlb-tlb*0.01*{y-2026}"
    put_row(CF, Cs, row, "d_tlb", "定期貸款 B（2026 年起；取代 2030 年擔保票據）", "US$M", tlbcells, "#,##0"); row += 1
    put_row(CF, Cs, row, "d_ssn30", "2030 年擔保票據", "US$M", {**{yc(2025): "=tlb_old"}, **{yc(y): "=0" for y in YRS if y >= 2026}}, "#,##0"); row += 1
    if s in ("A", "B"):
        cells = {yc(y): f"=IF({y}<refi_yr,ssn_31,0)" for y in YRS}
        cells2 = {yc(y): f"=IF({y}>=refi_yr,ssn_31*(1+0),0)" for y in YRS}
    else:
        cells = {yc(y): f"=IF({y}<2031,ssn_31,0)" for y in YRS}
        cells2 = {yc(y): f"=IF({y}>=2031,ssn_31,0)" for y in YRS}
    put_row(CF, Cs, row, "d_ssn31", "2031 年擔保票據", "US$M", cells, "#,##0"); row += 1
    put_row(CF, Cs, row, "d_refi", "2031 年票據再融資貸款", "US$M", cells2, "#,##0"); row += 1
    uns = {2025: "uns", 2026: "uns", 2027: "uns-32", 2028: "uns-32-22", 2029: "uns-32-22-26"}
    put_row(CF, Cs, row, "d_uns", "無擔保票據", "US$M", {yc(y): "=" + uns.get(y, "0") for y in YRS}, "#,##0"); row += 1
    put_row(CF, Cs, row, "d_re", "不動產售後租回融資", "US$M", {yc(y): "=re_fin" for y in YRS}, "#,##0"); row += 1
    put_row(CF, Cs, row, "d_bi", "受益權融資", "US$M", {yc(y): f"=bi_adv*{yc(y)}{Cs['bi']}" for y in YRS}, "#,##0"); row += 1
    put_row(CF, Cs, row, "d_oth", "車隊融資＋融資租賃", "US$M", {yc(y): "=fleet+finls" for y in YRS}, "#,##0"); row += 1
    put_row(CF, Cs, row, "d_fp", "存貨融資（floor plan）", "US$M", {yc(y): f"=fp_ratio*{yc(y)}{Cs['inv']}" for y in YRS}, "#,##0"); row += 1
    put_row(CF, Cs, row, "d_new", "新增外部融資（累計）", "US$M", {**{yc(2025): "=0"}, **{yc(y): f"={yc(y-1)}{row}+{yc(y)}{Cs['need']}" for y in YRS if y >= 2026}}, "#,##0"); row += 1
    put_row(CF, Cs, row, "debt", "債務合計", "US$M", {yc(y): f"=SUM({yc(y)}{Cs['d_tlb']}:{yc(y)}{Cs['d_new']})" for y in YRS}, "#,##0"); row += 1
    # 利息
    refi_rate = "(sofr+refi_sp)" if s in ("A", "B") else "(sofr+stress_sp)"
    refi_year = "refi_yr" if s in ("A", "B") else "2031"
    m9 = 7.5 if s in ("A", "B") else 5
    cells = {yc(2026): "=ytd_int+int_lq-tlb*(ssn_r-swap_fix-tlb_sp)*1.5/12"}
    for y in FC:
        c = yc(y); o = yc(y - 1)
        cells[c] = (f"={o}{Cs['d_tlb']}*(IF({y}<=2028,swap_fix,sofr)+tlb_sp)"
                    f"+IF({y}={refi_year},ssn_31*(ssn_r*{m9}/12+{refi_rate}*(12-{m9})/12),{o}{Cs['d_ssn31']}*ssn_r+{o}{Cs['d_refi']}*{refi_rate})"
                    f"+{o}{Cs['d_uns']}*uns_r+{o}{Cs['d_re']}*re_r+{o}{Cs['d_bi']}*bi_r+fleet*fleet_r+finls*finls_r+{o}{Cs['d_fp']}*fp_r"
                    f"+{o}{Cs['d_new']}*(sofr+new_r_sp)+int_cal*Thesis!{c}{T['orig']}/Thesis!{yc(2026)}{T['orig']}-{o}{Cs['cash']}*cash_y")
    put_row(CF, Cs, row, "int", "利息費用（淨）", "US$M", cells, "#,##0"); row += 1
    for y in YRS:
        if y >= 2026:
            PL[f"{yc(y)}{Ps['int']}"] = f"=CF!{yc(y)}{Cs['int']}"; PL[f"{yc(y)}{Ps['int']}"].number_format = "#,##0"
    PL[f"{yc(2025)}{Ps['int']}"] = "=fy25_int"
    row += 2
for j, s in enumerate(SC):
    col = "CDE"[j]; Cs = C[s]
    CF[f"{col}4"] = f"=SUM({yc(2027)}{Cs['gate']}:{yc(2031)}{Cs['gate']})"
    CF[f"{col}5"] = f"=SUM({yc(2027)}{Cs['need']}:{yc(2031)}{Cs['need']})"
    CF[f"{col}6"] = f"=MIN({yc(2027)}{Cs['cash']}:{yc(2031)}{Cs['cash']})"
    CF[f"{col}7"] = f"=SUM({yc(2027)}{Cs['dist']}:{yc(2031)}{Cs['dist']})"
    for rr_ in (4, 5, 6, 7): CF[f"{col}{rr_}"].number_format = "#,##0"
# 2025 應付：以 Q 頁 FY2025 期末 1,100（資產負債表）；修正為直接數值
for s in SC:
    CF[f"{yc(2025)}{C[s]['ap']}"] = 1100; CF[f"{yc(2025)}{C[s]['ap']}"].font = BLUE

# =====================================================================================
# Credit — 信用風險：公司債信、貸款去化、消費信用、盈餘品質、清償順位、壓力測試
# =====================================================================================
CR = wb.create_sheet("Credit")
CR["A1"] = "Credit — 信用風險六個角度：(1) 公司債信指標（三情境逐年）(2) 消費信用（管理組合）(3) 出售利得對損失與利差的敏感度 (4) 貸款去化與資金斷點 (5) 盈餘品質與關係人 (6) 清償順位與資產覆蓋；另附壓力格"
CR["A2"] = "評等門檻為 Moody's（2025-10）與 S&P（2025-03）公開說法；2026 年評等動作未找到。所有比率以調整後 EBITDA（公司口徑）計，另列扣除貸款出售利得的版本"
hdr(CR, 3, ["項目", "單位"] + [("FY%dA" % y if y == 2025 else "FY%dE" % y) for y in YRS])
K = {s: {} for s in SC}
row = 5
for s in SC:
    Ks = K[s]; Cs = C[s]; Ps = P[s]; T = R[s]
    CR.cell(row, 1, "(1) 公司債信指標｜" + SCN[s]).font = BOLD; row += 1
    yy = [y for y in YRS if y >= 2026]
    put_row(CR, Ks, row, "debt", "債務合計", "US$M", {yc(y): f"=CF!{yc(y)}{Cs['debt']}" for y in yy}, "#,##0"); row += 1
    put_row(CR, Ks, row, "cash", "現金", "US$M", {yc(y): f"=CF!{yc(y)}{Cs['cash']}" for y in yy}, "#,##0"); row += 1
    put_row(CR, Ks, row, "tra", "TRA 餘額（期末）", "US$M", {yc(y): f"=tra_0-tra_q4-SUM(CF!{yc(2027)}{Cs['tra']}:{yc(y)}{Cs['tra']})" if y >= 2027 else "=tra_0-tra_q4" for y in yy}, "#,##0"); row += 1
    put_row(CR, Ks, row, "ebitda", "調整後 EBITDA", "US$M", {yc(y): f"=PL!{yc(y)}{Ps['ebitda']}" for y in yy}, "#,##0"); row += 1
    put_row(CR, Ks, row, "ebitda_xg", "調整後 EBITDA 扣除貸款出售利得", "US$M", {yc(y): f"={yc(y)}{Ks['ebitda']}-Thesis!{yc(y)}{T['gos']}" for y in yy}, "#,##0"); row += 1
    put_row(CR, Ks, row, "lev", "債務／EBITDA", "x", {yc(y): f"={yc(y)}{Ks['debt']}/{yc(y)}{Ks['ebitda']}" for y in yy}, "0.00"); row += 1
    put_row(CR, Ks, row, "nlev", "淨債務／EBITDA", "x", {yc(y): f"=({yc(y)}{Ks['debt']}-{yc(y)}{Ks['cash']})/{yc(y)}{Ks['ebitda']}" for y in yy}, "0.00"); row += 1
    put_row(CR, Ks, row, "lev_xg", "債務／扣除出售利得的 EBITDA", "x", {yc(y): f"=IF({yc(y)}{Ks['ebitda_xg']}>0,{yc(y)}{Ks['debt']}/{yc(y)}{Ks['ebitda_xg']},\"EBITDA≤0\")" for y in yy}, "0.00"); row += 1
    put_row(CR, Ks, row, "lev_tra", "（債務＋TRA）／EBITDA", "x", {yc(y): f"=({yc(y)}{Ks['debt']}+{yc(y)}{Ks['tra']})/{yc(y)}{Ks['ebitda']}" for y in yy}, "0.00"); row += 1
    put_row(CR, Ks, row, "cov", "（EBITDA − 資本支出）／利息（利息為淨收入時標示）", "x", {yc(y): f"=IF(CF!{yc(y)}{Cs['int']}<=0,\"淨利息收入\",({yc(y)}{Ks['ebitda']}-CF!{yc(y)}{Cs['capex']})/CF!{yc(y)}{Cs['int']})" for y in yy}, "0.00"); row += 1
    put_row(CR, Ks, row, "focf", "自由現金流／債務", "%", {yc(y): f"=CF!{yc(y)}{Cs['fcf']}/{yc(y)}{Ks['debt']}" for y in yy}, "0.0%"); row += 1
    put_row(CR, Ks, row, "moody", "Moody's 升評門檻（兩項皆符合）", "—", {yc(y): f"=IF(AND({yc(y)}{Ks['lev']}<moody_lev,OR(CF!{yc(y)}{Cs['int']}<=0,({yc(y)}{Ks['ebitda']}-CF!{yc(y)}{Cs['capex']})/CF!{yc(y)}{Cs['int']}>moody_cov)),\"符合\",\"未符合\")" for y in yy}); row += 1
    put_row(CR, Ks, row, "sp", "S&P 正向門檻（兩項皆符合）", "—", {yc(y): f"=IF(AND({yc(y)}{Ks['lev']}<sp_lev,{yc(y)}{Ks['focf']}>sp_focf),\"符合\",\"未符合\")" for y in yy}); row += 1
    put_row(CR, Ks, row, "take", "已承諾貸款去化 ÷ 貸款發放（Ally＋整批出售，每 12 個月）", "%", {yc(y): f"=(ally_cap+fp_unused/fp_yrs)/Thesis!{yc(y)}{T['orig']}" for y in (2026, 2027)}, "0%"); row += 1
    put_row(CR, Ks, row, "take2", "已承諾＋證券化年化量 ÷ 貸款發放", "%", {yc(y): f"=(ally_cap+fp_unused/fp_yrs+abs_rr)/Thesis!{yc(y)}{T['orig']}" for y in (2026, 2027)}, "0%"); row += 1
    put_row(CR, Ks, row, "cush", "股權緩衝：現價市值 ÷ 債務（2026 年；市場給債權人的保護）", "x", {yc(2026): f"=px*sh_fx/{yc(2026)}{Ks['debt']}"}, "0.0"); row += 1
    row += 1
# (2) 消費信用
A_ = F["abs_managed"]
CR.cell(row, 1, "(2) 消費信用：Carvana 發放、Bridgecrest 服務的管理組合（含已出售）｜來源：" + ABS).font = BOLD; row += 1
hdr(CR, row, ["項目", "單位", "FY2021", "FY2022", "FY2023", "FY2024", "FY2025", "2025H1", "2026H1"]); S1 = {"hdr": row}; row += 1
put_row(CR, S1, row, "bal", "期末本金餘額（全部借款人）", "US$M", {get_column_letter(3 + i): v for i, v in enumerate([9613.0, 12262.1, 13528.3, 16210.5, 21570.4, 18463.2, 25959.6])}, "#,##0"); row += 1
put_row(CR, S1, row, "avg", "平均月底餘額（全部）", "US$M", {get_column_letter(3 + i): v for i, v in enumerate(A_["avg_all"] + A_["avg_h1_all"])}, "#,##0"); row += 1
put_row(CR, S1, row, "nco", "淨沖銷（全部）", "US$M", {get_column_letter(3 + i): v for i, v in enumerate(A_["nco_all"] + A_["nco_h1_all"])}, "#,##0.0"); row += 1
put_row(CR, S1, row, "avgp", "平均月底餘額（Prime：Deal Score ≥ 50）", "US$M", {get_column_letter(3 + i): v for i, v in enumerate(A_["avg_prime"] + A_["avg_h1_prime"])}, "#,##0"); row += 1
put_row(CR, S1, row, "ncop", "淨沖銷（Prime）", "US$M", {get_column_letter(3 + i): v for i, v in enumerate(A_["nco_prime"] + A_["nco_h1_prime"])}, "#,##0.0"); row += 1
for rr_ in (S1["bal"], S1["avg"], S1["nco"], S1["avgp"], S1["ncop"]):
    for i in range(7): CR.cell(rr_, 3 + i).font = BLUE
ann = lambda col: 2 if col in ("H", "I") else 1
put_row(CR, S1, row, "ncor", "年化淨沖銷率（全部）", "%", {c: f"={c}{S1['nco']}*{ann(c)}/{c}{S1['avg']}" for c in "CDEFGHI"}, "0.00%"); row += 1
put_row(CR, S1, row, "ncorp", "年化淨沖銷率（Prime）", "%", {c: f"={c}{S1['ncop']}*{ann(c)}/{c}{S1['avgp']}" for c in "CDEFGHI"}, "0.00%"); row += 1
put_row(CR, S1, row, "ncorn", "年化淨沖銷率（非 Prime，推得＝全部 − Prime）", "%", {c: f"=({c}{S1['nco']}-{c}{S1['ncop']})*{ann(c)}/({c}{S1['avg']}-{c}{S1['avgp']})" for c in "CDEFGHI"}, "0.00%"); row += 1
put_row(CR, S1, row, "lag", "落後 12 個月淨沖銷率（年化淨沖銷 ÷ 一年前餘額；排除快速成長的分母稀釋）", "%", {"D": f"=D{S1['nco']}/C{S1['bal']}", "E": f"=E{S1['nco']}/D{S1['bal']}", "F": f"=F{S1['nco']}/E{S1['bal']}", "G": f"=G{S1['nco']}/F{S1['bal']}", "I": f"=I{S1['nco']}*2/H{S1['bal']}"}, "0.00%"); row += 1
put_row(CR, S1, row, "dq", "61 天以上逾期率（全部，期末）", "%", {get_column_letter(3 + i): v / 100 for i, v in enumerate(A_["dq61_all_fy"] + A_["dq61_all"])}, "0.00%"); row += 1
put_row(CR, S1, row, "g", "管理組合餘額年增", "%", {"D": f"=D{S1['bal']}/C{S1['bal']}-1", "E": f"=E{S1['bal']}/D{S1['bal']}-1", "F": f"=F{S1['bal']}/E{S1['bal']}-1", "G": f"=G{S1['bal']}/F{S1['bal']}-1", "I": f"=I{S1['bal']}/H{S1['bal']}-1"}, "0.0%"); row += 1
CR.cell(row, 1, "註：2025-12 起承銷不再使用 FICO（改用 VantageScore 與自有 CarvanaScore 第八代）；Prime 定義為自有 Deal Score ≥ 50，非 FICO 口徑（同一公開說明書）"); row += 2
# (3) 出售利得敏感度
CR.cell(row, 1, "(3) 出售利得率的信用與利差傳導（情境 C 推導）").font = BOLD; row += 1
S2 = {}
put_row(CR, S2, row, "b", "情境 B 2027 年利得率", "%", {"C": "=gos27_B"}, "0.00%"); row += 1
put_row(CR, S2, row, "l", "− 損失傳導：loss_pt × 終身損失增幅", "pt", {"C": "=-loss_pt*dcnl_C"}, "0.00%"); row += 1
put_row(CR, S2, row, "d", "− 利差傳導：存續期間 × 利差擴大", "pt", {"C": "=-dur*dspr_C"}, "0.00%"); row += 1
put_row(CR, S2, row, "c", "情境 C 2027 年利得率", "%", {"C": f"=C{S2['b']}+C{S2['l']}+C{S2['d']}"}, "0.00%"); row += 1
put_row(CR, S2, row, "per1", "利得率每降 1 pt，2027 年調整後 EBITDA 減少（情境 B 發放量）", "US$M", {"C": f"=Thesis!{yc(2027)}{R['B']['orig']}*0.01"}, "#,##0"); row += 1
put_row(CR, S2, row, "per1s", "　占情境 B 2027 年調整後 EBITDA", "%", {"C": f"=C{S2['per1']}/PL!{yc(2027)}{P['B']['ebitda']}"}, "0.0%"); row += 2
# (4) 資金斷點
CR.cell(row, 1, "(4) 貸款去化與資金斷點（2026-09-30 起算；若買方與 ABS 市場停止承接，放款先停在倉儲額度、再用現金）").font = BOLD; row += 1
S3 = {}
put_row(CR, S3, row, "o", "每月貸款發放（Q3 2026 ÷ 3）", "US$M", {"C": "=lq_orig/3"}, "#,##0"); row += 1
put_row(CR, S3, row, "op", "每月營運現金（扣除出售利得；Q3 調整後 EBITDA − 利得 − 利息 − 資本支出，÷ 3）", "US$M", {"C": "=(lq_ebitda-lq_gos-lq_int-lq_capex)/3"}, "#,##0"); row += 1
put_row(CR, S3, row, "avail", "可動用現金（現金 − 最低營運現金）", "US$M", {"C": "=lq_cash-min_cash"}, "#,##0"); row += 1
put_row(CR, S3, row, "whl", "倉儲額度可容納的貸款（未用額度 ÷ 墊款成數）", "US$M", {"C": "=wh_unused/wh_adv"}, "#,##0"); row += 1
CR.cell(row, 1, "情境").font = BOLD; CR.cell(row, 3, "全部停止").font = BOLD; CR.cell(row, 4, "Ally 照常承購").font = BOLD; S3["h"] = row; row += 1
put_row(CR, S3, row, "on", "每月須自行持有的新貸款", "US$M", {"C": f"=C{S3['o']}", "D": f"=MAX(0,C{S3['o']}-ally_cap/12)"}, "#,##0"); row += 1
put_row(CR, S3, row, "m1", "倉儲額度用盡月數", "月", {c: f"=IF({c}{S3['on']}>0,C{S3['whl']}/{c}{S3['on']},99)" for c in "CD"}, "0.0"); row += 1
put_row(CR, S3, row, "b1", "倉儲期間每月現金消耗（自備款 − 營運現金）", "US$M", {c: f"={c}{S3['on']}*(1-wh_adv)-C{S3['op']}" for c in "CD"}, "#,##0"); row += 1
put_row(CR, S3, row, "b2", "倉儲用盡後每月現金消耗", "US$M", {c: f"={c}{S3['on']}-C{S3['op']}" for c in "CD"}, "#,##0"); row += 1
put_row(CR, S3, row, "runway", "資金斷點：可撐月數（不減少放款）", "月", {c: (f"=IF({c}{S3['b1']}<=0,IF({c}{S3['b2']}<=0,99,{c}{S3['m1']}+C{S3['avail']}/{c}{S3['b2']}),"
                                                                          f"IF(C{S3['avail']}<={c}{S3['m1']}*{c}{S3['b1']},C{S3['avail']}/{c}{S3['b1']},{c}{S3['m1']}+(C{S3['avail']}-{c}{S3['m1']}*{c}{S3['b1']})/{c}{S3['b2']}))") for c in "CD"}, "0.0"); row += 1
CR.cell(row, 1, "讀法：撐不過即須減少自營放款（改由第三方放款或現金買車），零售台數與其他毛利同步下降；倉儲額度到期日分布 2026-10 至 2027-11（10-Q 附註 9），續約本身亦受信用循環影響"); row += 2
# (5) 盈餘品質
CR.cell(row, 1, "(5) 盈餘品質與關係人（季度；Q 頁）").font = BOLD; row += 1
hdr(CR, row, ["項目", "單位", "Q4'24", "Q1'25", "Q2'25", "Q3'25", "Q4'25", "Q1'26", "Q2'26", "Q3'26", "FY2025", "TTM"]); S4 = {"hdr": row}; row += 1
for key, lab, unit, fmt in (("gos_sh", "出售利得 ÷ 調整後 EBITDA", "%", "0.0%"), ("gosrp_sh", "（出售利得＋關係人其他營收）÷ 調整後 EBITDA", "%", "0.0%"),
                            ("ebitda_xg", "調整後 EBITDA 扣除出售利得", "US$M", "#,##0"), ("ebitda_xg_m", "扣除出售利得的 EBITDA 率", "%", "0.0%"), ("gos_m", "出售利得率", "%", "0.00%")):
    cells = {get_column_letter(3 + i): f"=Q!{get_column_letter(3 + i)}{QR[key]}" for i in range(8)}
    cells["K"] = f"=Q!M{QR[key]}"
    put_row(CR, S4, row, key, lab, unit, cells, fmt); row += 1
CR[f"L{S4['gosrp_sh']}"] = "=ttm_gosrp"; CR[f"L{S4['gosrp_sh']}"].number_format = "0.0%"
put_row(CR, S4, row, "gotham", "對照：Gotham City Research（放空方）稱 2025 年前三季此比例約 75%", "%", {"K": 0.75}, "0%"); CR[f"K{row}"].font = BLUE; row += 1
put_row(CR, S4, row, "lev_xg", "債務／扣除出售利得的 EBITDA（TTM，2026-09-30）", "x", {"L": "=ttm_lev_xg"}, "0.00"); row += 1
CR.cell(row, 1, "關係人：VSC 由 DriveTime（Garcia 家族控制）承保與管理；Bridgecrest（同一家族）服務 Carvana 發放的貸款；TRA 負債 2,130 中 1,645 應付關係人（10-Q 附註 6）。2025-01 Hindenburg、2026-01 Gotham 指控未揭露關係人交易，公司否認；2025 年 10-K 如期申報、會計師 Grant Thornton 未更換"); row += 2
# (6) 清償順位與資產覆蓋
CR.cell(row, 1, "(6) 清償順位與資產覆蓋（2026-09-30；折扣率為判斷）").font = BOLD; row += 1
hdr(CR, row, ["資產", "單位", "帳面", "回收折扣後比例", "可回收"]); S5 = {"hdr": row}; row += 1
assets = (("a_cash", "現金", "=lq_cash", 1.0), ("a_fr", "待售貸款", "=lq_frhfs", 0.9), ("a_inv", "存貨", "=lq_inv", 0.75), ("a_bi", "證券化受益權", "=lq_bi", 0.6),
          ("a_ar", "應收帳款", "=320", 0.8), ("a_ppe", "不動產與設備", "=2914", 0.5))
for k, lab, v, h in assets:
    put_row(CR, S5, row, k, lab, "US$M", {"C": v, "D": h, "E": f"=C{row}*D{row}"}, "#,##0"); CR[f"D{row}"].number_format = "0%"; CR[f"D{row}"].font = BLUE; row += 1
put_row(CR, S5, row, "a_tot", "可回收資產合計（不含遞延所得稅資產 2,905）", "US$M", {"C": f"=SUM(C{S5['a_cash']}:C{S5['a_ppe']})", "E": f"=SUM(E{S5['a_cash']}:E{S5['a_ppe']})"}, "#,##0"); row += 1
hdr(CR, row, ["求償（依順位）", "單位", "金額", "累計", "累計覆蓋（可回收 ÷ 累計求償）"]); row += 1
claims = (("c_abf", "資產擔保融資（floor plan、受益權融資、不動產、車隊、融資租賃）", "=71+427+re_fin+fleet+finls"),
          ("c_sec", "擔保債（定期貸款 B＋2031 年票據）", "=tlb+ssn_31"), ("c_uns", "無擔保票據", "=uns"), ("c_tra", "TRA（應收稅款協議；破產時一般無擔保債權）", "=tra_0"))
prev = None
for k, lab, v in claims:
    put_row(CR, S5, row, k, lab, "US$M", {"C": v, "D": (f"=D{prev}+C{row}" if prev else f"=C{row}"), "E": f"=$E${S5['a_tot']}/D{row}"}, "#,##0"); CR[f"E{row}"].number_format = "0.00\"x\""; prev = row; row += 1
row += 1
# (7) 壓力格：情境 B 2027 年債務／EBITDA
CR.cell(row, 1, "(7) 壓力格：情境 B 2027 年債務／調整後 EBITDA（列＝利得率變化、欄＝台數變化；變動 SG&A 隨台數、固定不變；債務不變）").font = BOLD; row += 1
S6 = {"hdr": row}
dus = [-0.2, -0.1, 0.0, 0.1]; dgs = [-0.04, -0.03, -0.02, -0.01, 0.0]
CR.cell(row, 1, "利得率變化 ＼ 台數變化"); CR.cell(row, 1).font = BOLD
for j, du in enumerate(dus):
    c = CR.cell(row, 3 + j, du); c.number_format = "+0%;-0%;0%"; c.font = BOLD
row += 1
B = R["B"]; e27 = yc(2027)
contrib = f"(Thesis!{e27}{B['gpu']}-Thesis!{e27}{B['sga_var_pu']})"
for dg in dgs:
    CR.cell(row, 1, dg).number_format = "+0.0%;-0.0%;0.0%"; CR.cell(row, 1).font = BOLD
    for j, du in enumerate(dus):
        col = get_column_letter(3 + j)
        CR[f"{col}{row}"] = (f"=Credit!{e27}{K['B']['debt']}/(PL!{e27}{P['B']['ebitda']}+Thesis!{e27}{B['u']}*{col}${S6['hdr']}*{contrib}/1000000"
                             f"+Thesis!{e27}{B['u']}*(1+{col}${S6['hdr']})*Thesis!{e27}{B['opu']}*$A{row}/1000000)")
        CR[f"{col}{row}"].number_format = "0.00"
    row += 1
S6["end"] = row - 1
CR.column_dimensions["A"].width = 60
ROWMAP["Credit"] = {"K": K, "S1": S1, "S2": S2, "S3": S3, "S4": S4, "S5": S5, "S6": S6}

# =====================================================================================
# Val — 估值：DCF（UFCF）＋倍數法 → 情境目標價 → 機率加權 → 五級評等
# =====================================================================================
V = wb.create_sheet("Val")
V["A1"] = "Val — 估值。DCF：FY2027–FY2035 無槓桿自由現金流＋永續價值（以 2026 年底為基準點折現）；倍數法：2027 年調整後 EBITDA × 同業中位數 × 成長溢價。股權＝EV − 淨債務（2026-09-30）− TRA 負債"
hdr(V, 3, ["項目", "單位", "值", "", "說明"])
VR = {}
vrows = [("ke", "股權成本 ke（CAPM：rf＋β × ERP）", "%", "=rf+beta*erp", "0.00%", "β 為 Blume 調整"),
         ("E", "股權市值（現價 × 完全轉換股數）", "US$M", "=px*sh_fx", "#,##0", ""),
         ("D", "債務（2026-09-30，含融資租賃）", "US$M", "=lq_debt_bs", "#,##0", ""),
         ("wacc", "WACC", "%", "=C5*C6/(C6+C7)+kd*(1-t_val)*C7/(C6+C7)", "0.00%", "市值權重"),
         ("nd", "淨債務（2026-09-30）", "US$M", "=lq_debt_bs-lq_cash", "#,##0", ""),
         ("tra", "TRA 負債（視同債務）", "US$M", "=tra_0", "#,##0", "遞延所得稅資產的價值已反映在現金稅率，不另加"),
         ("peer", "同業 EV/EBITDA 中位數（LAD、AN）", "x", "=MEDIAN(peer_lad,peer_an)", "0.00", ""),
         ("mult", "適用倍數（中位數 × 成長溢價）", "x", "=C11*prem", "0.00", ""),
         ("mkt_ev", "市場隱含 EV（股權市值＋淨債務＋TRA）", "US$M", "=C6+C9+C10", "#,##0", ""),
         ("mkt_mult", "市場隱含 EV ÷ 近四季調整後 EBITDA", "x", "=C13/ttm_ebitda", "0.0", "")]
rv = 5
for k, lab, u, f, fmt, note in vrows:
    V.cell(rv, 1, lab); V.cell(rv, 2, u); V.cell(rv, 3, f).number_format = fmt; V.cell(rv, 5, note); VR[k] = rv; dn("v_" + k, f"Val!$C${rv}"); rv += 1
V["C8"] = "=C5*C6/(C6+C7)+kd*(1-t_val)*C7/(C6+C7)"
rv += 1
VS = {s: {} for s in SC}
hdr(V, rv, ["DCF｜項目", "單位"] + [("FY%dE" % y) for y in FC]); V.cell(rv, 1).value = "DCF（各情境）"; VR["dcf_hdr"] = rv; rv += 1
FCc = lambda y: get_column_letter(3 + y - 2027)   # Val 頁年度欄：C＝2027 … K＝2035
for s in SC:
    Vs = VS[s]; Ps = P[s]; Cs = C[s]
    V.cell(rv, 1, SCN[s]).font = BOLD; rv += 1
    def vput(key, lab, unit, f, fmt="#,##0"):
        global rv
        V.cell(rv, 1, lab); V.cell(rv, 2, unit)
        for y in FC:
            V[f"{FCc(y)}{rv}"] = f(y); V[f"{FCc(y)}{rv}"].number_format = fmt
        Vs[key] = rv; rv += 1
    vput("ebit", "營業利益", "US$M", lambda y: f"=PL!{yc(y)}{Ps['oi']}")
    vput("tax", "− 營業利益的現金稅（2027–2028 用 ctr，之後用有效稅率）", "US$M", lambda y: (f"=-MAX(0,PL!{yc(y)}{Ps['oi']})*ctr" if y <= 2028 else f"=-MAX(0,PL!{yc(y)}{Ps['oi']})*" + ("(etr_27+(etr_30-etr_27)*2/3)" if y == 2029 else "etr_30")))
    vput("da", "＋折舊攤銷＋SBC", "US$M", lambda y: f"=CF!{yc(y)}{Cs['da']}")
    vput("capex", "− 資本支出", "US$M", lambda y: f"=-CF!{yc(y)}{Cs['capex']}")
    vput("wc", "− 營運資金與風險自留", "US$M", lambda y: f"=CF!{yc(y)}{Cs['dwc']}+CF!{yc(y)}{Cs['rr']}")
    vput("ufcf", "無槓桿自由現金流", "US$M", lambda y: f"=SUM({FCc(y)}{Vs['ebit']}:{FCc(y)}{Vs['wc']})")
    vput("df", "折現因子", "x", lambda y: f"=1/(1+$C${VR['wacc']})^{y-2026}", "0.000")
    vput("pv", "現值", "US$M", lambda y: f"={FCc(y)}{Vs['ufcf']}*{FCc(y)}{Vs['df']}")
    V.cell(rv, 1, "永續價值現值"); V.cell(rv, 2, "US$M"); V[f"C{rv}"] = f"=K{Vs['ufcf']}*(1+g_tv)/($C${VR['wacc']}-g_tv)*K{Vs['df']}"; V[f"C{rv}"].number_format = "#,##0"; Vs["tv"] = rv; rv += 1
    V.cell(rv, 1, "EV（DCF）"); V.cell(rv, 2, "US$M"); V[f"C{rv}"] = f"=SUM(C{Vs['pv']}:K{Vs['pv']})+C{Vs['tv']}"; V[f"C{rv}"].number_format = "#,##0"; Vs["ev"] = rv; rv += 1
    V.cell(rv, 1, "永續價值占 EV"); V.cell(rv, 2, "%"); V[f"C{rv}"] = f"=C{Vs['tv']}/C{Vs['ev']}"; V[f"C{rv}"].number_format = "0%"; Vs["tvsh"] = rv; rv += 1
    V.cell(rv, 1, "永續價值隱含 2035 年 EV／調整後 EBITDA（對照倍數法的適用倍數）"); V.cell(rv, 2, "x"); V[f"C{rv}"] = f"=C{Vs['tv']}/K{Vs['df']}/PL!{yc(2035)}{Ps['ebitda']}"; V[f"C{rv}"].number_format = "0.0"; Vs["tvm"] = rv; rv += 1
    V.cell(rv, 1, "每股價值（DCF）"); V.cell(rv, 2, "US$"); V[f"C{rv}"] = f"=(C{Vs['ev']}-$C${VR['nd']}-$C${VR['tra']})/sh_fx"; V[f"C{rv}"].number_format = "0.00"; Vs["ps_dcf"] = rv; rv += 1
    V.cell(rv, 1, "每股價值（倍數法：2027 年調整後 EBITDA × 適用倍數）"); V.cell(rv, 2, "US$"); V[f"C{rv}"] = f"=(PL!{yc(2027)}{Ps['ebitda']}*$C${VR['mult']}-$C${VR['nd']}-$C${VR['tra']})/sh_fx"; V[f"C{rv}"].number_format = "0.00"; Vs["ps_mult"] = rv; rv += 1
    V.cell(rv, 1, "情境目標價（DCF × w＋倍數法 ×（1 − w））"); V.cell(rv, 2, "US$"); V[f"C{rv}"] = f"=w_dcf*C{Vs['ps_dcf']}+(1-w_dcf)*C{Vs['ps_mult']}"; V[f"C{rv}"].number_format = "0.00"; Vs["tp"] = rv; rv += 1
    V.cell(rv, 1, "現金流序列（C＝2026 年底以市場 EV 買入；D–L＝FY2027–FY2035 無槓桿自由現金流；L 另加 2035 年以適用倍數出場）"); V.cell(rv, 2, "US$M")
    V[f"C{rv}"] = f"=-$C${VR['mkt_ev']}"
    for y in FC:
        col = get_column_letter(4 + y - 2027)
        V[f"{col}{rv}"] = f"={FCc(y)}{Vs['ufcf']}" + (f"+PL!{yc(2035)}{Ps['ebitda']}*$C${VR['mult']}" if y == 2035 else "")
        V[f"{col}{rv}"].number_format = "#,##0"
    V[f"C{rv}"].number_format = "#,##0"; Vs["irr_cf"] = rv; rv += 1
    V.cell(rv, 1, "市場隱含報酬率（以現價買入、持有至 2035 年的 IRR）"); V.cell(rv, 2, "%"); V[f"C{rv}"] = f"=IRR(C{Vs['irr_cf']}:L{Vs['irr_cf']})"; V[f"C{rv}"].number_format = "0.0%"; Vs["irr"] = rv; rv += 1
    rv += 1
# 彙總
VR["sum_hdr"] = rv; hdr(V, rv, ["彙總", "單位", "情境 A", "情境 B", "情境 C"]); rv += 1
for k, lab, f, fmt in (("s_dcf", "每股價值（DCF）", "ps_dcf", "0.00"), ("s_mult", "每股價值（倍數法）", "ps_mult", "0.00"), ("s_tp", "情境目標價", "tp", "0.00"), ("s_tvsh", "永續價值占 EV", "tvsh", "0%"), ("s_tvm", "永續價值隱含 2035 年 EV／EBITDA", "tvm", "0.0"), ("s_irr", "市場隱含報酬率（IRR，對照 WACC）", "irr", "0.0%")):
    V.cell(rv, 1, lab); V.cell(rv, 2, {"0%": "%", "0.0%": "%", "0.0": "x"}.get(fmt, "US$"))
    for j, s in enumerate(SC):
        V.cell(rv, 3 + j, f"=C{VS[s][f]}").number_format = fmt
    VR[k] = rv; rv += 1
V.cell(rv, 1, "情境機率"); V.cell(rv, 2, "%"); V.cell(rv, 3, "=p_A").number_format = "0%"; V.cell(rv, 4, "=p_B").number_format = "0%"; V.cell(rv, 5, "=p_C").number_format = "0%"; VR["prob"] = rv; rv += 1
V.cell(rv, 1, "機率加權市場隱含報酬率（IRR）"); V.cell(rv, 2, "%"); V.cell(rv, 3, f"=SUMPRODUCT(C{VR['s_irr']}:E{VR['s_irr']},C{VR['prob']}:E{VR['prob']})").number_format = "0.0%"; VR["irr_w"] = rv; dn("v_irr_w", f"Val!$C${rv}"); rv += 1
rv += 1
VR["res_hdr"] = rv; hdr(V, rv, ["結論", "單位", "值", "", "說明"]); rv += 1
res = [("tp", "加權目標價", "US$", f"=SUMPRODUCT(C{VR['s_tp']}:E{VR['s_tp']},C{VR['prob']}:E{VR['prob']})", "0.00", "情境 25／50／25"),
       ("px", "現價（2026-10-28 收盤，推得）", "US$", "=px", "0.00", ""),
       ("up", "上檔空間（加權目標價 ÷ 現價 − 1）", "%", f"=C{rv}/C{rv+1}-1", "+0.0%;-0.0%", ""),
       ("dcf_w", "加權每股價值：DCF", "US$", f"=SUMPRODUCT(C{VR['s_dcf']}:E{VR['s_dcf']},C{VR['prob']}:E{VR['prob']})", "0.00", ""),
       ("mult_w", "加權每股價值：倍數法", "US$", f"=SUMPRODUCT(C{VR['s_mult']}:E{VR['s_mult']},C{VR['prob']}:E{VR['prob']})", "0.00", ""),
       ("an", "分析師平均目標價（對照）", "US$", "=an_tgt", "0.00", "TIKR 2026-10-13"),
       ("req30", "市場隱含：2030 年須達到的調整後 EBITDA（市場 EV 以 WACC 滾至 2030 年 ÷ 適用倍數）", "US$M", f"=C{VR['mkt_ev']}*(1+C{VR['wacc']})^4/C{VR['mult']}", "#,##0", ""),
       ("b30", "情境 B 2030 年調整後 EBITDA", "US$M", f"=PL!{yc(2030)}{P['B']['ebitda']}", "#,##0", ""),
       ("req_u", "市場隱含：2030 年零售台數（以情境 B 2030 年每台 EBITDA 換算）", "輛", f"=C{rv+6}/(C{rv+7}/Thesis!{yc(2030)}{R['B']['u']})", "#,##0", ""),
       ("rec", "投資建議", "—", None, "@", "五級固定規則：≥ +50% 強力買進、+25%～+50% 買進、−25%～+25% 中立、−50%～−25% 放空、≤ −50% 強力放空")]
for k, lab, u, f, fmt, note in res:
    V.cell(rv, 1, lab); V.cell(rv, 2, u)
    if f: V.cell(rv, 3, f).number_format = fmt
    V.cell(rv, 5, note); VR[k] = rv; rv += 1
V[f"C{VR['up']}"] = f"=C{VR['tp']}/C{VR['px']}-1"
V[f"C{VR['req_u']}"] = f"=C{VR['req30']}/(C{VR['b30']}/Thesis!{yc(2030)}{R['B']['u']})"
V[f"C{VR['rec']}"] = f'=IF(C{VR["up"]}>=0.5,"強力買進",IF(C{VR["up"]}>=0.25,"買進",IF(C{VR["up"]}>-0.25,"中立",IF(C{VR["up"]}>-0.5,"放空","強力放空"))))'
V[f"C{VR['rec']}"].fill = KEY
dn("v_tp_w", f"Val!$C${VR['tp']}"); dn("v_up", f"Val!$C${VR['up']}")
V.column_dimensions["A"].width = 64
ROWMAP["Val"] = {"VR": VR, "VS": VS}

# =====================================================================================
# Recon — 對帳：公司指引、共識、評等機構、放空報告
# =====================================================================================
RC = wb.create_sheet("Recon")
RC["A1"] = "Recon — 對帳（不作校準）：模型值與公司指引、市場共識、評等機構指標、放空報告主張逐列比較"
hdr(RC, 3, ["項目", "單位", "模型（情境 B）／實績", "外部值", "差距", "外部來源", "標記", "說明"])
B = P["B"]; KB = K["B"]
recs = [
    ("2026 年調整後 EBITDA vs 公司指引「區間高端」（取 2.85–3.0B 中點）", "US$M", f"=PL!{yc(2026)}{B['ebitda']}", "=(guide_lo+guide_hi)/2+75", PR3 + "：2.7–3.0B 的 higher end", "Interested-party", "共識約 2.95B（TIKR 2026-10-13）"),
    ("2026 年調整後 EBITDA vs 市場共識", "US$M", f"=PL!{yc(2026)}{B['ebitda']}", 2950, "TIKR（2026-10-13）：全年約 2.95B（Q3 約 795、Q4 約 710）", "Interested-party", "—"),
    ("Q3 2026 零售台數：實績 vs 共識", "輛", f"={q('units','J')}", "=cons_q3_units", "Zacks（2026-10-23）", "Interested-party", "實績高 3.4%"),
    ("Q3 2026 營收：實績 vs 共識", "US$M", f"={q('rev_tot','J')}", "=cons_q3_rev", "Zacks（2026-10-23）", "Interested-party", "—"),
    ("Q3 2026 調整後 EBITDA：實績 vs 共識", "US$M", f"={q('adj_ebitda','J')}", "=cons_q3_ebitda", "StockStory（2026-10-28）", "Interested-party", "—"),
    ("Q3 2026 零售 GPU：實績 vs 共識", "US$", f"={q('gp_ret','J')}*1000000/{q('units','J')}", "=cons_q3_gpr", "Zacks", "Interested-party", "—"),
    ("Q3 2026 其他 GPU：實績 vs 共識", "US$", f"={q('rev_oth','J')}*1000000/{q('units','J')}", "=cons_q3_gpo", "Zacks", "Interested-party", "利率上升下仍持平"),
    ("Q3 2026 批發 GPU：實績 vs 共識", "US$", f"={q('gp_whs','J')}*1000000/{q('units','J')}", "=cons_q3_gpw", "Zacks", "Interested-party", "—"),
    ("情境 B 2030 年營業利益率 vs 公司中期目標（300 萬台）", "%", f"=PL!{yc(2030)}{B['oi_m']}", "=om_tgt", PR3, "Interested-party", "公司目標對應 300 萬台；情境 B 2030 年約 160 萬台"),
    ("債務／EBITDA（2026 年底 vs 2026-09-30 TTM）", "x", f"=Credit!{yc(2026)}{KB['lev']}", "=ttm_lev", "Q 頁 TTM", "Derived", "—"),
    ("淨債務／EBITDA（2026 年底 vs 2026-09-30 TTM）", "x", f"=Credit!{yc(2026)}{KB['nlev']}", "=ttm_nlev", "Q 頁 TTM（公司 Q2 口徑 1.0x）", "Derived", "—"),
    ("（EBITDA − 資本支出）／利息（2026 年 vs Moody's 升評門檻）", "x", f"=Credit!{yc(2026)}{KB['cov']}", "=moody_cov", "Moody's 2025-10-02", "Interested-party", "2025-06 時為約 2.8x"),
    ("（出售利得＋關係人其他營收）÷ 調整後 EBITDA（TTM vs Gotham 主張）", "%", "=ttm_gosrp", 0.75, "Gotham City Research（2026-01-28，經 Hedge Fund Alpha 報導）", "Interested-party", "放空方；本模型以公司揭露數字推得，口徑可能不同"),
    ("加權目標價 vs 分析師平均", "US$", "=v_tp_w", "=an_tgt", "TIKR（2026-10-13）", "Interested-party", "—"),
    ("市場隱含 EV ÷ TTM 調整後 EBITDA vs 同業中位數", "x", f"=Val!C{VR['mkt_mult']}", f"=Val!C{VR['peer']}", "stockanalysis.com（LAD、AN，近四季）", "Verified", "—"),
]
rr_ = 4; RCR = {}
for i, (lab, u, m, e, src, tag, note) in enumerate(recs):
    RC.cell(rr_, 1, lab); RC.cell(rr_, 2, u); RC.cell(rr_, 3, m); c = RC.cell(rr_, 4, e)
    if not (isinstance(e, str) and e.startswith("=")): c.font = BLUE
    RC.cell(rr_, 5, f"=C{rr_}-D{rr_}"); RC.cell(rr_, 6, src); RC.cell(rr_, 7, tag); RC.cell(rr_, 8, note)
    fmt = {"US$M": "#,##0", "輛": "#,##0", "x": "0.00", "%": "0.0%", "US$": "0.00"}[u]
    for cc in (3, 4, 5): RC.cell(rr_, cc).number_format = fmt
    RCR[i] = rr_; rr_ += 1
RC.cell(rr_ + 1, 1, "未找到：2026 年 Moody's／S&P 評等動作；2031 年票據贖回價格表；ADESA 全面改建成本；TRA 償付時程；非 Prime 貸款的終身損失預期（144A 發行未公開）")
ROWMAP["Recon"] = RCR
RC.column_dimensions["A"].width = 60

# =====================================================================================
# Sens — 敏感度（建置時完整重算；scripts/sens.py 產生 data/sens_v02.json）
# =====================================================================================
SE = wb.create_sheet("Sens")
SE["A1"] = "Sens — 一次動一個輸入（低／高），完整重算整本模型；輸出：加權目標價、情境 B 2027 年債務／EBITDA、情境 C 收支關卡（FY2027–FY2031 合計）"
hdr(SE, 3, ["輸入", "鍵", "低值", "高值", "基準：加權目標價", "低值：目標價變化", "高值：目標價變化", "低值：B 2027 債務／EBITDA", "高值：B 2027 債務／EBITDA", "低值：C 收支關卡", "高值：C 收支關卡"])
SENS_KEYS = ["q4_seq", "g27_B", "g28_B", "g30_B", "g35_B", "gpr27_B", "gprLT_B", "gos27_B", "gosLT_B", "sga_fix", "sga_var", "eff_B", "capcost", "util",
             "dcnl_C", "dspr_C", "dur", "etr_30", "ctr", "tra_yrs", "beta", "g_tv", "prem", "inv_d", "refi_sp", "gprLT_C", "gosLT_C", "asp_g"]
for i, k in enumerate(SENS_KEYS):
    r_ = 4 + i
    lab = next((it[1] for it in rows if it[0] == k), k)
    SE.cell(r_, 1, lab); SE.cell(r_, 2, k)
    if SENS and k in SENS["cases"]:
        cs = SENS["cases"][k]; base = SENS["base"]
        SE.cell(r_, 3, cs["lo_val"]); SE.cell(r_, 4, cs["hi_val"])
        SE.cell(r_, 5, base["tp"]).number_format = "0.00"
        SE.cell(r_, 6, cs["lo"]["tp"] / base["tp"] - 1).number_format = "+0.0%;-0.0%"
        SE.cell(r_, 7, cs["hi"]["tp"] / base["tp"] - 1).number_format = "+0.0%;-0.0%"
        SE.cell(r_, 8, cs["lo"]["lev"]).number_format = "0.00"; SE.cell(r_, 9, cs["hi"]["lev"]).number_format = "0.00"
        SE.cell(r_, 10, cs["lo"]["gateC"]).number_format = "#,##0"; SE.cell(r_, 11, cs["hi"]["gateC"]).number_format = "#,##0"
SE.cell(5 + len(SENS_KEYS), 1, "基準值：加權目標價（美元）"); SE.cell(5 + len(SENS_KEYS), 3, SENS["base"]["tp"] if SENS else None)
SE.cell(6 + len(SENS_KEYS), 1, "基準值：情境 B 2027 年債務／EBITDA"); SE.cell(6 + len(SENS_KEYS), 3, SENS["base"]["lev"] if SENS else None)
SE.cell(7 + len(SENS_KEYS), 1, "基準值：情境 C 收支關卡（FY2027–FY2031 合計）"); SE.cell(7 + len(SENS_KEYS), 3, SENS["base"]["gateC"] if SENS else None)
SE.column_dimensions["A"].width = 52
ROWMAP["Sens"] = {"first": 4, "last": 3 + len(SENS_KEYS)}

# =====================================================================================
# Sources
# =====================================================================================
SO = wb.create_sheet("Sources")
SO["A1"] = "Sources — 出處（擷取日 2026-10-27／28）"
hdr(SO, 3, ["序", "出處", "內容", "標記"])
srcs = [
    ("SEC EDGAR：Carvana 8-K 附件 99.1 股東信 Q4 2024–Q2 2026（CIK 1690820）", "季度營收線、毛利、台數、調整後 EBITDA 對帳、流動性", "Interested-party"),
    ("SEC EDGAR：Carvana 10-Q Q2 2026（2026-07-29）附註 6–10、19", "關係人、貸款出售協議（Ally、整批出售、證券化）、受益權、債務、股權", "Interested-party"),
    ("SEC EDGAR：Carvana 10-K FY2025（2026-02-18）", "評價備抵釋放、TRA、會計師 Grant Thornton", "Interested-party"),
    ("SEC EDGAR：8-K 2026-08-14（定期貸款 B 信用協議）；Carvana 新聞稿 2026-08-12", "1.66B、SOFR＋2.25%、2033 到期、贖回 2030 年票據；淨債務／EBITDA 1.0x", "Interested-party"),
    ("SEC EDGAR：8-K 2026-05-06（章程修正，1 拆 5）", "股票分割", "Interested-party"),
    ("SEC EDGAR：Carvana Auto Receivables Trust 2026-P3 424B5（2026-08-20，CIK 1770373）", "管理組合逾期與淨沖銷（全部／Prime）、承銷模型變更", "Interested-party"),
    ("SEC XBRL companyfacts CIK0001690820", "季末資產負債表、2024 年九個月現金流", "Interested-party"),
    ("stockanalysis.com：CVNA、LAD、AN、KMX 統計頁（2026-10-23～27）", "現價、Beta、EV/EBITDA", "Verified"),
    ("sofrrate.com 2026-10 SOFR 每日表", "SOFR 3.93%（2026-10-26）", "Verified"),
    ("Yahoo Finance／Zacks（2026-10-23）Q3 2026 共識", "台數 210,299、營收 7.67B、EPS 0.47、各 GPU", "Interested-party"),
    ("TIKR（2026-10-13）", "2026 年 EBITDA 共識約 2.95B、目標價平均約 82", "Interested-party"),
    ("Investing.com：Moody's 升 B2（2025-10-02）、S&P 升 B（2025-03）", "評等門檻與槓桿指標", "Interested-party"),
    ("Hedge Fund Alpha（2026-01-28）、24/7 Wall St.（2026-01-29）、Yahoo Finance（2026-02-12）", "Gotham City Research 放空報告主張與公司回應", "Interested-party"),
    ("Cox Automotive Q3 2026 Manheim 指數新聞稿", "2026-09 指數 205.9、年減 0.6%", "Verified"),
    ("ESP Data（引 Fitch）", "次級車貸 ABS 60 天以上逾期率 2026-01 達 6.90%（舊方法）、2026-07 6.13%（新方法）", "Analogy"),
    ("YinchenChang/company-models：oracle v0.4、docs/plan/個股模型原則_命題加營收分部.md；crwv-model v4.5", "模型架構、估值規則、無風險利率輸入", "—"),
]
for i, (a, b, c) in enumerate(srcs, 1):
    SO.cell(3 + i, 1, i); SO.cell(3 + i, 2, a); SO.cell(3 + i, 3, b); SO.cell(3 + i, 4, c)
SO.column_dimensions["B"].width = 80; SO.column_dimensions["C"].width = 70

# =====================================================================================
# README（放最前）
# =====================================================================================
RM = wb.create_sheet("README", 0)
RM["A1"] = f"Carvana（CVNA）個股模型 v{VERSION}（{BUILD_DATE}）— 命題＋營收線＋收支平衡＋信用風險"
RM["A2"] = "命題：Carvana 的高成長與獲利，毛利約三成、調整後 EBITDA 約一半來自「把車貸賣掉的利得」；只要消費信用與 ABS 市場正常，它能自我融資，一旦信用循環轉向，利得、放款與台數會同時被壓縮——問題是它的現金與債務結構撐不撐得住"
RM["A3"] = "核心問題：在信用循環轉差時，Carvana 能否不靠新股與新債維持收支平衡？債務／EBITDA 會不會越過評等門檻？"
RM["A4"] = "重點觀察（數字取自本活頁簿）"
B = P["B"]; CC = P["C"]
RM["A5"] = (f'="① 收支：情境 B 的 FY2027–FY2031 收支關卡合計 "&TEXT(CF!D4,"#,##0")&"（百萬美元），情境 C 為 "&TEXT(CF!E4,"#,##0")&"；情境 C 期間最低現金 "&TEXT(CF!E6,"#,##0")&"，新增外部融資需求 "&TEXT(CF!E5,"#,##0")&"。"')
RM["A6"] = (f'="② 債信：2026-09-30 TTM 債務／EBITDA "&TEXT(ttm_lev,"0.0")&"x，扣除貸款出售利得後為 "&TEXT(ttm_lev_xg,"0.0")&"x；情境 C 2027 年債務／EBITDA "&TEXT(Credit!{yc(2027)}{K["C"]["lev"]},"0.0")&"x（Moody\'s 升評門檻 4.0x）。"')
RM["A7"] = (f'="③ 貸款去化：已承諾承購（Ally＋整批出售）只涵蓋情境 B 2027 年發放量的 "&TEXT(Credit!{yc(2027)}{K["B"]["take"]},"0%")&"；若所有買方與 ABS 同時停止，現金與倉儲額度約可撐 "&TEXT(Credit!C{ROWMAP["Credit"]["S3"]["runway"]},"0.0")&" 個月（Ally 照常則 "&TEXT(Credit!D{ROWMAP["Credit"]["S3"]["runway"]},"0.0")&" 個月）。"')
RM["A8"] = (f'="④ 消費信用：管理組合年化淨沖銷率 2026 上半年 "&TEXT(Credit!I{ROWMAP["Credit"]["S1"]["ncor"]},"0.00%")&"（非 Prime 推得 "&TEXT(Credit!I{ROWMAP["Credit"]["S1"]["ncorn"]},"0.0%")&"），以一年前餘額計為 "&TEXT(Credit!I{ROWMAP["Credit"]["S1"]["lag"]},"0.00%")&"——改善有一部分來自組合快速成長的分母稀釋。"')
RM["A9"] = (f'="⑤ 估值：加權目標價 "&TEXT(v_tp_w,"0.0")&" 美元，相對現價 "&TEXT(v_up,"+0%;-0%")&"（"&Val!C{VR["rec"]}&"）；市場隱含 2030 年須有約 "&TEXT(Val!C{VR["req_u"]}/10000,"0")&" 萬台零售量（以情境 B 每台 EBITDA 換算；情境 B 為 "&TEXT(Thesis!{yc(2030)}{R["B"]["u"]}/10000,"0")&" 萬台）。三情境 IRR 機率加權 "&TEXT(v_irr_w,"0.0%")&"，低於 WACC "&TEXT(v_wacc,"0.0%")&"——結論取決於折現率。"')
RM["A11"] = "工作表：README → In（輸入）→ Q（事實總帳）→ Thesis（命題：台數 × 單位經濟）→ PL（營收線 → 損益）→ CF（收支模組、債務表）→ Credit（信用風險七節）→ Val（DCF＋倍數）→ Recon（對帳）→ Sens（敏感度）→ Sources"
RM["A12"] = "會計年度＝曆年；金額百萬美元；每台金額美元。FY2026E＝前三季實績＋第四季模型（v0.2 起）；公司自 Q4 2026 起不再提供季度與年度指引，2027 年起改以 GAAP 指標（中期目標：300 萬台、營業利益率 12.5%，2030–2035）"
RM["A15"] = "v0.1 → v0.2（2026-10-29）：納入 Q3 2026 實績（台數 217,385，高於共識 3.4%；零售 GPU 3,556 高於共識、批發 GPU 708 低於共識、其他 GPU 2,733 持平；SG&A 805 偏高）；2026 年改為前三季實績＋Q4 模型；債務表改為實際（定期貸款 B 1,660 以利率交換固定 4.055% 至 2028-07、2030 年票據已以溢價 50 贖回、車隊融資 142）；SG&A 變動成本 2,700 → 2,900；出售貸款存續期間 1.8 → 2.0（貸款池模擬）；現價改為 2026-10-28 收盤（推得）59.03；TTM 改為 Q4 2025–Q3 2026"
RM["A13"] = "營收分部層：Carvana 只有一個應報導部門（10-Q 附註 2）；營收列採公司固定揭露的三條營收線（零售、批發、其他），其他營收下拆「貸款出售利得」與「非貸款」兩個命題子列，PL 頁核對列為 0"
RM["A14"] = "與專案既有模型的關係：比照 company-models 的 Oracle v0.4（命題＋分部架構、五級評等、DCF 50%＋倍數 50%、rf 5.00%）與 CoreWeave 收支模型「Backlog 不是現金」的收支精神；本模型不讀 Tokenomics（Carvana 與算力無關）"
RM["A16"] = "已知限制"
RM["A17"] = ("1. 非 Prime（144A）證券化的靜態池損失未公開，非 Prime 淨沖銷率以「全部 − Prime」推得。2. 利息以期初餘額計，忽略年內再融資時點（再融資年度另以月數拆）。"
             "3. TRA 以平均攤付處理，未隨課稅所得與未來交換變動。4. 產能建置成本、維護資本支出、利得率對損失與利差的傳導係數皆為 Assumed。"
             "5. 估值折現基準點為 2026 年底，未計 2026 年第四季現金流。6. 關係人與放空報告主張無法以公開資料驗證，只列對照。7. 2026 年評等動作未找到，門檻沿用 2025 年公開說法。")
for c in ("A5", "A6", "A7", "A8", "A9", "A17", "A2", "A3"):
    RM[c].alignment = Alignment(wrap_text=True, vertical="top")
RM.column_dimensions["A"].width = 160

out = sys.argv[1] if len(sys.argv) > 1 else "/tmp/carvana_raw.xlsx"
wb.save(out)
if os.environ.get("CVNA_ROWMAP"):
    json.dump({"R": R, "P": P, "C": C, "ROWMAP": {k: v for k, v in ROWMAP.items()}, "CHK": CHK, "TBLK": TBLK}, open(os.environ["CVNA_ROWMAP"], "w"), ensure_ascii=False, indent=1, default=str)
print("saved", out)
