"""Generates the 5 deliverables (xlsx x2, docx x2, pptx) from data/*.csv. Run: .venv/bin/python src/build_docs.py"""
import os, numpy as np, pandas as pd
from scipy import stats
from statsmodels.stats.proportion import confint_proportions_2indep, proportions_ztest
from statsmodels.stats.power import NormalIndPower
from statsmodels.stats.proportion import proportion_effectsize
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Font, PatternFill, Alignment
from docx import Document
from docx.shared import Pt
from pptx import Presentation
from pptx.util import Inches, Pt as PPt

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__))); os.chdir(ROOT)
OUT = "outputs"; N = "Moiseyenko"
cols = ["campaign","date","spend","impressions","reach","clicks","searches","view_content","add_to_cart","purchase"]
c = pd.read_csv("data/control_group.csv", sep=";", header=0, names=cols)
t = pd.read_csv("data/test_group.csv", sep=";", header=0, names=cols)
for d in (c, t): d["date"] = pd.to_datetime(d["date"], format="%d.%m.%Y")
raw = pd.concat([c, t], ignore_index=True)
raw["group"] = np.where(raw.campaign.str.startswith("Control"), "Control", "Test")
raw["excluded"] = raw[cols[3:]].isna().any(axis=1)           # 5.08 control: only Spend filled
c = c.dropna().reset_index(drop=True)                          # option A: drop only the broken control day
m = ["spend","impressions","reach","clicks","searches","view_content","add_to_cart","purchase"]

# ---------- numbers for text ----------
rows = []
for k in m:
    r = stats.ttest_ind(t[k], c[k], equal_var=False)
    rows.append((k, c[k].mean(), t[k].mean(), (t[k].mean()/c[k].mean()-1)*100, r.statistic, r.pvalue, stats.mannwhitneyu(t[k], c[k]).pvalue))
T = pd.DataFrame(rows, columns=["metric","ctrl","test","diff","t","p","p_mw"]).set_index("metric")
tot = {g: d[m].sum() for g, d in (("c", c), ("t", t))}
nd = {"c": len(c), "t": len(t)}
def pooled(g):
    s = tot[g]; return dict(ctr=s.clicks/s.impressions, cpm=s.spend/s.impressions*1000, cpc=s.spend/s.clicks, cpa=s.spend/s.purchase,
        s2=s.searches/s.clicks, s3=s.view_content/s.searches, s4=s.add_to_cart/s.view_content, s5=s.purchase/s.add_to_cart, cr=s.purchase/s.clicks)
pc_, pt_ = pooled("c"), pooled("t")
# daily CPA / step comparison
for d in (c, t): d["cpa"] = d.spend/d.purchase; d["vc_atc"] = d.add_to_cart/d.view_content; d["atc_pur"] = d.purchase/d.add_to_cart
cpa_p = stats.ttest_ind(t.cpa, c.cpa, equal_var=False).pvalue
vcatc_p = stats.ttest_ind(t.vc_atc, c.vc_atc, equal_var=False).pvalue
atcpur_p = stats.ttest_ind(t.atc_pur, c.atc_pur, equal_var=False).pvalue
for d in (c, t):
    d["ctr"] = d.clicks/d.impressions; d["cpc"] = d.spend/d.clicks; d["cvr"] = d.purchase/d.clicks; d["atc_rate"] = d.add_to_cart/d.clicks
    d["s_rate"] = d.searches/d.clicks; d["vc_s"] = d.view_content/d.searches
RM = [("Конверсія клік→покупка (Purchase/Clicks)","cvr",4),("CTR (Clicks/Impressions)","ctr",4),("CPC, $ (Spend/Clicks)","cpc",2),
      ("Add-to-Cart Rate (AddToCart/Clicks)","atc_rate",4),("Клік→пошук","s_rate",4),("Пошук→перегляд контенту","vc_s",4),
      ("Перегляд→кошик","vc_atc",4),("Кошик→покупка","atc_pur",4),("CPA, $ (Spend/Purchase)","cpa",2)]
RR = []
for nm, k, dg in RM:
    r = stats.ttest_ind(t[k], c[k], equal_var=False); RR.append((nm, c[k].mean(), t[k].mean(), r.statistic, r.pvalue, stats.mannwhitneyu(t[k], c[k]).pvalue, dg))
shap_t = stats.shapiro(t.purchase).pvalue
# CI task 5
xc, nc, xt, nt = tot["c"].purchase, tot["c"].add_to_cart, tot["t"].purchase, tot["t"].add_to_cart
pcv, ptv = xc/nc, xt/nt; diff = ptv-pcv; se = np.sqrt(pcv*(1-pcv)/nc + ptv*(1-ptv)/nt); z = stats.norm.ppf(.975)
lo, hi = diff-z*se, diff+z*se
base = tot["c"].purchase/tot["c"].clicks
ss = {r: int(np.ceil(NormalIndPower().solve_power(proportion_effectsize(base*(1+r), base), alpha=.05, power=.8))) for r in (.05,.10,.15)}
pct = lambda x: f"{x*100:.1f}%"
print(T.round(4)); print(pc_, pt_); print(diff, lo, hi, cpa_p, vcatc_p, atcpur_p, ss, shap_t)

HDR = Font(bold=True, color="FFFFFF"); FILL = PatternFill("solid", fgColor="305496")
def hdr(ws, row, ncols):
    for i in range(1, ncols+1):
        c_ = ws.cell(row, i); c_.font = HDR; c_.fill = FILL; c_.alignment = Alignment(wrap_text=True, horizontal="center")


from openpyxl.styles import Border, Side
import math
_th = Side(style="thin", color="B4C6E7"); BORD = Border(left=_th, right=_th, top=_th, bottom=_th)
NAVY, MID, LIGHT, KEY = "1F3864", "2F5597", "EAF1FB", "FFF2CC"
def fillc(c): return PatternFill("solid", fgColor=c)
def banner(ws, row, c1, c2, text, size=14, color=NAVY, height=28):
    ws.merge_cells(start_row=row, start_column=c1, end_row=row, end_column=c2)
    c_ = ws.cell(row, c1, text); c_.font = Font(bold=True, size=size, color="FFFFFF"); c_.fill = fillc(color)
    c_.alignment = Alignment(vertical="center", wrap_text=True, indent=1); ws.row_dimensions[row].height = height
    for cc in range(c1, c2+1): ws.cell(row, cc).fill = fillc(color)
def body(ws, r1, r2, c1, c2, wrap_cols=(), widths=None, key_rows=()):
    for r in range(r1, r2+1):
        lines = 1
        for cc in range(c1, c2+1):
            c_ = ws.cell(r, cc); c_.border = BORD; c_.fill = fillc(KEY if r in key_rows else LIGHT)
            c_.alignment = Alignment(wrap_text=cc in wrap_cols or c_.alignment.wrap_text, vertical="top" if cc in wrap_cols else "center")
            if cc in wrap_cols and c_.value and widths:
                lines = max(lines, math.ceil(len(str(c_.value))/(widths[cc-1]*1.05)))
        if lines > 1: ws.row_dimensions[r].height = 15*lines+4
        if r in key_rows:
            for cc in range(c1, c2+1): ws.cell(r, cc).font = Font(bold=True)

import re
def ua_numbers(t_):
    """єдиний формат чисел у тексті: крапка як десятковий роздільник, без роздільника тисяч"""
    if not isinstance(t_, str): return t_
    t_ = re.sub(r"(?<![\d,.])0,(\d)", r"0.\1", t_)                                   # 0,05 -> 0.05
    return re.sub(r"(?<![\d,.])([1-9]\d{0,2})((?:[, ]\d{3})+)(?!\d)",
                  lambda m: m.group(1) + re.sub(r"[, ]", "", m.group(2)), t_)    # 15,044 / 15 044 -> 15044

def data_sheet(wb):
    ws = wb.active; ws.title = "Data"
    names = ["Group","Campaign","Date","Spend [USD]","Impressions","Reach","Website Clicks","Searches","View Content","Add to Cart","Purchase","Excluded (missing data)"]
    ws.append(names); hdr(ws, 1, len(names))
    for _, r in raw.iterrows():
        v = [None if pd.isna(r[k]) else float(r[k]) for k in cols[2:]]
        ws.append([r.group, r.campaign, r.date.to_pydatetime()] + v + ["YES" if r.excluded else ""])
    for row in ws.iter_rows(min_row=2, min_col=3, max_col=3):
        row[0].number_format = "dd.mm.yyyy"
    for i, w in enumerate([10,18,12,12,13,12,14,11,13,12,10,22], 1): ws.column_dimensions[chr(64+i)].width = w
    ws.freeze_panes = "A2"
    return ws, len(raw)+1

# ---------- Task 4 xlsx ----------
wb = Workbook(); ws, last = data_sheet(wb)
f = wb.create_sheet("Funnel")
f["A1"] = "Воронка Control vs Test (дані: Data; день 5.08 у контролі з порожніми метриками не враховується в SUMIFS/AVERAGEIFS)"; f["A1"].font = Font(bold=True, size=12)
f.append([]); f.append(["Крок", "Control, сума", "Test, сума", "Control, в середньому за день", "Test, в середньому за день", "Δ середнього за день, %"]); hdr(f, 3, 6)
stages = [("Покази","E"),("Кліки","G"),("Пошук","H"),("Перегляд контенту","I"),("Додавання в кошик","J"),("Покупка","K")]
D = lambda col: f"Data!${col}$2:${col}${last}"
for i, (nm, col) in enumerate(stages):
    r = 4+i
    f.cell(r,1,nm)
    f.cell(r,2,f'=SUMIFS({D(col)},{D("A")},"Control")'); f.cell(r,3,f'=SUMIFS({D(col)},{D("A")},"Test")')
    f.cell(r,4,f'=AVERAGEIFS({D(col)},{D("A")},"Control")'); f.cell(r,5,f'=AVERAGEIFS({D(col)},{D("A")},"Test")')
    f.cell(r,6,f"=E{r}/D{r}-1"); f.cell(r,6).number_format = "0.0%"
    for cc in (2,3,4,5): f.cell(r,cc).number_format = "#,##0"
f["A11"] = "Кількість днів у даних"; f["B11"] = f'=COUNTIFS({D("A")},"Control",{D("L")},"<>YES")'; f["C11"] = f'=COUNTIFS({D("A")},"Test",{D("L")},"<>YES")'
f["A12"] = "Spend, сума ($)"; f["B12"] = f'=SUMIFS({D("D")},{D("A")},"Control",{D("L")},"<>YES")'; f["C12"] = f'=SUMIFS({D("D")},{D("A")},"Test",{D("L")},"<>YES")'
f["A13"] = "CPA = Spend / Purchase ($)"; f["B13"] = "=B12/B9"; f["C13"] = "=C12/C9"
f["A14"] = "CPM ($)"; f["B14"] = "=B12/B4*1000"; f["C14"] = "=C12/C4*1000"
f["A15"] = "CTR (кліки / покази)"; f["B15"] = "=B5/B4"; f["C15"] = "=C5/C4"
for r in (12,13,14): 
    for cc in "BC": f[f"{cc}{r}"].number_format = "#,##0.00"
for cc in "BC": f[f"{cc}15"].number_format = "0.00%"
f["A17"] = "Конверсія між кроками"; f["A17"].font = Font(bold=True)
f.append([]); f.cell(18,1,"Перехід"); f.cell(18,2,"Control"); f.cell(18,3,"Test"); f.cell(18,4,"Δ, п.п."); f.cell(18,5,"% від кліків: Control"); f.cell(18,6,"% від кліків: Test"); hdr(f, 18, 6)
for i in range(5):
    r = 19+i; a, b = 4+i, 5+i
    f.cell(r,1,f"{stages[i][0]} → {stages[i+1][0]}")
    f.cell(r,2,f"=B{b}/B{a}"); f.cell(r,3,f"=C{b}/C{a}"); f.cell(r,4,f"=(C{r}-B{r})*100")
    for cc in (2,3): f.cell(r,cc).number_format = "0.0%"
    f.cell(r,4).number_format = "+0.0;-0.0"
    f.cell(r,5,f"=B{b}/B$5"); f.cell(r,6,f"=C{b}/C$5")
    for cc in (5,6): f.cell(r,cc).number_format = "0.0%"
f.column_dimensions["A"].width = 42
for cc in "BCDEF": f.column_dimensions[cc].width = 19
ch = BarChart(); ch.type = "bar"; ch.title = "Воронка: середнє за день (від кліків)"; ch.style = 10
ch.add_data(Reference(f, min_col=4, max_col=5, min_row=3, max_row=9), titles_from_data=True); ch.set_categories(Reference(f, min_col=1, min_row=5, max_row=9))
# rows 5..9 only (clicks..purchase); rebuild refs without impressions
ch = BarChart(); ch.type = "bar"; ch.title = "Воронка: у середньому за день (кліки → покупка)"
ch.add_data(Reference(f, min_col=4, max_col=5, min_row=5, max_row=9), titles_from_data=False)
ch.set_categories(Reference(f, min_col=1, min_row=5, max_row=9))
from openpyxl.chart.series import SeriesLabel
ch.series[0].tx = SeriesLabel(v="Control"); ch.series[1].tx = SeriesLabel(v="Test")
ch.y_axis.scaling.orientation = "minMax"; ch.x_axis.scaling.orientation = "maxMin"; ch.height, ch.width = 9, 18
ch.x_axis.delete = False; ch.y_axis.delete = False
f.add_chart(ch, "H3")
ch2 = BarChart(); ch2.type = "col"; ch2.title = "Конверсія між кроками воронки"
ch2.add_data(Reference(f, min_col=2, max_col=3, min_row=19, max_row=23), titles_from_data=False)
ch2.set_categories(Reference(f, min_col=1, min_row=19, max_row=23))
ch2.series[0].tx = SeriesLabel(v="Control"); ch2.series[1].tx = SeriesLabel(v="Test"); ch2.height, ch2.width = 9, 18
ch2.y_axis.number_format = "0%"; ch2.x_axis.delete = False; ch2.y_axis.delete = False
f.add_chart(ch2, "H22")
ins = wb.create_sheet("Висновки")
lines = [
 "ВИСНОВКИ ПО ВОРОНЦІ",
 f"1. Верх воронки: у тесті на {-T.loc['impressions','diff']:.0f}% менше показів на день і на {-T.loc['reach','diff']:.0f}% менший охват, при +{T.loc['spend','diff']:.0f}% витрат. CPM ${pt_['cpm']:.1f} vs ${pc_['cpm']:.1f}, але CTR {pct(pt_['ctr'])} vs {pct(pc_['ctr'])}. Імовірне пояснення (гіпотеза, у даних перевірити неможливо): реклама показується вужчій, активнішій аудиторії.",
 f"2. Кліків на день більше на {T.loc['clicks','diff']:.0f}%, але різниця не значима (p={T.loc['clicks','p']:.2f}).",
 f"3. Середина: клік→пошук {pct(pt_['s2'])} vs {pct(pc_['s2'])}; пошук→перегляд контенту {pct(pt_['s3'])} vs {pct(pc_['s3'])} (Test vs Control).",
 f"4. Головна просадка — перегляд→кошик: {pct(pt_['s4'])} vs {pct(pc_['s4'])}. Додавань у кошик на день на {-T.loc['add_to_cart','diff']:.0f}% менше (p={T.loc['add_to_cart','p']:.4f}).",
 f"5. Кошик→покупка: {pct(pt_['s5'])} vs {pct(pc_['s5'])}. Хто дійшов до кошика, купує значно частіше.",
 f"6. Покупок на день порівну ({T.loc['purchase','test']:.0f} vs {T.loc['purchase','ctrl']:.0f}, p={T.loc['purchase','p']:.2f}), витрати вищі → CPA ${pt_['cpa']:.2f} vs ${pc_['cpa']:.2f} ({(pt_['cpa']/pc_['cpa']-1)*100:+.0f}% у сумі); по днях різниця CPA не значима (p={cpa_p:.2f}).",
 "7. Висновок: нова стратегія не збільшила кількість покупок, а покупка обійшлась дорожче (+12% у сумі, хоча по днях це не значимо). Кроки воронки рухаються в різні боки: перегляд→кошик падає, кошик→покупка зростає.",
 "",
 "Примітка про дані: 5.08.2019 у контролі є лише Spend ($1 835), решта метрик порожні — день виключено з розрахунків (прапорець в колонці L листа Data). У тесті всі 30 днів збережено, тому суми порівнюються і як сума, і як середнє за день.",
]
for i, l in enumerate(lines, 1): ins.cell(i,1,ua_numbers(l)).alignment = Alignment(wrap_text=True, vertical="top")
ins["A1"].font = Font(bold=True, size=12); ins.column_dimensions["A"].width = 130
f["A1"].value = None; banner(f, 1, 1, 6, "Воронка Control vs Test — суми та середнє за день (дані: лист Data; 5.08 у контролі виключено)", size=12, height=32)
body(f, 4, 9, 1, 6); body(f, 11, 15, 1, 3); body(f, 19, 23, 1, 6)
f["A17"].font = Font(bold=True, color=MID)
ins["A1"].value = None; banner(ins, 1, 1, 1, "Висновки по воронці", size=14)
for r in range(2, len(lines)+1):
    ins.cell(r,1).alignment = Alignment(wrap_text=True, vertical="top")
    if ins.cell(r,1).value:
        body(ins, r, r, 1, 1, wrap_cols=(1,), widths=[130])
wb.save(f"{OUT}/{N}_Final Project-4.xlsx")

# ---------- Task 5 xlsx ----------
wb = Workbook(); ws, last = data_sheet(wb)
s = wb.create_sheet("CI", 0)
s["A1"] = "95% довірчий інтервал різниці конверсії кроку Add to Cart → Purchase (Test − Control)"; s["A1"].font = Font(bold=True, size=12)
items = [
 ("Control: Add to Cart (n)", f'=SUMIFS({D("J")},{D("A")},"Control")', "#,##0"),
 ("Control: Purchase (x)", f'=SUMIFS({D("K")},{D("A")},"Control")', "#,##0"),
 ("Control: конверсія p1 = x/n", "=B4/B3", "0.00%"),
 ("Test: Add to Cart (n)", f'=SUMIFS({D("J")},{D("A")},"Test")', "#,##0"),
 ("Test: Purchase (x)", f'=SUMIFS({D("K")},{D("A")},"Test")', "#,##0"),
 ("Test: конверсія p2 = x/n", "=B7/B6", "0.00%"),
 ("Різниця d = p2 − p1", "=B8-B5", "0.00%"),
 ("Стандартна похибка SE = √(p1(1−p1)/n1 + p2(1−p2)/n2)", "=SQRT(B5*(1-B5)/B3+B8*(1-B8)/B6)", "0.0000%"),
 ("Рівень довіри", 0.95, "0%"),
 ("z-критичне = NORM.S.INV(1−(1−CL)/2)", "=NORM.S.INV(1-(1-B11)/2)", "0.000"),
 ("Похибка (margin) = z·SE", "=B12*B10", "0.00%"),
 ("Нижня межа ДІ", "=B9-B13", "0.00%"),
 ("Верхня межа ДІ", "=B9+B13", "0.00%"),
 ("z-статистика (d / SE під H0, pooled)", "=B9/SQRT(((B4+B7)/(B3+B6))*(1-(B4+B7)/(B3+B6))*(1/B3+1/B6))", "0.00"),
 ("p-value (двосторонній)", '=IF(2*NORM.S.DIST(-ABS(B16),TRUE)<0.0001,"< 0.0001",2*NORM.S.DIST(-ABS(B16),TRUE))', "0.0000"),
]
s.append([]); s.cell(2,1,"Показник"); s.cell(2,2,"Значення"); hdr(s, 2, 2)
for i, (a, b, nf) in enumerate(items):
    s.cell(3+i,1,a); s.cell(3+i,2,b).number_format = nf
for r in (14,15): s.cell(r,1).font = Font(bold=True); s.cell(r,2).font = Font(bold=True)
s["A20"] = "Інтерпретація"; s["A20"].font = Font(bold=True)
IL = [
 f"Конверсія кошик→покупка: Control — {pcv*100:.2f}%, Test — {ptv*100:.2f}%.",
 f"Різниця (Test − Control) = {diff*100:.2f} п.п.",
 f"95% довірчий інтервал: [{lo*100:.2f}; {hi*100:.2f}] п.п.",
 "Обидві межі додатні — нуль в інтервал не входить.",
 f"Висновок: різниця {diff*100:+.2f} п.п. статистично значуща (95% ДІ не включає 0, p < 0,05): після додавання в кошик у тестовій групі купують помітно частіше.",
 f"Масштаб: це +{(ptv/pcv-1)*100:.0f}% відносно контролю ({pcv*100:.1f}% → {ptv*100:.1f}%).",
 f"Контекст: у тесті в кошик потрапляє на {-T.loc['add_to_cart','diff']:.0f}% менше людей на день, тому покупок на день порівну ({T.loc['purchase','test']:.0f} проти {T.loc['purchase','ctrl']:.0f}). Висока конверсія кроку не означає перемоги тесту: вона лише компенсує втрати на попередньому кроці: після перегляду товару в кошик додають значно рідше (47% проти 67%).",
 "Застереження: розрахунок вважає кошики незалежними, а дані — денні агрегати, тому реальна невизначеність більша.",
 "Менша кількість кошиків, імовірно, означає відсів менш налаштованих людей, тож частину росту конверсії може пояснювати склад аудиторії, а не якість оплати.",
]
for i, l in enumerate(IL):
    r = 21+i; s.cell(r,1,ua_numbers(l)); s.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
    s.cell(r,1).alignment = Alignment(wrap_text=True, vertical="center", indent=1); s.row_dimensions[r].height = 22 if len(l) < 85 else (38 if len(l) < 170 else 56)
s.column_dimensions["A"].width = 58; s.column_dimensions["B"].width = 16
s["A1"].value = None; banner(s, 1, 1, 4, "95% довірчий інтервал різниці конверсії кроку Add to Cart → Purchase (Test − Control)", size=12, height=32)
body(s, 3, 17, 1, 2, key_rows=(9, 14, 15))
for r in range(21, 21+len(IL)):
    for cc in range(1, 5): s.cell(r, cc).fill = fillc(KEY if r == 25 else LIGHT); s.cell(r, cc).border = BORD
s.cell(25,1).font = Font(bold=True)
s["A20"].font = Font(bold=True, color=MID)
s["B11"].font = Font(bold=True, color="0000FF")
wb.save(f"{OUT}/{N}_Final Project-5.xlsx")

# ---------- docx helpers ----------
def new_doc(title):
    d = Document(); d.styles["Normal"].font.name = "Calibri"; d.styles["Normal"].font.size = Pt(11); d.add_heading(title, 0); return d
def table(d, header, rows, style="Light Grid Accent 1"):
    tb = d.add_table(rows=1, cols=len(header)); tb.style = style
    for i, h in enumerate(header): tb.rows[0].cells[i].text = h
    for r in rows:
        cs = tb.add_row().cells
        for i, v in enumerate(r): cs[i].text = str(v)
    d.add_paragraph()
def bullets(d, items):
    for i in items: d.add_paragraph(i, style="List Bullet")

# ---------- Task 1 xlsx ----------
wb = Workbook(); w = wb.active; w.title = "Тест-план"
w["A1"] = "Тест-план: A/B-тест нового підходу до онлайн-реклами"; w["A1"].font = Font(bold=True, size=14)
bc, bt = raw[raw.group == "Control"].spend.sum(), raw[raw.group == "Test"].spend.sum()
n = lambda x: f"{x:,.0f}".replace(",", " ")
plan = [
 ("1. Мета тесту", None),
 ("Гіпотеза H0", "Новий підхід не змінює конверсію клік→покупка та CPA порівняно з поточним."),
 ("Гіпотеза H1", "Новий підхід змінює (очікуємо: покращує) конверсію клік→покупка та/або знижує CPA. Двосторонній тест, α = 0,05, потужність 80%."),
 ("Що хочемо з'ясувати", "Чи дає нове розміщення більше покупок на кліки й на витрачений долар, і чи не погіршує воно інші етапи воронки."),
 ("2. Метрики", None),
 ("Основна метрика (North Star)", "Конверсія клік→покупка = Purchase / Website Clicks. Додатково для бізнес-рішення — CPA = Spend / Purchase."),
 ("Вторинні метрики", "CTR (кліки/покази), CPC (витрати/кліки), конверсії кроків воронки: клік→пошук, пошук→перегляд, перегляд→кошик, кошик→покупка."),
 ("Захисні метрики (guardrails)", "Spend, Reach, Impressions, CPM — щоб не допустити неконтрольованого росту витрат і падіння охвату."),
 ("3. Дизайн тесту", None),
 ("Групи", "Control Campaign — поточний підхід. Test Campaign — новий підхід."),
 ("Розподіл", "Випадковий поділ 50/50 на рівні користувача (cookie / user_id); кожен користувач бачить лише одну версію. Групи працюють одночасно."),
 ("Бюджет у наших даних", f"Control: ${n(bc)} за серпень (разом із $1 835 за 5.08, коли метрики воронки порожні); Test: ${n(bt)} (+{(bt/bc-1)*100:.0f}% у сумі за місяць; на день +{T.loc['spend','diff']:.0f}%, бо в контролі 29 днів із даними). Різницю треба враховувати: порівнюємо відношення (конверсії, CPA), а не абсолютні суми."),
 ("Умови", "Однакові аудиторії та креативи, окрім змінної тесту; налаштування кампаній не змінюємо під час тесту."),
 ("4. Тривалість", None),
 ("Період", "Мінімум 2 повні тижні, щоб покрити всі дні тижня; у нашому експерименті — 1–30 серпня 2019 (30 днів, по 4+ рази кожен день тижня)."),
 ("Дострокова зупинка", "Лише за суттєвого збитку (наприклад, конверсія впала вдвічі). Інакше не зупиняємо за проміжним p-value — це завищує хибнопозитивні результати."),
 ("5. Розмір вибірки", None),
 ("Розрахунок", f"Базова конверсія клік→покупка у контролі {base*100:.2f}%. Щоб помітити відносний ріст на 10% (α=0,05, потужність 80%), потрібно ≈15 тис. кліків на групу — див. лист «Розмір вибірки» (формула, параметри змінюються)."),
 ("Наші дані", f"Control (29 днів, без 5.08): {n(tot['c'].impressions)} показів, {n(tot['c'].clicks)} кліків, {n(tot['c'].purchase)} покупок. Test (30 днів): {n(tot['t'].impressions)} показів, {n(tot['t'].clicks)} кліків, {n(tot['t'].purchase)} покупок. Кліків приблизно в 10 разів більше за необхідний мінімум, тож вибірки достатньо для виявлення навіть невеликого ефекту."),
 ("6. Критерії успіху", None),
 ("Статистичні методи", "Конверсії (пропорції): z-тест для двох пропорцій і довірчий інтервал різниці. Денні метрики: t-тест Велча + Манна-Уїтні. Основна метрика одна; решта пояснювальні, при багатьох порівняннях — поправка Бонферроні/Холма."),
 ("Тест успішний, якщо", "p < 0,05 і конверсія клік→покупка (або CPA) у Test значимо краща, а guardrails не погіршились."),
 ("Практична значущість", "Навіть значима різниця має бути економічно виправданою: покращення має перекривати додаткові витрати. Якщо ефект статистично значимий, але дрібний або CPA гірший — не впроваджуємо."),
 ("Неоднозначний результат", "Якщо кроки воронки рухаються в різні боки — продовжуємо тест або повторюємо його після виправлень."),
 ("7. Ризики та обмеження", None),
 ("Незіставність груп", f"У даних витрати на день у тесті на {T.loc['spend','diff']:.0f}% вищі, а покази на {-T.loc['impressions','diff']:.0f}% і охват на {-T.loc['reach','diff']:.0f}% нижчі, ніж у контролі: розподіл не 50/50. Потрібна перевірка SRM (chi-square) і з'ясування причини."),
 ("Якість даних", "Пропуски в даних (5.08 у контролі: заповнено лише Spend, метрики воронки порожні) — щоденна перевірка повноти та правило, як виключати такі дні."),
 ("Ефект новизни", "Нова реклама може давати тимчасовий сплеск CTR лише через новизну; перші 2–3 дні інтерпретуємо обережно."),
 ("Сезонність і зовнішні фактори", "Один місяць (серпень) — результати можуть не перенестися на інші періоди; немає даних про конкуренцію й зміни на платформі."),
 ("Агреговані дані", "Денні агрегати не дають аналізу на рівні користувача й автокореляції днів — потрібен трекінг подій на рівні користувача (подій, атрибутів, груп)."),
]
w.append([]); w.append(["Розділ", "Опис"]); hdr(w, 3, 2)
SEC = PatternFill("solid", fgColor="D9E1F2")
for k, v in plan:
    w.append([k, ua_numbers(v)]); r = w.max_row
    if v is None:
        for cc in (1, 2): w.cell(r, cc).fill = SEC
        w.cell(r,1).font = Font(bold=True, size=12)
    else:
        w.cell(r,1).font = Font(bold=True); w.cell(r,2).alignment = Alignment(wrap_text=True, vertical="top"); w.cell(r,1).alignment = Alignment(vertical="top", wrap_text=True)
w.column_dimensions["A"].width = 32; w.column_dimensions["B"].width = 100
w.unmerge_cells if False else None
banner(w, 1, 1, 2, "Тест-план: A/B-тест нового підходу до онлайн-реклами")
for r in range(4, w.max_row+1):
    if w.cell(r,2).value is None:
        for cc in (1,2): w.cell(r,cc).fill = fillc(MID); w.cell(r,cc).font = Font(bold=True, size=12, color="FFFFFF"); w.cell(r,cc).border = BORD
        w.row_dimensions[r].height = 22
    else:
        body(w, r, r, 1, 2, wrap_cols=(1,2), widths=[32,100])
        w.cell(r,1).font = Font(bold=True)
z_ = wb.create_sheet("Розмір вибірки")
z_["A1"] = "Розрахунок розміру вибірки (кліків на групу) для порівняння двох пропорцій"; z_["A1"].font = Font(bold=True, size=12)
z_.append([]); z_.append(["Параметр", "Значення"]); hdr(z_, 3, 2)
rows_ = [("Базова конверсія клік→покупка p1 (контроль)", float(base), "0.00%"),
 ("MDE, відносний ефект", 0.10, "0%"), ("p2 = p1·(1+MDE)", "=B4*(1+B5)", "0.00%"),
 ("p̄ = (p1 + p2) / 2  (середня очікувана конверсія двох груп)", "=(B4+B6)/2", "0.00%"),
 ("α (двосторонній)", 0.05, "0.00"), ("Потужність (1−β)", 0.80, "0%"),
 ("z(α/2)", "=NORM.S.INV(1-B8/2)", "0.000"), ("z(β)", "=NORM.S.INV(B9)", "0.000"),
 ("n на групу = ( z(α/2)·√(2·p̄·(1−p̄)) + z(β)·√(p1·(1−p1) + p2·(1−p2)) )² / (p2 − p1)²",
  "=ROUNDUP((B10*SQRT(2*B7*(1-B7))+B11*SQRT(B4*(1-B4)+B6*(1-B6)))^2/(B6-B4)^2,0)", "#,##0"),
 ("Середньо кліків на день на групу (з даних)", float(c.clicks.mean()), "#,##0"), ("Днів для набору вибірки", "=ROUNDUP(B12/B13,0)", "0")]
for i, (a, b_, nf) in enumerate(rows_): z_.cell(4+i,1,a); z_.cell(4+i,2,b_).number_format = nf
z_["A16"] = "Чутливість: скільки кліків на групу потрібно для різного MDE (ті ж α і потужність)"; z_["A16"].font = Font(bold=True)
for i, m_ in enumerate((0.05, 0.10, 0.15)):
    r_ = 17+i; p2 = f"($B$4*(1+{m_}))"; pb = f"(($B$4+{p2})/2)"
    z_.cell(r_,1,f"MDE {int(m_*100)}%")
    z_.cell(r_,2,f"=ROUNDUP(($B$10*SQRT(2*{pb}*(1-{pb}))+$B$11*SQRT($B$4*(1-$B$4)+{p2}*(1-{p2})))^2/({p2}-$B$4)^2,0)").number_format = "#,##0"
z_.column_dimensions["A"].width = 78; z_.column_dimensions["B"].width = 16
z_["A1"].value = None; banner(z_, 1, 1, 2, "Розрахунок розміру вибірки (кліків на групу) для порівняння двох пропорцій", size=12)
body(z_, 4, 14, 1, 2, wrap_cols=(1,), widths=[78,16], key_rows=(12,14))
body(z_, 17, 19, 1, 2); z_["A16"].font = Font(bold=True, color=MID)
for r in (5,8,9): z_.cell(r,2).font = Font(bold=True, color="0000FF")   # вхідні параметри синім
z_["A21"] = "Сині значення — вхідні параметри, їх можна змінювати; решта рахується формулами."; z_["A21"].font = Font(italic=True, color="666666")
wb.save(f"{OUT}/{N}_Final Project-1.xlsx")

# ---------- Task 3 docx ----------
d = new_doc("Результати A/B-тесту та інтерпретація")
d.add_heading("Дані та метод", 1)
d.add_paragraph(f"Дані: щоденні агрегати Control ({len(c)} дн.) і Test ({len(t)} дн.), серпень 2019. День 5.08 у контролі виключено — там заповнено лише Spend, усі метрики воронки порожні. Для порівняння груп використано t-тест Велча (stats.ttest_ind, equal_var=False) для денних значень, перевірку непараметричним Манна-Уїтні, рівень значущості 0,05. "
                f"Тест Шапіро-Вілка для Purchase у тестовій групі дає p={shap_t:.2f} < 0,05, тобто розподіл відхиляється від нормального; тому поряд із t-тестом Велча наведено непараметричний тест Манна-Уїтні як підтвердження (висновки за обома збігаються).")
d.add_heading("Результати за метриками (середнє за день)", 1)
lab = {"spend":"Spend, $","impressions":"Покази","reach":"Охват","clicks":"Кліки","searches":"Пошуки","view_content":"Перегляд контенту","add_to_cart":"Додавання в кошик","purchase":"Покупки"}
table(d, ["Метрика","Control","Test","Δ, %","p (Welch)","p (Mann-Whitney)","Значимо?"],
      [[lab[k], f"{r.ctrl:,.0f}", f"{r.test:,.0f}", f"{r['diff']:+.1f}", f"{r.p:.4f}", f"{r.p_mw:.4f}", "так" if r.p < .05 else "ні"] for k, r in T.iterrows()])
d.add_heading("Похідні показники", 1)
table(d, ["Показник","Control","Test"], [
 ["CTR", pct(pc_['ctr']), pct(pt_['ctr'])], ["CPM, $", f"{pc_['cpm']:.1f}", f"{pt_['cpm']:.1f}"], ["CPC, $", f"{pc_['cpc']:.2f}", f"{pt_['cpc']:.2f}"],
 ["CPA, $", f"{pc_['cpa']:.2f}", f"{pt_['cpa']:.2f}"], ["Перегляд→кошик", pct(pc_['s4']), pct(pt_['s4'])], ["Кошик→покупка", pct(pc_['s5']), pct(pt_['s5'])]])
d.add_paragraph("Примітка: у таблиці «Похідні показники» значення пораховані по сумах за весь період (сума чисельників ÷ сума знаменників). У таблиці нижче — середнє денних відношень, тому значення можуть відрізнятися (наприклад, CPA $4.41 і $4.92 проти $5.05 і $5.90). Висновок за обома способами однаковий.")
d.add_heading("Щоденні відсоткові метрики (Welch t-test по днях)", 1)
d.add_paragraph("Кожен день рахуємо відношення, потім порівнюємо середні між групами (t = Test − Control, t-статистика Велча).")
table(d, ["Метрика","Control (mean)","Test (mean)","t-статистика","p-value (Welch)","p (Mann-Whitney)","Значимо?"],
      [[nm, f"{a:.{dg}f}", f"{b:.{dg}f}", f"{ts:.2f}", f"{pv:.4f}", f"{pm:.4f}", "так" if pv < .05 else "ні"] for nm, a, b, ts, pv, pm, dg in RR])
d.add_paragraph(f"Значимість по днях: CPA p={cpa_p:.2f}; перегляд→кошик p={vcatc_p:.4f}; кошик→покупка p={atcpur_p:.4f}.")
d.add_heading("Інтерпретація", 1)
bullets(d, [
 f"Ключовий результат: кількість покупок на день статистично не відрізняється (p={T.loc['purchase','p']:.2f}). Доказів, що новий підхід збільшує покупки, немає.",
 f"Тест коштує дорожче: витрати на день +{T.loc['spend','diff']:.0f}% (p={T.loc['spend','p']:.3f}). CPA ${pt_['cpa']:.2f} проти ${pc_['cpa']:.2f} ({(pt_['cpa']/pc_['cpa']-1)*100:+.0f}%) — тобто покупка стала дорожчою; по днях ця різниця не значима (p={cpa_p:.2f}) через великий розкид, але напрямок невигідний.",
 f"Охват і покази різко впали (покази {T.loc['impressions','diff']:.0f}%, охват {T.loc['reach','diff']:.0f}%, p<0,001), при цьому CTR виріс до {pct(pt_['ctr'])} з {pct(pc_['ctr'])} — імовірно, нова схема показує рекламу вужчій аудиторії з вищою залученістю (це припущення, у даних про аудиторію немає), але за вищою ціною за 1000 показів.",
 f"Додавання в кошик значимо менше ({T.loc['add_to_cart','diff']:.0f}%, p={T.loc['add_to_cart','p']:.4f}), але конверсія кошик→покупка значимо вища. Це схоже на зміну якості трафіку/цільової сторінки, а не на універсальне покращення.",
 "Різниці в кліках, пошуках і переглядах контенту не значимі.",
 f"За відсотковими метриками значимі CTR (p={RR[1][4]:.4f}), Add-to-Cart Rate (p={RR[3][4]:.4f}), перегляд→кошик (p={RR[6][4]:.4f}) та кошик→покупка (p={RR[7][4]:.4f}); конверсія клік→покупка (p={RR[0][4]:.2f}) і CPC (p={RR[2][4]:.2f}) — ні.",
 "Рекомендація: не масштабувати новий підхід на основі цього тесту. Варто з'ясувати причину падіння перегляд→кошик (лендінг, ціна, релевантність) і повторити тест з виправленням. Також перевірити повноту даних (5.08)."])
d.add_heading("Обмеження", 1)
bullets(d, ["Дані агреговані по днях, 29–30 спостережень на групу — потужність обмежена для метрик із великим розкидом (кліки, CPA).",
            "Денні спостереження можуть бути автокорельовані; користувачі в групах не відстежуються індивідуально.",
            "Багато порівнянь без поправки: основна метрика — кількість покупок/CPA, решта — пояснювальні."])
def _paras(doc):
    for p_ in doc.paragraphs: yield p_
    for tb_ in doc.tables:
        for row_ in tb_.rows:
            for cell_ in row_.cells:
                for p_ in cell_.paragraphs: yield p_
for p_ in _paras(d):
    for run_ in p_.runs:
        run_.text = ua_numbers(run_.text)
        if run_.text.strip() == "0.0000": run_.text = "< 0.0001"
d.save(f"{OUT}/{N}_Final Project-3.docx")

# ---------- Task 2 pptx ----------
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
rgb = lambda h: RGBColor.from_string(h)
C_NAVY, C_MID, C_LIGHT, C_ACC, C_TXT = "1F3864", "2F5597", "EAF1FB", "F2A900", "1F2937"
p = Presentation(); p.slide_width, p.slide_height = Inches(13.333), Inches(7.5)
BLANK = p.slide_layouts[6]
def rect(s, x, y, w, h, color, shape=MSO_SHAPE.RECTANGLE):
    r = s.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h)); r.fill.solid(); r.fill.fore_color.rgb = rgb(color); r.line.fill.background(); r.shadow.inherit = False; return r
def text(s, x, y, w, h, t, size=18, color=C_TXT, bold=False, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP):
    tb = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h)); tf = tb.text_frame; tf.word_wrap = True; tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Inches(0.1)
    para = tf.paragraphs[0]; para.text = t; para.alignment = align; f_ = para.font; f_.size = PPt(size); f_.bold = bold; f_.color.rgb = rgb(color); f_.name = "Calibri"; return tb
def base(title):
    s = p.slides.add_slide(BLANK); rect(s, 0, 0, 13.333, 1.1, C_NAVY); rect(s, 0, 1.1, 13.333, 0.07, C_ACC)
    text(s, 0.5, 0.15, 12.3, 0.8, title, size=28, color="FFFFFF", bold=True, anchor=MSO_ANCHOR.MIDDLE)
    text(s, 12.2, 7.0, 1.0, 0.35, str(len(p.slides)), size=11, color="7F8FA9", align=PP_ALIGN.RIGHT); return s
def slide(title, body=None, table_data=None, widths=None):
    s = base(title)
    if body:
        n_ = len(body); h = min(1.05, 5.5/n_ - 0.12); y = 1.5
        for l in body:
            rect(s, 0.5, y, 12.3, h, C_LIGHT); rect(s, 0.5, y, 0.1, h, C_MID)
            text(s, 0.75, y, 11.9, h, l, size=18, anchor=MSO_ANCHOR.MIDDLE); y += h + 0.12
    if table_data:
        rows, cols_ = len(table_data), len(table_data[0]); rh = min(0.62, 5.4/rows)
        gt = s.shapes.add_table(rows, cols_, Inches(0.5), Inches(1.5), Inches(12.3), Inches(rh*rows)); tbl = gt.table
        tblPr = gt._element.graphic.graphicData.tbl.tblPr
        for att in ("bandRow", "firstRow"): tblPr.set(att, "0")
        if widths:
            k = 12.3/sum(widths)
            for j, w_ in enumerate(widths): tbl.columns[j].width = Inches(w_*k)
        for i, row in enumerate(table_data):
            tbl.rows[i].height = Inches(rh)
            for j, v in enumerate(row):
                cell = tbl.cell(i, j); cell.text = v; cell.vertical_anchor = MSO_ANCHOR.MIDDLE
                cell.fill.solid(); cell.fill.fore_color.rgb = rgb(C_MID if i == 0 else (C_LIGHT if i % 2 == 0 else "FFFFFF"))
                cell.margin_left = cell.margin_right = Inches(0.1)
                for para in cell.text_frame.paragraphs:
                    para.font.size = PPt(14 if i == 0 else 12); para.font.bold = (i == 0 or j == 0); para.font.name = "Calibri"
                    para.font.color.rgb = rgb("FFFFFF" if i == 0 else C_TXT)
    return s
# title slide
s = p.slides.add_slide(BLANK); rect(s, 0, 0, 13.333, 7.5, C_NAVY); rect(s, 1.2, 3.55, 4.0, 0.08, C_ACC)
text(s, 1.1, 1.8, 11.0, 1.7, "Трекінг-план A/B-тесту онлайн-реклами", size=40, color="FFFFFF", bold=True, anchor=MSO_ANCHOR.BOTTOM)
text(s, 1.1, 3.8, 11.0, 1.0, "Події, атрибути та метрики порівняння Control vs Test", size=22, color="D6E0F2")
text(s, 1.1, 6.5, 11.0, 0.5, "Final Project — Завдання №2", size=14, color="9FB2D4")
# event flow
s = base("Шлях користувача: події воронки")
ev = ["ad_impression", "ad_click", "session_start", "site_search", "view_content", "add_to_cart", "begin_checkout", "purchase"]
w_ = 12.3/len(ev)
for i, e in enumerate(ev):
    sh = rect(s, 0.5 + i*w_, 2.6, w_ + 0.12, 1.2, C_MID if i < len(ev)-1 else C_ACC, MSO_SHAPE.CHEVRON if i else MSO_SHAPE.PENTAGON)
    lab = e.replace("_", "_\n", 1) if (e != "purchase" and e != "add_to_cart") else ("add_\nto_cart" if e == "add_to_cart" else e)
    text(s, 0.5 + i*w_ + (0.38 if i else 0.15), 2.6, w_ - 0.45, 1.2, lab, size=12, color="FFFFFF" if i < len(ev)-1 else C_NAVY, bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
text(s, 0.5, 4.2, 12.3, 0.5, "Кожна подія несе ab_group, user_id і session_id — це дозволяє порівнювати Control і Test на кожному кроці та по кожному користувачу.", size=16, color=C_TXT)
for i, (lab, d) in enumerate((("Верх воронки", "ad_impression → ad_click: охват, CTR, вартість кліку"), ("Середина", "session_start → view_content: якість трафіку та інтерес до продукту"), ("Низ воронки", "add_to_cart → purchase: намір і конверсія в оплату"))):
    x = 0.5 + i*4.15; rect(s, x, 5.0, 4.0, 1.6, C_LIGHT); rect(s, x, 5.0, 4.0, 0.08, C_MID)
    text(s, x+0.1, 5.15, 3.8, 0.4, lab, size=16, color=C_NAVY, bold=True); text(s, x+0.1, 5.6, 3.8, 1.0, d, size=14)
slide("Мета та принципи", ["Відстежувати поведінку кожного користувача по всій воронці: покази → клік → сайт → пошук → перегляд → кошик → покупка.",
      "Кожна подія містить ab_group, щоб порівнювати групи на рівні подій і користувачів, а не лише денних агрегатів.",
      "Один user_id/anonymous_id і session_id наскрізно від кліку на рекламі до покупки.",
      "Єдина схема назв (snake_case) та обов'язкових атрибутів; валідація на dev перед запуском."])
slide("Спільні атрибути (для кожної події)", None, [
 ["Атрибут","Тип","Опис / навіщо"],
 ["event_name, event_timestamp (UTC)","string, datetime","Що і коли сталося; побудова послідовностей"],
 ["user_id / anonymous_id","string","Ідентифікація користувача; рахунок унікальних користувачів"],
 ["session_id","string","Групування подій у сесію"],
 ["ab_group (control / test), experiment_id","string","Ключ порівняння груп; перевірка SRM"],
 ["campaign_id, ad_id, utm_source / medium / campaign / content","string","Атрибуція витрат та креативів"],
 ["device_type, os, browser, country, language","string","Сегментація, перевірка балансу груп"],
 ["page_url, referrer","string","Контекст сторінки"]], [4.3, 2.4, 5.8])
slide("Події воронки та специфічні атрибути", None, [
 ["Подія","Тригер","Специфічні атрибути"],
 ["ad_impression","Показ оголошення","placement, creative_id, bid/cost, viewable (bool)"],
 ["ad_click","Клік по оголошенню","placement, creative_id, click_cost"],
 ["session_start / landing_view","Відкриття сайту","landing_page, load_time_ms"],
 ["site_search","Пошук на сайті","search_query, results_count"],
 ["view_content","Перегляд сторінки продукту","product_id, category, price, currency"],
 ["add_to_cart","Додано в кошик","product_id, quantity, price, cart_value"],
 ["begin_checkout","Початок оформлення","cart_value, items_count, coupon"],
 ["purchase","Успішна оплата","order_id, revenue, currency, items, payment_method, is_new_customer"]], [3.0, 3.5, 6.0])
slide("Порівняння Control vs Test по кроках воронки", None, [
 ["Крок воронки","Метрика","Що показує","Формула"],
 ["Impression → Click","CTR","Наскільки реклама чіпляє","clicks / impressions"],
 ["Click → Search","Search Rate","Частка відвідувачів, що шукають","searches / clicks"],
 ["Search → View Content","View Rate","Чи доходять до товару","view_content / searches"],
 ["View → Add to Cart","Add-to-Cart Rate","Інтерес до продукту","add_to_cart / view_content"],
 ["Cart → Purchase","Cart Conversion","Якість чекауту й ціни","purchase / add_to_cart"],
 ["Click → Purchase","Purchase CVR (основна)","Підсумкова ефективність трафіку","purchase / clicks"],
 ["Вся воронка","CPA","Вартість покупки","spend / purchase"]], [2.8, 2.6, 3.7, 3.4])
slide("Метрики порівняння груп", ["Рівень користувача: унікальні користувачі на кожному кроці, конверсія між кроками (user-level), конверсія клік→покупка.",
      "Вартість: CPM, CPC, CPA = Spend / Purchase, ROAS = revenue / spend, AOV.",
      "Час: time_to_purchase (від ad_click до purchase), кількість сесій до покупки.",
      "Утримання/якість: повторні покупки за 7/30 днів, відмови (bounce), повернення в кошик.",
      "Статистика: z-тест для пропорцій, t-тест/Манна-Уїтні для середніх, ДІ різниці, bootstrap по користувачах."])
slide("Якість даних та QA", ["Перевірка SRM: співвідношення users у групах 50/50 (chi-square).",
      "Щоденні дашборд-алерти на пропуски: подія/метрика без значень (як 5.08 у контролі без метрик воронки).",
      "Воронка не має «стрибків»: purchase ≤ add_to_cart ≤ view_content за користувачами; дедуплікація подій по event_id.",
      "Збіг трекінгу з рекламним кабінетом: Spend, Impressions, Clicks (допустиме розходження <5%).",
      "Однакові правила для обох груп; не міняти розмітку під час тесту; фіксація версій трекінгу."])
slide("Що аналізуємо з цього треку", ["Де саме групи розходяться: наприклад, у тесті нижча конверсія перегляд→кошик і вища кошик→покупка — трекінг на рівні користувачів дозволить зрозуміти чому (device, джерело, продукт, ціна).",
      "Сегменти, у яких новий підхід працює краще/гірше (device, country, new/returning).",
      "Чи змінюється склад аудиторії у групах (охват/покази різко відрізнялись).",
      "Дашборди: воронка по групах, CPA/ROAS по днях, SRM, якість даних."])
p.save(f"{OUT}/{N}_Final Project-2.pptx")
print("done")
