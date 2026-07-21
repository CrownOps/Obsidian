import openpyxl
from openpyxl.styles import (Font, PatternFill, Alignment, Border, Side,
                              GradientFill)
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import ColorScaleRule, DataBarRule
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.series import SeriesLabel

wb = openpyxl.Workbook()

# ── 공통 스타일 ──────────────────────────────────────────
def side(color="BFBFBF", style="thin"):
    return Side(style=style, color=color)

def border(color="BFBFBF"):
    s = side(color)
    return Border(left=s, right=s, top=s, bottom=s)

def fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)

def font(bold=False, size=11, color="000000", italic=False):
    return Font(bold=bold, size=size, color=color, italic=italic)

def align(h="left", v="center", wrap=False):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

NAVY   = "1F4E79"
BLUE   = "2E75B6"
LBLUE  = "D6E4F0"
GREEN  = "375623"
LGREEN = "E2EFDA"
RED    = "C00000"
LRED   = "FCE4D6"
YELLOW = "FFF2CC"
LGRAY  = "F2F2F2"
WHITE  = "FFFFFF"

def header_cell(ws, row, col, value, bg=NAVY, fg=WHITE, size=11, bold=True,
                h="center", colspan=None, rowspan=None):
    cell = ws.cell(row=row, column=col, value=value)
    cell.font = font(bold=bold, size=size, color=fg)
    cell.fill = fill(bg)
    cell.alignment = align(h=h, wrap=True)
    cell.border = border(bg)
    return cell

def data_cell(ws, row, col, value, bg=WHITE, bold=False, h="right",
              num_fmt=None, italic=False):
    cell = ws.cell(row=row, column=col, value=value)
    cell.font = font(bold=bold, color="000000", italic=italic)
    cell.fill = fill(bg)
    cell.alignment = align(h=h)
    cell.border = border()
    if num_fmt:
        cell.number_format = num_fmt
    return cell

def merge_header(ws, row, c1, c2, value, bg=NAVY, fg=WHITE, size=11):
    ws.merge_cells(start_row=row, start_column=c1,
                   end_row=row,   end_column=c2)
    cell = ws.cell(row=row, column=c1, value=value)
    cell.font = font(bold=True, size=size, color=fg)
    cell.fill = fill(bg)
    cell.alignment = align(h="center")
    cell.border = border(bg)
    return cell

# ════════════════════════════════════════════════════════════
# SHEET 1 : 손실 계산기
# ════════════════════════════════════════════════════════════
ws1 = wb.active
ws1.title = "① 손실 계산기"
ws1.sheet_view.showGridLines = False
ws1.column_dimensions["A"].width = 2
ws1.column_dimensions["B"].width = 28
ws1.column_dimensions["C"].width = 18
ws1.column_dimensions["D"].width = 18
ws1.column_dimensions["E"].width = 18
ws1.column_dimensions["F"].width = 18
ws1.column_dimensions["G"].width = 2

# 제목
ws1.row_dimensions[1].height = 8
ws1.row_dimensions[2].height = 36
ws1.merge_cells("B2:F2")
c = ws1["B2"]
c.value = "기공소 리메이크 손실 계산기"
c.font = Font(bold=True, size=18, color=WHITE)
c.fill = fill(NAVY)
c.alignment = align(h="center")

ws1.row_dimensions[3].height = 6

# ── 섹션 1: 기공소 기본 정보 입력 ──
r = 4
merge_header(ws1, r, 2, 6, "🔧  기본 입력값  (노란색 셀을 수정하세요)", bg=BLUE, size=12)
ws1.row_dimensions[r].height = 24

r=5; ws1.row_dimensions[r].height=20
header_cell(ws1, r, 2, "항목", bg=BLUE, size=10)
header_cell(ws1, r, 3, "값",   bg=BLUE, size=10)
header_cell(ws1, r, 4, "단위", bg=BLUE, size=10)
header_cell(ws1, r, 5, "기준/참고",   bg=BLUE, size=10)
header_cell(ws1, r, 6, "출처", bg=BLUE, size=10)

inputs = [
    ("월 총 케이스 수",          200,      "건/월",  "중소 기공소 평균",         "업계 추정"),
    ("리메이크 발생률",          0.10,     "%",      "국내 평균 10%",            "치의신보"),
    ("평균 기공료 (케이스당)",   90000,    "원",     "메탈~지르코니아 평균",     "업계 추정"),
    ("재료비 비율",              0.30,     "%",      "기공료의 약 30%",          "업계 추정"),
    ("기공사 시간당 인건비",     18000,    "원/시간","연봉 3,765만÷2,080h",      "고용노동부 2021"),
    ("케이스당 평균 제작 시간",  2.5,      "시간",   "보철 종류 혼합 기준",      "업계 추정"),
]

INPUT_ROWS = {}
for i, (label, val, unit, ref, src) in enumerate(inputs, start=6):
    rr = r + i - 5
    ws1.row_dimensions[rr].height = 22
    data_cell(ws1, rr, 2, label, bg=LGRAY, h="left")
    c = data_cell(ws1, rr, 3, val,   bg=YELLOW, bold=True, h="center")
    if unit == "%":
        c.number_format = "0.0%"
    elif unit in ("원", "원/시간"):
        c.number_format = "#,##0"
    else:
        c.number_format = "0.0"
    data_cell(ws1, rr, 4, unit,  bg=LGRAY, h="center")
    data_cell(ws1, rr, 5, ref,   bg=LGRAY, h="left")
    data_cell(ws1, rr, 6, src,   bg=LGRAY, h="left", italic=True)
    INPUT_ROWS[label] = rr

# named references (row numbers)
R_CASES    = INPUT_ROWS["월 총 케이스 수"]
R_RATE     = INPUT_ROWS["리메이크 발생률"]
R_PRICE    = INPUT_ROWS["평균 기공료 (케이스당)"]
R_MAT      = INPUT_ROWS["재료비 비율"]
R_WAGE     = INPUT_ROWS["기공사 시간당 인건비"]
R_HOURS    = INPUT_ROWS["케이스당 평균 제작 시간"]

gap = R_HOURS + 1
ws1.row_dimensions[gap].height = 10

# ── 섹션 2: 손실 계산 ──
r2 = gap + 1
merge_header(ws1, r2, 2, 6, "💸  손실 계산 결과", bg="375623", size=12)
ws1.row_dimensions[r2].height = 24

r2 += 1; ws1.row_dimensions[r2].height = 20
header_cell(ws1, r2, 2, "항목",       bg="375623", size=10)
header_cell(ws1, r2, 3, "계산식",     bg="375623", size=10)
header_cell(ws1, r2, 4, "월간 손실",  bg="375623", size=10)
header_cell(ws1, r2, 5, "연간 손실",  bg="375623", size=10)
header_cell(ws1, r2, 6, "비고",       bg="375623", size=10)

# 계산행 정의: (label, formula_monthly, note)
calc_rows = [
    ("① 월 리메이크 건수",
     f"=C{R_CASES}*C{R_RATE}",
     None,
     "건",
     f"=C{R_CASES}*C{R_RATE}*12"),
    ("② 재료비 손실",
     f"=C{R_CASES}*C{R_RATE}*C{R_PRICE}*C{R_MAT}",
     "케이스수×재제작률×기공료×재료비율",
     "원",
     f"=C{R_CASES}*C{R_RATE}*C{R_PRICE}*C{R_MAT}*12"),
    ("③ 인건비 손실",
     f"=C{R_CASES}*C{R_RATE}*C{R_HOURS}*C{R_WAGE}",
     "케이스수×재제작률×시간×시급",
     "원",
     f"=C{R_CASES}*C{R_RATE}*C{R_HOURS}*C{R_WAGE}*12"),
    ("④ 기회비용",
     f"=C{R_CASES}*C{R_RATE}*C{R_PRICE}*(1-C{R_MAT})",
     "리메이크 건수×기공료×(1-재료비율)",
     "원",
     f"=C{R_CASES}*C{R_RATE}*C{R_PRICE}*(1-C{R_MAT})*12"),
]

CALC_R = {}
for i, (label, formula, note, unit, formula_y) in enumerate(calc_rows, start=1):
    rr = r2 + i
    ws1.row_dimensions[rr].height = 22
    data_cell(ws1, rr, 2, label, bg=LGREEN, h="left", bold=True)

    # 월간
    cm = ws1.cell(row=rr, column=3, value=formula)
    cm.font = font()
    cm.fill = fill(LGREEN)
    cm.alignment = align(h="right")
    cm.border = border()
    if unit == "원":
        cm.number_format = "#,##0"
    else:
        cm.number_format = "0.0"

    # 계산식 설명 (col D → 월간, col E → 연간)
    # 재배치: B=항목, C=월간, D=연간, E=비고 (헤더와 맞춤)
    # 위 헤더: B항목 C계산식 D월간 E연간 F비고 → 실제 데이터를 맞춰야함
    # 헤더 순서: B=항목, C=계산식(월간값), D=월간손실, E=연간손실, F=비고
    # C에 월간 값 공식, D에 월간 값(=C), E에 연간값, F에 비고

    # 재정의: C = 월간 값 공식, D = 월간 표시(=C), E = 연간, F = 비고
    cm.value = formula

    cd = ws1.cell(row=rr, column=4, value=f"=C{rr}")
    cd.font = font()
    cd.fill = fill(LGREEN)
    cd.alignment = align(h="right")
    cd.border = border()
    cd.number_format = "#,##0" if unit == "원" else "0.0"

    ce = ws1.cell(row=rr, column=5, value=formula_y)
    ce.font = font(bold=True)
    ce.fill = fill(LGREEN)
    ce.alignment = align(h="right")
    ce.border = border()
    ce.number_format = "#,##0" if unit == "원" else "0.0"

    cf = ws1.cell(row=rr, column=6, value=note or "")
    cf.font = font(italic=True, size=9)
    cf.fill = fill(LGREEN)
    cf.alignment = align(h="left")
    cf.border = border()

    CALC_R[label] = rr

# 합계행
r_total = r2 + len(calc_rows) + 1
ws1.row_dimensions[r_total].height = 28
data_cell(ws1, r_total, 2, "🔴 총 손실 합계 (②+③+④)", bg=LRED, bold=True, h="left")

r2_row = CALC_R["② 재료비 손실"]
r3_row = CALC_R["③ 인건비 손실"]
r4_row = CALC_R["④ 기회비용"]

ct = ws1.cell(row=r_total, column=3,
              value=f"=C{r2_row}+C{r3_row}+C{r4_row}")
ct.font = font(bold=True, color=RED)
ct.fill = fill(LRED)
ct.alignment = align(h="right")
ct.border = border(RED)
ct.number_format = "#,##0"

ws1.cell(row=r_total, column=4,
         value=f"=D{r2_row}+D{r3_row}+D{r4_row}").number_format = "#,##0"
ws1.cell(row=r_total, column=4).font = font(bold=True, color=RED)
ws1.cell(row=r_total, column=4).fill = fill(LRED)
ws1.cell(row=r_total, column=4).alignment = align(h="right")
ws1.cell(row=r_total, column=4).border = border(RED)

ct_y = ws1.cell(row=r_total, column=5,
                value=f"=E{r2_row}+E{r3_row}+E{r4_row}")
ct_y.font = font(bold=True, size=13, color=RED)
ct_y.fill = fill(LRED)
ct_y.alignment = align(h="right")
ct_y.border = border(RED)
ct_y.number_format = "#,##0"

ws1.cell(row=r_total, column=6, value="연간 총 손실").font = font(bold=True, color=RED)
ws1.cell(row=r_total, column=6).fill = fill(LRED)
ws1.cell(row=r_total, column=6).border = border(RED)

ws1.freeze_panes = "B5"

# ════════════════════════════════════════════════════════════
# SHEET 2 : 시나리오 분석
# ════════════════════════════════════════════════════════════
ws2 = wb.create_sheet("② 시나리오 분석")
ws2.sheet_view.showGridLines = False
for col, w in zip("ABCDEFGH", [2,28,16,16,16,16,16,2]):
    ws2.column_dimensions[get_column_letter(ord(col)-64)].width = w

ws2.row_dimensions[1].height = 8
ws2.row_dimensions[2].height = 36
ws2.merge_cells("B2:G2")
c = ws2["B2"]
c.value = "리메이크 개선 시나리오별 절감 효과"
c.font = Font(bold=True, size=18, color=WHITE)
c.fill = fill(NAVY)
c.alignment = align(h="center")

ws2.row_dimensions[3].height = 6

r = 4
merge_header(ws2, r, 2, 7, "📊  시나리오별 연간 손실 비교", bg=BLUE, size=12)
ws2.row_dimensions[r].height = 24

r=5; ws2.row_dimensions[r].height=22
for col, label in enumerate(["시나리오","재제작률","월 리메이크","연간 재료비 손실","연간 인건비 손실","연간 기회비용","연간 총 손실"], start=2):
    header_cell(ws2, r, col, label, bg=BLUE, size=10)

# 시나리오 데이터 (재제작률, 라벨, 배경색)
scenarios = [
    ("현재 (국내 평균)",    0.10, LRED),
    ("글로벌 평균",         0.04, YELLOW),
    ("목표 개선 (5%)",      0.05, YELLOW),
    ("우수 기공소 (3%)",    0.03, LGREEN),
    ("이상적 목표 (1%)",    0.01, LGREEN),
]

# ① 시트 입력값 참조 (같은 파일 내)
REF = "'① 손실 계산기'"
CASES = f"{REF}!C{R_CASES}"
PRICE = f"{REF}!C{R_PRICE}"
MAT   = f"{REF}!C{R_MAT}"
WAGE  = f"{REF}!C{R_WAGE}"
HOURS = f"{REF}!C{R_HOURS}"

scen_rows = []
for i, (label, rate, bg) in enumerate(scenarios, start=6):
    rr = r + i - 5
    ws2.row_dimensions[rr].height = 22
    data_cell(ws2, rr, 2, label, bg=bg, h="left", bold=(i==6))
    rc = ws2.cell(row=rr, column=3, value=rate)
    rc.font = font(bold=True)
    rc.fill = fill(bg)
    rc.alignment = align(h="center")
    rc.border = border()
    rc.number_format = "0.0%"

    remake_cnt = f"={CASES}*C{rr}"
    mat_loss   = f"={CASES}*C{rr}*{PRICE}*{MAT}*12"
    lab_loss   = f"={CASES}*C{rr}*{HOURS}*{WAGE}*12"
    opp_loss   = f"={CASES}*C{rr}*{PRICE}*(1-{MAT})*12"
    total_loss = f"=D{rr}+E{rr}+F{rr}"

    for col, (val, fmt) in enumerate([(remake_cnt,"0.0"),(mat_loss,"#,##0"),
                                       (lab_loss,"#,##0"),(opp_loss,"#,##0"),
                                       (total_loss,"#,##0")], start=4):
        cell = ws2.cell(row=rr, column=col, value=val)
        cell.font = font(bold=(col==8))
        cell.fill = fill(bg)
        cell.alignment = align(h="right")
        cell.border = border()
        cell.number_format = fmt

    scen_rows.append(rr)

# 절감액 행
gap2 = scen_rows[-1] + 1
ws2.row_dimensions[gap2].height = 10

r_sav = gap2 + 1
merge_header(ws2, r_sav, 2, 7, "💡  CrownOps 도입 시 절감 효과 (현재 → 목표 개선 5%)", bg="7030A0", fg=WHITE, size=12)
ws2.row_dimensions[r_sav].height = 24

r_sav2 = r_sav + 1
ws2.row_dimensions[r_sav2].height = 28
data_cell(ws2, r_sav2, 2, "연간 절감액", bg=YELLOW, bold=True, h="left")
# 현재(row scen_rows[0]) - 목표(row scen_rows[2])
saving_formula = f"=G{scen_rows[0]}-G{scen_rows[2]}"
cs = ws2.cell(row=r_sav2, column=3, value=saving_formula)
cs.font = Font(bold=True, size=14, color="7030A0")
cs.fill = fill(YELLOW)
cs.alignment = align(h="right")
cs.border = border()
cs.number_format = "#,##0"
ws2.merge_cells(start_row=r_sav2, start_column=3, end_row=r_sav2, end_column=4)

data_cell(ws2, r_sav2, 5, "원/년", bg=YELLOW, h="left")
data_cell(ws2, r_sav2, 6, "→ ROI 근거로 활용 가능", bg=YELLOW, h="left", italic=True)

ws2.freeze_panes = "B6"

# ════════════════════════════════════════════════════════════
# SHEET 3 : 보철 종류별 분석
# ════════════════════════════════════════════════════════════
ws3 = wb.create_sheet("③ 보철 종류별 분석")
ws3.sheet_view.showGridLines = False
for col, w in zip(range(1,10), [2,22,12,12,12,14,14,14,2]):
    ws3.column_dimensions[get_column_letter(col)].width = w

ws3.row_dimensions[1].height = 8
ws3.row_dimensions[2].height = 36
ws3.merge_cells("B2:H2")
c = ws3["B2"]
c.value = "보철 종류별 리메이크 손실 분석"
c.font = Font(bold=True, size=18, color=WHITE)
c.fill = fill(NAVY)
c.alignment = align(h="center")

ws3.row_dimensions[3].height = 6

r = 4
merge_header(ws3, r, 2, 8, "🦷  보철 종류별 월간 손실", bg=BLUE, size=12)
ws3.row_dimensions[r].height = 24

r=5; ws3.row_dimensions[r].height=22
for col, label in enumerate(["보철 종류","월 케이스","리메이크율","기공료(원)","재료비율",
                               "월 손실(재료+인건비)","월 기회비용"], start=2):
    header_cell(ws3, r, col, label, bg=BLUE, size=10)

# 보철 종류별 데이터 (국내 재제작률 기준)
prosthetics = [
    ("메탈 크라운",      40, 0.058, 65000,  0.25),
    ("PFM 크라운",       50, 0.070, 95000,  0.30),
    ("지르코니아 크라운",60, 0.075, 110000, 0.28),
    ("임플란트 보철",    25, 0.080, 130000, 0.32),
    ("가철성 의치",      15, 0.098, 150000, 0.40),
    ("심미보철",         10, 0.095, 200000, 0.35),
]

wage_ref  = f"'① 손실 계산기'!C{R_WAGE}"
hours_ref = f"'① 손실 계산기'!C{R_HOURS}"

type_rows = []
for i, (name, cases, rate, price, mat) in enumerate(prosthetics, start=6):
    rr = r + i - 5
    ws3.row_dimensions[rr].height = 22
    bg = LBLUE if i % 2 == 0 else WHITE

    data_cell(ws3, rr, 2, name,  bg=bg, h="left", bold=True)
    data_cell(ws3, rr, 3, cases, bg=bg, h="center", num_fmt="0")
    c_rate = ws3.cell(row=rr, column=4, value=rate)
    c_rate.number_format = "0.0%"
    c_rate.fill = fill(bg); c_rate.alignment = align(h="center"); c_rate.border = border()
    data_cell(ws3, rr, 5, price, bg=bg, num_fmt="#,##0")
    c_mat = ws3.cell(row=rr, column=6, value=mat)
    c_mat.number_format = "0.0%"
    c_mat.fill = fill(bg); c_mat.alignment = align(h="center"); c_mat.border = border()

    # 월 손실(재료비+인건비)
    mat_loss = f"=C{rr}*D{rr}*E{rr}*F{rr}+C{rr}*D{rr}*{hours_ref}*{wage_ref}"
    c7 = ws3.cell(row=rr, column=7, value=mat_loss)
    c7.number_format = "#,##0"; c7.fill = fill(bg)
    c7.alignment = align(h="right"); c7.border = border()

    # 월 기회비용
    opp = f"=C{rr}*D{rr}*E{rr}*(1-F{rr})"
    c8 = ws3.cell(row=rr, column=8, value=opp)
    c8.number_format = "#,##0"; c8.fill = fill(bg)
    c8.alignment = align(h="right"); c8.border = border()

    type_rows.append(rr)

# 합계
r_sum3 = type_rows[-1] + 1
ws3.row_dimensions[r_sum3].height = 26
data_cell(ws3, r_sum3, 2, "합계", bg=LRED, bold=True, h="left")
ws3.cell(row=r_sum3, column=3,
         value=f"=SUM(C{type_rows[0]}:C{type_rows[-1]})").number_format = "0"
ws3.cell(row=r_sum3, column=3).fill = fill(LRED)
ws3.cell(row=r_sum3, column=3).alignment = align(h="center"); ws3.cell(row=r_sum3, column=3).border = border()
ws3.cell(row=r_sum3, column=3).font = font(bold=True)

for col in [7, 8]:
    c = ws3.cell(row=r_sum3, column=col,
                 value=f"=SUM({get_column_letter(col)}{type_rows[0]}:{get_column_letter(col)}{type_rows[-1]})")
    c.number_format = "#,##0"
    c.font = font(bold=True, color=RED)
    c.fill = fill(LRED)
    c.alignment = align(h="right")
    c.border = border(RED)

ws3.freeze_panes = "B6"

# ════════════════════════════════════════════════════════════
# SHEET 4 : 가정 & 출처
# ════════════════════════════════════════════════════════════
ws4 = wb.create_sheet("④ 가정 & 출처")
ws4.sheet_view.showGridLines = False
ws4.column_dimensions["A"].width = 2
ws4.column_dimensions["B"].width = 30
ws4.column_dimensions["C"].width = 40
ws4.column_dimensions["D"].width = 30
ws4.column_dimensions["E"].width = 2

ws4.row_dimensions[1].height = 8
ws4.row_dimensions[2].height = 36
ws4.merge_cells("B2:D2")
c = ws4["B2"]
c.value = "모델 가정 및 데이터 출처"
c.font = Font(bold=True, size=18, color=WHITE)
c.fill = fill(NAVY)
c.alignment = align(h="center")

ws4.row_dimensions[3].height = 6

r = 4; ws4.row_dimensions[r].height = 24
merge_header(ws4, r, 2, 4, "📋  핵심 가정 & 출처", bg=BLUE, size=12)

r=5; ws4.row_dimensions[r].height=20
header_cell(ws4, r, 2, "항목",     bg=BLUE)
header_cell(ws4, r, 3, "가정/수치", bg=BLUE)
header_cell(ws4, r, 4, "출처",     bg=BLUE)

assumptions = [
    ("국내 평균 리메이크율",    "10% (글로벌 평균 3.8%의 약 2.5배)",      "치의신보 / NCBI PMC7005880"),
    ("보철 재료비 비율",        "기공료의 약 25~40% (종류별 상이)",         "업계 추정"),
    ("기공사 시간당 인건비",    "약 18,000원 (연봉 3,765만 ÷ 2,080시간)", "고용노동부 고용형태별근로실태조사 2021"),
    ("케이스당 제작 시간",      "평균 2.5시간 (보철 종류 혼합 기준)",       "업계 추정"),
    ("기회비용 정의",           "리메이크 시간에 신규 케이스를 처리했을 경우의 수익 손실", "자체 산정"),
    ("리메이크 주요 원인",      "인상체 변형 92.2% / 정보 오류 68.6% / 기공사 오류 62.7%", "치의신보"),
    ("국내 기공소 수",          "약 4,500개소 (추정)",                      "업계 추정 (공식 통계 확인 필요)"),
    ("월 평균 케이스 수",       "중소 기공소 기준 200건/월",                 "업계 추정"),
    ("글로벌 비교 기준",        "미국 dental lab 평균 3.8% (n=3,750 크라운)", "NCBI PMC7005880 (2020)"),
    ("CrownOps 개선 목표",      "리메이크율 10% → 5% (수주/커뮤니케이션 개선)", "자체 설정"),
]

for i, (item, assumption, source) in enumerate(assumptions, start=6):
    rr = r + i - 5
    ws4.row_dimensions[rr].height = 22
    bg = LGRAY if i % 2 == 0 else WHITE
    data_cell(ws4, rr, 2, item,       bg=bg, h="left", bold=True)
    data_cell(ws4, rr, 3, assumption, bg=bg, h="left")
    data_cell(ws4, rr, 4, source,     bg=bg, h="left", italic=True)

# 한계 섹션
r_lim = r + len(assumptions) + 7
ws4.row_dimensions[r_lim].height = 24
merge_header(ws4, r_lim, 2, 4, "⚠️  모델 한계 및 추가 검증 필요 항목", bg=RED, size=12)

limits = [
    "전국 기공소 정확한 수 → data.go.kr '전국치과기공소표준데이터' CSV 다운로드 필요",
    "보철 종류별 월 케이스 수 분포 → 기공소 직접 인터뷰 또는 설문 필요",
    "실제 기공료 단가 → 지역별·규모별 편차 존재, 직접 확인 필요",
    "리메이크 이후 재작업 여부 추적 데이터 → 현재 국내 데이터 없음",
    "CrownOps 도입에 따른 실질 개선율 → 실증 데이터 확보 후 업데이트",
]

for i, text in enumerate(limits, start=1):
    rr = r_lim + i
    ws4.row_dimensions[rr].height = 22
    ws4.merge_cells(start_row=rr, start_column=2, end_row=rr, end_column=4)
    c = ws4.cell(row=rr, column=2, value=f"• {text}")
    c.font = Font(size=10, color="C00000")
    c.fill = fill(LRED)
    c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    c.border = border()

# ── 저장 ────────────────────────────────────────────────────
out = "(26.07.20) 기공소 리메이크 손실 계산 모델.xlsx"
wb.save(out)
print(f"완료: {out}")
