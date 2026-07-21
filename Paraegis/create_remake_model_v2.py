"""
기공소 리메이크 손실 계산 모델 v2
- 국내 수치 중심 개편
- 공공데이터 기반 기공소 수 반영 (4,663개 영업 중)
- 해외 수치는 참고용 환율 환산 포함
- TAM/SAM/SOM 시트 추가
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

wb = openpyxl.Workbook()

# ── 공통 스타일 ──────────────────────────────────────────
def sd(c="BFBFBF", s="thin"): return Side(style=s, color=c)
def bd(c="BFBFBF"): s=sd(c); return Border(left=s, right=s, top=s, bottom=s)
def fl(h): return PatternFill("solid", fgColor=h)
def ft(bold=False, sz=11, color="000000", italic=False):
    return Font(bold=bold, size=sz, color=color, italic=italic, name="맑은 고딕")
def al(h="left", v="center", wrap=False):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

NAVY="1F4E79"; BLUE="2E75B6"; LBLUE="D6E4F0"
GREEN="375623"; LGREEN="E2EFDA"; RED="C00000"; LRED="FCE4D6"
YELLOW="FFF2CC"; LGRAY="F2F2F2"; WHITE="FFFFFF"; PURPLE="7030A0"
LPURPLE="EAD1F5"; ORANGE="C55A11"; LORANGE="FCE4D6"

def hcell(ws,r,c,v,bg=NAVY,fg=WHITE,sz=11,bold=True,h="center",wrap=True):
    cell=ws.cell(row=r,column=c,value=v)
    cell.font=ft(bold=bold,sz=sz,color=fg)
    cell.fill=fl(bg); cell.alignment=al(h=h,wrap=wrap); cell.border=bd(bg)
    return cell

def dcell(ws,r,c,v,bg=WHITE,bold=False,h="right",nf=None,italic=False,color="000000"):
    cell=ws.cell(row=r,column=c,value=v)
    cell.font=ft(bold=bold,color=color,italic=italic)
    cell.fill=fl(bg); cell.alignment=al(h=h); cell.border=bd()
    if nf: cell.number_format=nf
    return cell

def mh(ws,r,c1,c2,v,bg=NAVY,fg=WHITE,sz=12):
    ws.merge_cells(start_row=r,start_column=c1,end_row=r,end_column=c2)
    cell=ws.cell(row=r,column=c1,value=v)
    cell.font=ft(bold=True,sz=sz,color=fg)
    cell.fill=fl(bg); cell.alignment=al(h="center"); cell.border=bd(bg)
    return cell

def title(ws,txt):
    ws.row_dimensions[1].height=8
    ws.row_dimensions[2].height=38
    ws.merge_cells("B2:G2")
    c=ws["B2"]; c.value=txt
    c.font=Font(bold=True,size=18,color=WHITE,name="맑은 고딕")
    c.fill=fl(NAVY); c.alignment=al(h="center")
    ws.row_dimensions[3].height=8
    ws.sheet_view.showGridLines=False

def set_cols(ws, widths):
    for i,(col,w) in enumerate(widths,1):
        ws.column_dimensions[col].width=w

# ════════════════════════════════════════════════════════════
# SHEET 1 : 손실 계산기 (국내 중심)
# ════════════════════════════════════════════════════════════
ws1=wb.active; ws1.title="① 손실 계산기"
title(ws1,"기공소 리메이크 손실 계산기  (국내 기준)")
set_cols(ws1,[("A",2),("B",30),("C",18),("D",12),("E",20),("F",22),("G",2)])

# 섹션: 입력값
r=4; ws1.row_dimensions[r].height=26
mh(ws1,r,2,6,"🔧  기본 입력값  (노란색 셀 수정 → 자동 계산)",bg=BLUE,sz=12)

r=5; ws1.row_dimensions[r].height=20
for c,v in [(2,"항목"),(3,"값"),(4,"단위"),(5,"기준/참고"),(6,"출처")]:
    hcell(ws1,r,c,v,bg=BLUE,sz=10)

inputs=[
    ("월 총 케이스 수",         200,    "건/월",  "중소 기공소 평균 기준",       "업계 추정"),
    ("리메이크 발생률 (국내)",  0.10,   "%",      "국내 평균 10% (글로벌 4%의 2.5배)", "치의신보"),
    ("보철 종류 혼합 평균 기공료", 95000,"원",   "메탈~임플란트 가중 평균",     "업계 추정"),
    ("재료비 비율",             0.30,   "%",      "기공료의 약 30% (종류별 상이)", "업계 추정"),
    ("기공사 시간당 인건비",    18000,  "원/시간","연봉 3,765만÷2,080시간",      "고용노동부 2021"),
    ("케이스당 평균 제작 시간", 2.5,    "시간",   "보철 종류 혼합 기준",         "업계 추정"),
]

IROWS={}
for i,(lbl,val,unit,ref,src) in enumerate(inputs,start=6):
    rr=r+i-5; ws1.row_dimensions[rr].height=22
    dcell(ws1,rr,2,lbl,bg=LGRAY,h="left",bold=True)
    c=dcell(ws1,rr,3,val,bg=YELLOW,bold=True,h="center")
    nf={"원":"#,##0","원/시간":"#,##0","%":"0.0%"}.get(unit,"0.0")
    c.number_format=nf
    dcell(ws1,rr,4,unit,bg=LGRAY,h="center")
    dcell(ws1,rr,5,ref,bg=LGRAY,h="left")
    dcell(ws1,rr,6,src,bg=LGRAY,h="left",italic=True,color="595959")
    IROWS[lbl]=rr

R_CASES=IROWS["월 총 케이스 수"]
R_RATE =IROWS["리메이크 발생률 (국내)"]
R_PRICE=IROWS["보철 종류 혼합 평균 기공료"]
R_MAT  =IROWS["재료비 비율"]
R_WAGE =IROWS["기공사 시간당 인건비"]
R_HOURS=IROWS["케이스당 평균 제작 시간"]

gap=R_HOURS+1; ws1.row_dimensions[gap].height=12

# 섹션: 손실 계산
r2=gap+1; ws1.row_dimensions[r2].height=26
mh(ws1,r2,2,6,"💸  손실 계산 결과",bg=GREEN,sz=12)

r2+=1; ws1.row_dimensions[r2].height=20
for c,v in [(2,"항목"),(3,"계산 내역"),(4,"단위"),(5,"월간 손실"),(6,"연간 손실")]:
    hcell(ws1,r2,c,v,bg=GREEN,sz=10)

calcs=[
    ("① 월 리메이크 건수",  f"=C{R_CASES}*C{R_RATE}",  "건"),
    ("② 재료비 손실",       f"=C{R_CASES}*C{R_RATE}*C{R_PRICE}*C{R_MAT}", "원"),
    ("③ 인건비 손실",       f"=C{R_CASES}*C{R_RATE}*C{R_HOURS}*C{R_WAGE}", "원"),
    ("④ 기회비용",          f"=C{R_CASES}*C{R_RATE}*C{R_PRICE}*(1-C{R_MAT})", "원"),
]
descs=[
    "케이스 수 × 재제작률",
    "케이스 수 × 재제작률 × 기공료 × 재료비율",
    "케이스 수 × 재제작률 × 시간 × 시급",
    "리메이크 건수 × 기공료 × (1-재료비율) [신규케이스 기회손실]",
]

CROWS={}
for i,(lbl,formula,unit) in enumerate(calcs,start=1):
    rr=r2+i; ws1.row_dimensions[rr].height=22
    dcell(ws1,rr,2,lbl,bg=LGREEN,h="left",bold=True)
    dcell(ws1,rr,3,descs[i-1],bg=LGREEN,h="left",italic=True,color="595959")
    dcell(ws1,rr,4,unit,bg=LGREEN,h="center")

    nf="0.0" if unit=="건" else "#,##0"
    cm=ws1.cell(row=rr,column=5,value=formula)
    cm.font=ft(bold=True); cm.fill=fl(LGREEN)
    cm.alignment=al(h="right"); cm.border=bd(); cm.number_format=nf

    cy_formula=formula.replace("=","=")+"*12" if unit!="건" else formula+"*12"
    # 건은 12배 불필요하므로 연간은 비워둠
    if unit!="건":
        cy=ws1.cell(row=rr,column=6,value=f"=E{rr}*12")
        cy.font=ft(bold=True); cy.fill=fl(LGREEN)
        cy.alignment=al(h="right"); cy.border=bd(); cy.number_format=nf
    else:
        dcell(ws1,rr,6,"(월 단위)",bg=LGREEN,h="center",color="595959",italic=True)

    CROWS[lbl]=rr

# 합계
r_tot=r2+len(calcs)+1; ws1.row_dimensions[r_tot].height=30
dcell(ws1,r_tot,2,"🔴 총 손실 합계 (②+③+④)",bg=LRED,bold=True,h="left")
dcell(ws1,r_tot,3,"직접 손실 + 기회비용 합산",bg=LRED,h="left",italic=True,color="595959")
dcell(ws1,r_tot,4,"원",bg=LRED,h="center")

r2_r=CROWS["② 재료비 손실"]
r3_r=CROWS["③ 인건비 손실"]
r4_r=CROWS["④ 기회비용"]

for col,formula in [(5,f"=E{r2_r}+E{r3_r}+E{r4_r}"),(6,f"=F{r2_r}+F{r3_r}+F{r4_r}")]:
    c=ws1.cell(row=r_tot,column=col,value=formula)
    c.font=Font(bold=True,size=13,color=RED,name="맑은 고딕")
    c.fill=fl(LRED); c.alignment=al(h="right"); c.border=bd(RED)
    c.number_format="#,##0"

ROW_TOTAL_ANNUAL=r_tot  # 연간 총 손실 → F열

ws1.freeze_panes="B5"

# ════════════════════════════════════════════════════════════
# SHEET 2 : 시나리오 분석
# ════════════════════════════════════════════════════════════
ws2=wb.create_sheet("② 시나리오 분석")
title(ws2,"재제작률 개선 시나리오별 절감 효과")
set_cols(ws2,[("A",2),("B",24),("C",14),("D",18),("E",18),("F",18),("G",18),("H",2)])

r=4; ws2.row_dimensions[r].height=26
mh(ws2,r,2,7,"📊  시나리오별 연간 손실 비교  (기공소 1개소 기준, 월 200케이스)",bg=BLUE,sz=12)

r=5; ws2.row_dimensions[r].height=20
for c,v in [(2,"시나리오"),(3,"재제작률"),(4,"연간 재료비 손실"),(5,"연간 인건비 손실"),(6,"연간 기회비용"),(7,"연간 총 손실")]:
    hcell(ws2,r,c,v,bg=BLUE,sz=10)

REF="'① 손실 계산기'"
C=f"{REF}!C{R_CASES}"; PR=f"{REF}!C{R_PRICE}"
MA=f"{REF}!C{R_MAT}";  WA=f"{REF}!C{R_WAGE}"; HO=f"{REF}!C{R_HOURS}"

scenarios=[
    ("현재 (국내 평균 10%)",  0.10, LRED,   True),
    ("글로벌 평균 참고 (4%)", 0.04, LGRAY,  False),
    ("목표: 5% (개선 후)",    0.05, YELLOW, False),
    ("우수 기공소 기준 (3%)", 0.03, LGREEN, False),
    ("디지털 워크플로우 (1%)",0.01, LGREEN, False),
]

SROWS=[]
for i,(lbl,rate,bg,bold) in enumerate(scenarios,start=6):
    rr=r+i-5; ws2.row_dimensions[rr].height=22
    dcell(ws2,rr,2,lbl,bg=bg,h="left",bold=bold)
    c=ws2.cell(row=rr,column=3,value=rate)
    c.font=ft(bold=bold); c.fill=fl(bg); c.alignment=al(h="center")
    c.border=bd(); c.number_format="0.0%"

    for col,formula in [
        (4,f"={C}*C{rr}*{PR}*{MA}*12"),
        (5,f"={C}*C{rr}*{HO}*{WA}*12"),
        (6,f"={C}*C{rr}*{PR}*(1-{MA})*12"),
        (7,f"=D{rr}+E{rr}+F{rr}"),
    ]:
        cell=ws2.cell(row=rr,column=col,value=formula)
        cell.font=ft(bold=bold,color=RED if (col==7 and bold) else "000000")
        cell.fill=fl(bg); cell.alignment=al(h="right")
        cell.border=bd(); cell.number_format="#,##0"
    SROWS.append(rr)

# 절감액
gap2=SROWS[-1]+1; ws2.row_dimensions[gap2].height=12
r_s=gap2+1; ws2.row_dimensions[r_s].height=26
mh(ws2,r_s,2,7,"💡  CrownOps 도입 효과 — 현재(10%) → 목표(5%) 개선 시",bg=PURPLE,sz=12)

r_s2=r_s+1; ws2.row_dimensions[r_s2].height=28
dcell(ws2,r_s2,2,"기공소 1개소 연간 절감액",bg=LPURPLE,bold=True,h="left")
c=ws2.cell(row=r_s2,column=3,value=f"=G{SROWS[0]}-G{SROWS[2]}")
c.font=Font(bold=True,size=14,color=PURPLE,name="맑은 고딕")
c.fill=fl(LPURPLE); c.alignment=al(h="right"); c.border=bd(); c.number_format="#,##0"
ws2.merge_cells(f"C{r_s2}:D{r_s2}")
dcell(ws2,r_s2,5,"원/년",bg=LPURPLE,h="left",bold=True)
dcell(ws2,r_s2,6,"→ 구독료 대비 ROI 근거",bg=LPURPLE,h="left",italic=True,color=PURPLE)

ROW_SAVING=r_s2  # 절감액 행 (C열)
ws2.freeze_panes="B6"

# ════════════════════════════════════════════════════════════
# SHEET 3 : TAM / SAM / SOM
# ════════════════════════════════════════════════════════════
ws3=wb.create_sheet("③ TAM·SAM·SOM")
title(ws3,"시장 규모 산정  (TAM · SAM · SOM)")
set_cols(ws3,[("A",2),("B",28),("C",20),("D",16),("E",22),("F",20),("G",2)])

# ── 기공소 수 입력 섹션
r=4; ws3.row_dimensions[r].height=26
mh(ws3,r,2,6,"🏭  기공소 수 기초 데이터  (공공데이터 확정)",bg=BLUE,sz=12)
r=5; ws3.row_dimensions[r].height=20
for c,v in [(2,"항목"),(3,"수치"),(4,"단위"),(5,"기준"),(6,"출처")]:
    hcell(ws3,r,c,v,bg=BLUE,sz=10)

lab_data=[
    ("전국 기공소 전체 등록 수",    8752, "개소","공공데이터 전수",      "건강_치과기공소.csv (2026.07)"),
    ("영업 중 기공소",              4663, "개소","영업/정상 상태",        "건강_치과기공소.csv (2026.07)"),
    ("폐업·기타",                  4089, "개소","폐업+삭제+취소 합산",   "건강_치과기공소.csv (2026.07)"),
    ("서울 영업 중 기공소",         1003, "개소","지역별 1위",            "건강_치과기공소.csv (2026.07)"),
    ("경기 영업 중 기공소",          774, "개소","지역별 2위",            "건강_치과기공소.csv (2026.07)"),
]

LAB_ROWS={}
for i,(lbl,val,unit,ref,src) in enumerate(lab_data,start=6):
    rr=r+i-5; ws3.row_dimensions[rr].height=22
    bg=LBLUE if i%2==0 else WHITE
    dcell(ws3,rr,2,lbl,bg=bg,h="left",bold=(i==7))  # 영업 중이 핵심
    dcell(ws3,rr,3,val,bg=bg,h="right",bold=(i==7),nf="#,##0")
    dcell(ws3,rr,4,unit,bg=bg,h="center")
    dcell(ws3,rr,5,ref,bg=bg,h="left")
    dcell(ws3,rr,6,src,bg=bg,h="left",italic=True,color="595959")
    LAB_ROWS[lbl]=rr

R_ACTIVE=LAB_ROWS["영업 중 기공소"]  # row 7 (value=4663)

gap3=R_ACTIVE+len(lab_data)-1+1; ws3.row_dimensions[gap3].height=12

# ── TAM/SAM/SOM
r2=gap3+1; ws3.row_dimensions[r2].height=26
mh(ws3,r2,2,6,"📐  시장 규모 계산",bg=NAVY,sz=12)

r2+=1; ws3.row_dimensions[r2].height=20
for c,v in [(2,"구분"),(3,"계산식"),(4,"비율/수량"),(5,"시장 규모 (연간)"),(6,"설명")]:
    hcell(ws3,r2,c,v,bg=NAVY,sz=10)

# 연간 손실 참조 → ① 시트 F열 r_tot (연간 총 손실)
ANNUAL_LOSS=f"'① 손실 계산기'!F{ROW_TOTAL_ANNUAL}"

# SAM 비율 (중소형 기공소 비중)
SAM_RATE=0.70
SOM_RATE=0.05  # 초기 3년 목표 점유율

tam_rows=[
    ("TAM  (전체 기공소 재제작 손실)",
     f"=C{R_ACTIVE}*{ANNUAL_LOSS}",
     f"영업 중 4,663개소 전체",
     fl(LBLUE),False),
    ("SAM  (중소형 기공소 타겟)",
     f"=E{r2+1}*{SAM_RATE}",
     f"TAM의 {int(SAM_RATE*100)}% (2~10인 규모)",
     fl(LGREEN),False),
    ("SOM  (초기 3년 목표 점유)",
     f"=E{r2+2}*{SOM_RATE}",
     f"SAM의 {int(SOM_RATE*100)}% (3년 누적 목표)",
     fl(YELLOW),True),
]

TSS_ROWS=[]
for i,(lbl,formula,note,fill_style,bold) in enumerate(tam_rows,start=1):
    rr=r2+i; ws3.row_dimensions[rr].height=28
    dcell(ws3,rr,2,lbl,bg=WHITE,h="left",bold=True)
    dcell(ws3,rr,3,"→",bg=WHITE,h="center",color="595959")
    dcell(ws3,rr,4,note,bg=WHITE,h="left",italic=True,color="595959")

    c=ws3.cell(row=rr,column=5,value=formula)
    c.font=Font(bold=True,size=13 if bold else 12,
                color=PURPLE if bold else ("375623" if i==1 else "000000"),
                name="맑은 고딕")
    c.fill=fill_style; c.alignment=al(h="right"); c.border=bd(); c.number_format="#,##0"

    desc=["재제작 손실을 100% 해소하는 이상적 시장 총량",
          "실제 타겟 세그먼트 (중소형 기공소)",
          "CrownOps 현실적 초기 목표 시장"][i-1]
    dcell(ws3,rr,6,desc,bg=WHITE,h="left",italic=True,color="595959")
    TSS_ROWS.append(rr)

# 비고: SAM 비율 조정 가능 안내
note_r=TSS_ROWS[-1]+2; ws3.row_dimensions[note_r].height=20
ws3.merge_cells(f"B{note_r}:F{note_r}")
c=ws3.cell(row=note_r,column=2,value="※ SAM 비율(70%)·SOM 비율(5%)은 ① 손실 계산기의 기본값 변경에 따라 자동 반영됩니다.")
c.font=ft(sz=9,italic=True,color="595959"); c.alignment=al(h="left")

ws3.freeze_panes="B5"

# ════════════════════════════════════════════════════════════
# SHEET 4 : 해외 참고 (환율 환산)
# ════════════════════════════════════════════════════════════
ws4=wb.create_sheet("④ 해외 참고 (환율 환산)")
title(ws4,"해외 수치 참고  (환율 환산 — 참고용)")
set_cols(ws4,[("A",2),("B",26),("C",14),("D",14),("E",16),("F",24),("G",2)])

r=4; ws4.row_dimensions[r].height=26
mh(ws4,r,2,6,"💱  환율 설정  (노란색 셀 수정 가능)",bg=ORANGE,fg=WHITE,sz=12)

r=5; ws4.row_dimensions[r].height=22
for c,v in [(2,"통화"),(3,"환율 (원)"),(4,"단위"),(5,"기준일"),(6,"비고")]:
    hcell(ws4,r,c,v,bg=ORANGE,fg=WHITE,sz=10)

fx_data=[
    ("USD (미국 달러)", 1380, "원/USD", "2026.07 기준", "환율은 직접 수정하세요"),
    ("EUR (유로)",      1510, "원/EUR", "2026.07 기준", "환율은 직접 수정하세요"),
]
FX_ROWS={}
for i,(lbl,rate,unit,date,note) in enumerate(fx_data,start=6):
    rr=r+i-5; ws4.row_dimensions[rr].height=22
    dcell(ws4,rr,2,lbl,bg=LGRAY,h="left",bold=True)
    dcell(ws4,rr,3,rate,bg=YELLOW,bold=True,h="right",nf="#,##0")
    dcell(ws4,rr,4,unit,bg=LGRAY,h="center")
    dcell(ws4,rr,5,date,bg=LGRAY,h="center")
    dcell(ws4,rr,6,note,bg=LGRAY,h="left",italic=True,color="595959")
    FX_ROWS[lbl]=rr

R_USD=FX_ROWS["USD (미국 달러)"]
R_EUR=FX_ROWS["EUR (유로)"]

gap4=R_EUR+2; ws4.row_dimensions[gap4].height=12
r2=gap4+1; ws4.row_dimensions[r2].height=26
mh(ws4,r2,2,6,"🌐  주요 해외 수치 → 원화 환산",bg=BLUE,sz=12)

r2+=1; ws4.row_dimensions[r2].height=20
for c,v in [(2,"항목"),(3,"원래 수치"),(4,"통화"),(5,"환산 금액 (원)"),(6,"출처/비고")]:
    hcell(ws4,r2,c,v,bg=BLUE,sz=10)

overseas=[
    ("재제작 1건 직접 손실 (하한)", 150, "USD", f"=C{{rr}}*C{R_USD}", "Spear Education / 업계 추정"),
    ("재제작 1건 직접 손실 (상한)", 400, "USD", f"=C{{rr}}*C{R_USD}", "Spear Education / 업계 추정"),
    ("디지털 케이스 재제작 1건",    127, "USD", f"=C{{rr}}*C{R_USD}", "DDS Lab"),
    ("exocad 퍼페추얼 라이선스",   4830, "EUR", f"=C{{rr}}*C{R_EUR}", "exocad 공식 (CAD 참고용)"),
    ("LabStar 월 구독료 (추정)",    150, "USD", f"=C{{rr}}*C{R_USD}", "3Shape LabStar"),
    ("미국 기공사 시급",             21.6,"USD", f"=C{{rr}}*C{R_USD}", "PayScale 2026 / 국내 18,000원과 비교"),
]

for i,(lbl,val,cur,formula_tpl,note) in enumerate(overseas,start=1):
    rr=r2+i; ws4.row_dimensions[rr].height=22
    bg=LBLUE if i%2==0 else WHITE
    dcell(ws4,rr,2,lbl,bg=bg,h="left")
    nf="#,##0" if val>=10 else "0.0"
    dcell(ws4,rr,3,val,bg=bg,h="right",nf=nf)
    dcell(ws4,rr,4,cur,bg=bg,h="center",bold=True)
    formula=formula_tpl.replace("{rr}",str(rr))
    c=ws4.cell(row=rr,column=5,value=formula)
    c.font=ft(bold=True); c.fill=fl(bg)
    c.alignment=al(h="right"); c.border=bd(); c.number_format="#,##0"
    dcell(ws4,rr,6,note,bg=bg,h="left",italic=True,color="595959")

note_r2=r2+len(overseas)+2; ws4.row_dimensions[note_r2].height=18
ws4.merge_cells(f"B{note_r2}:F{note_r2}")
c=ws4.cell(row=note_r2,column=2,value="※ 해외 수치는 참고용입니다. 국내 TAM/SAM/SOM 산정은 ③ 시트 국내 데이터 기준을 우선 사용하세요.")
c.font=ft(sz=9,italic=True,color=RED); c.alignment=al(h="left")

ws4.freeze_panes="B5"

# ════════════════════════════════════════════════════════════
# SHEET 5 : 가정 & 출처
# ════════════════════════════════════════════════════════════
ws5=wb.create_sheet("⑤ 가정 & 출처")
title(ws5,"모델 가정 및 데이터 출처")
set_cols(ws5,[("A",2),("B",28),("C",40),("D",24),("E",2)])

r=4; ws5.row_dimensions[r].height=26
mh(ws5,r,2,4,"📋  핵심 가정 & 출처",bg=BLUE,sz=12)
r=5; ws5.row_dimensions[r].height=20
for c,v in [(2,"항목"),(3,"가정/수치"),(4,"출처")]:
    hcell(ws5,r,c,v,bg=BLUE,sz=10)

assumptions=[
    ("국내 평균 리메이크율",      "10% (복수응답 조사 기준)",                    "치의신보 (dailydental.co.kr)"),
    ("글로벌 평균 리메이크율",    "3.8~4% (미국 단일 크라운 n=3,750)",          "PMC NCBI / PMC7005880"),
    ("재료비 비율",               "기공료의 약 25~40% (종류별 상이, 평균 30%)", "업계 추정"),
    ("기공사 시간당 인건비",      "약 18,000원 (연봉 3,765만÷2,080시간)",       "고용노동부 고용형태별근로실태조사 2021"),
    ("케이스당 제작 시간",        "평균 2.5시간 (혼합 기준)",                    "업계 추정"),
    ("기회비용 정의",             "리메이크 시간에 신규 케이스 처리 시 수익 손실","자체 산정"),
    ("영업 중 기공소 수",         "4,663개소 (영업/정상 상태)",                  "건강_치과기공소.csv / 공공데이터포털 2026.07"),
    ("SAM 비율",                  "전체의 70% (중소형 2~10인 규모 추정)",        "업계 추정"),
    ("SOM 비율",                  "SAM의 5% (초기 3년 목표)",                    "CrownOps 자체 설정"),
    ("재제작 주요 원인 1위",      "인상체·재료 변형 92.2%",                      "치의신보"),
    ("커뮤니케이션 기인 비율",    "60%+ 재제작이 생산 前 커뮤니케이션 문제",    "VCAD 케이스 스터디"),
    ("USD 환율",                  "1 USD = 1,380원 (2026.07 기준)",              "환율 시트 직접 수정 가능"),
    ("EUR 환율",                  "1 EUR = 1,510원 (2026.07 기준)",              "환율 시트 직접 수정 가능"),
]

for i,(item,assumption,source) in enumerate(assumptions,start=6):
    rr=r+i-5; ws5.row_dimensions[rr].height=22
    bg=LGRAY if i%2==0 else WHITE
    dcell(ws5,rr,2,item,bg=bg,h="left",bold=True)
    dcell(ws5,rr,3,assumption,bg=bg,h="left")
    dcell(ws5,rr,4,source,bg=bg,h="left",italic=True,color="595959")

# 한계 섹션
r_lim=r+len(assumptions)+7; ws5.row_dimensions[r_lim].height=26
mh(ws5,r_lim,2,4,"⚠️  모델 한계 및 추가 검증 필요 항목",bg=RED,sz=12)

limits=[
    "월 케이스 수 분포 → 기공소 규모별 실제 조사 필요 (현재 중소 200건 가정)",
    "실제 기공소 수취 기공료 단가 → 지역별·규모별 편차 존재, 직접 확인 필요",
    "재제작 무료처리 vs. 유료청구 비율 → 현재 전량 손실로 보수적 가정",
    "CrownOps 도입에 따른 실질 개선율 → 실증 데이터 확보 후 업데이트 필요",
    "보철 종류별 케이스 비중 → 믹스 비율에 따라 평균 기공료 변동 가능",
]

for i,text in enumerate(limits,start=1):
    rr=r_lim+i; ws5.row_dimensions[rr].height=22
    ws5.merge_cells(start_row=rr,start_column=2,end_row=rr,end_column=4)
    c=ws5.cell(row=rr,column=2,value=f"• {text}")
    c.font=ft(sz=10,color=RED); c.fill=fl(LRED)
    c.alignment=al(h="left",wrap=True); c.border=bd()

# ── 저장
out="(26.07.20) 기공소 리메이크 손실 계산 모델 v2.xlsx"
wb.save(out)
print(f"완료: {out}")
