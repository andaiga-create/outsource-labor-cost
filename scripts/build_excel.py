# -*- coding: utf-8 -*-
"""검증용 엑셀 보고서 생성 (update.bat 실행 시 build_data.py 가 호출)

시트
  대시보드   : 주요 지표 + 차트 (집계 시트 값을 참조)
  집계       : 연도별·월별·장비·호선·업체·선종 집계. 모든 금액은 RAW_통합 을 SUMIFS 로 다시 계산한 수식
  RAW_통합   : 대시보드에 들어간 전체 행 (원본 파일명 포함)
  제외내역   : 2단계 규칙으로 빠진 행과 사유
엑셀 파일은 원본처럼 GitHub 에 올라가지 않는다 (.gitignore: 엑셀보고서/).
"""
import os
from collections import defaultdict
from openpyxl import Workbook
from openpyxl.chart import BarChart, PieChart, LineChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.table import Table, TableStyleInfo

KINDS = ['AS', '김해공장']
SEGS = ['상선', '특수선', '유상']
BLANK = '(미분류)'
WON = '#,##0'
HEAD = PatternFill('solid', fgColor='4F46E5')
SUB = PatternFill('solid', fgColor='EEF1F6')
THIN = Border(bottom=Side(style='thin', color='D5DAE3'))
R = "RAW_통합!"
COL = {'연': '$B:$B', '월': '$C:$C', '구분': '$D:$D', '업체': '$E:$E', '금액': '$K:$K',
       '고객': '$O:$O', '호선': '$P:$P', '장비': '$Q:$Q', '선종': '$R:$R'}


def q(v):
    return '"' + str(v).replace('"', '""') + '"'


def sumifs(*conds):
    """conds: (열이름, 조건식 문자열) — 조건식은 엑셀 수식 조각(숫자, "문자", "<>유상")"""
    parts = [f"{R}{COL['금액']}"]
    for c, v in conds:
        parts += [f"{R}{COL[c]}", str(v)]
    return f"=SUMIFS({','.join(parts)})"


def header(ws, row, col, labels, widths=None):
    for i, t in enumerate(labels):
        c = ws.cell(row, col + i, t)
        c.font = Font(bold=True, color='FFFFFF')
        c.fill = HEAD
        c.alignment = Alignment(horizontal='center')


def title(ws, row, col, text, note=None):
    ws.cell(row, col, text).font = Font(bold=True, size=12)
    if note:
        ws.cell(row + 1, col, note).font = Font(size=9, color='7A8599')


def write(root, records, excluded, latest, settings):
    ly, lm = map(int, latest.split('-'))
    rev = {int(k): v for k, v in (settings.get('revenue') or {}).items()}
    tgt = (settings.get('target') or {}).get(str(ly))
    wb = Workbook()
    dash = wb.active
    dash.title = '대시보드'
    agg = wb.create_sheet('집계')
    raw = wb.create_sheet('RAW_통합')
    exc = wb.create_sheet('제외내역')

    # ---------------- RAW_통합 ----------------
    cols = ['입고일자', '연', '월', '구분', '거래처명', '품목번호', '품명', '단위', '수량', '단가', '공급가',
            '프로젝트코드', '프로젝트명', '비고', '고객', '호선', '장비', '상선/특수선', '원본파일']
    raw.append(cols)
    for r in sorted(records, key=lambda x: (x['date'], x['kind'], x['vendor'])):
        raw.append([r['date'], r['y'], r['m'], r['kind'], r['vendor'], r['pn'], r['name'], r['unit'], r['qty'],
                    r['price'], r['amt'], r['code'], r['pname'], r['note'], r['cust'], r['ship'], r['equip'],
                    r['seg'], r['file']])
    n = len(records) + 1
    for row in raw.iter_rows(min_row=2, max_row=n):
        row[0].number_format = 'yyyy-mm-dd'
        for c in (row[9], row[10]):
            c.number_format = WON
    raw.add_table(Table(displayName='RAW_통합', ref=f'A1:{L(len(cols))}{n}',
                        tableStyleInfo=TableStyleInfo(name='TableStyleLight9', showRowStripes=True)))
    for i, w in enumerate([11, 6, 5, 9, 20, 12, 10, 6, 7, 11, 12, 22, 30, 30, 8, 10, 10, 10, 40], 1):
        raw.column_dimensions[L(i)].width = w
    raw.freeze_panes = 'A2'

    # ---------------- 제외내역 ----------------
    ecols = ['사유', '구분', '입고일자', '거래처명', '품목번호', '품명', '공급가', '프로젝트코드', '프로젝트명', '비고', '원본파일']
    exc.append(ecols)
    for e in sorted(excluded, key=lambda x: (x['사유'], str(x['입고일']))):
        exc.append([e['사유'], e['구분'], e['입고일'], e.get('거래처명'), e.get('품목번호'), e.get('품명'),
                    e['공급가_'], e.get('프로젝트코드'), e.get('프로젝트명'), e.get('비고'), e['파일']])
    m = len(excluded) + 1
    for row in exc.iter_rows(min_row=2, max_row=m):
        row[2].number_format = 'yyyy-mm-dd'
        row[6].number_format = WON
    if excluded:
        exc.add_table(Table(displayName='제외내역', ref=f'A1:{L(len(ecols))}{m}',
                            tableStyleInfo=TableStyleInfo(name='TableStyleLight15', showRowStripes=True)))
    for i, w in enumerate([20, 9, 11, 20, 14, 16, 12, 24, 30, 30, 40], 1):
        exc.column_dimensions[L(i)].width = w
    exc.freeze_panes = 'A2'

    # ---------------- 집계 ----------------
    agg.column_dimensions['A'].width = 16
    for c in 'BCDEFGHI':
        agg.column_dimensions[c].width = 15
    agg['A1'] = '외주인건비 집계 (모든 금액은 RAW_통합 시트를 SUMIFS 로 다시 계산한 수식입니다)'
    agg['A1'].font = Font(bold=True, size=13)
    agg['A2'] = '기본: 유상 제외 (엑셀 월간 보고서 피벗과 동일). 선종별 표만 유상 포함.'
    agg['A2'].font = Font(size=9, color='7A8599')

    years = sorted({r['y'] for r in records})
    pos = {}
    # 1) 연도별
    r0 = 4
    title(agg, r0, 1, '연도별 외주인건비')
    header(agg, r0 + 1, 1, ['연도', 'AS', '김해공장', '합계(유상 제외)', '유상', '합계(유상 포함)', '매출액', '매출 대비'])
    for i, y in enumerate(years):
        rr = r0 + 2 + i
        agg.cell(rr, 1, y)
        agg.cell(rr, 2, sumifs(('연', y), ('구분', q('AS')), ('선종', q('<>유상'))))
        agg.cell(rr, 3, sumifs(('연', y), ('구분', q('김해공장')), ('선종', q('<>유상'))))
        agg.cell(rr, 4, f'=B{rr}+C{rr}')
        agg.cell(rr, 5, sumifs(('연', y), ('선종', q('유상'))))
        agg.cell(rr, 6, f'=D{rr}+E{rr}')
        agg.cell(rr, 7, rev.get(y))
        agg.cell(rr, 8, f'=IF(G{rr}="","",D{rr}/G{rr})')
        for c in range(2, 8):
            agg.cell(rr, c).number_format = WON
        agg.cell(rr, 8).number_format = '0.00%'
    pos['year'] = (r0 + 2, r0 + 1 + len(years))

    # 2) 최신 연도 월별
    r0 = pos['year'][1] + 3
    title(agg, r0, 1, f'{ly}년 월별 (유상 제외)')
    header(agg, r0 + 1, 1, ['월', '김해공장', 'AS', '합계', '누계', '전년 누계'])
    for i in range(12):
        mm, rr = i + 1, r0 + 2 + i
        agg.cell(rr, 1, f'{mm:02d}월')
        agg.cell(rr, 2, sumifs(('연', ly), ('월', mm), ('구분', q('김해공장')), ('선종', q('<>유상'))))
        agg.cell(rr, 3, sumifs(('연', ly), ('월', mm), ('구분', q('AS')), ('선종', q('<>유상'))))
        agg.cell(rr, 4, f'=B{rr}+C{rr}')
        agg.cell(rr, 5, f'=SUM($D${r0 + 2}:D{rr})' if mm <= lm else None)
        agg.cell(rr, 6, f'=SUMIFS({R}{COL["금액"]},{R}{COL["연"]},{ly - 1},{R}{COL["월"]},"<="&{mm},{R}{COL["선종"]},"<>유상")')
        for c in range(2, 7):
            agg.cell(rr, c).number_format = WON
    tr = r0 + 14
    agg.cell(tr, 1, '합계').font = Font(bold=True)
    for c in (2, 3, 4):
        agg.cell(tr, c, f'=SUM({L(c)}{r0 + 2}:{L(c)}{r0 + 13})').number_format = WON
        agg.cell(tr, c).font = Font(bold=True)
    pos['month'] = (r0 + 2, r0 + 13)
    pos['ytotal'] = f'D{tr}'

    # 3) Top 표 (최신 연도, 유상 제외) — 이름은 금액 순으로 정렬, 금액은 수식
    def top(r0, label, key, field, limit):
        tot = defaultdict(int)
        for r in records:
            if r['y'] == ly and r['seg'] != '유상':
                tot[r[key]] += r['amt']
        items = sorted(((k, v) for k, v in tot.items() if v), key=lambda x: -x[1])[:limit]
        title(agg, r0, 1, f'{ly}년 {label} (유상 제외)')
        header(agg, r0 + 1, 1, [field, '공급가'])
        for i, (k, _) in enumerate(items):
            rr = r0 + 2 + i
            agg.cell(rr, 1, k if k else BLANK)
            agg.cell(rr, 2, sumifs(('연', ly), (field if field in COL else '업체', q(k)), ('선종', q('<>유상'))))
            agg.cell(rr, 2).number_format = WON
        return r0 + 2, r0 + 1 + len(items)

    r0 = tr + 3
    pos['equip'] = top(r0, '장비별 TOP 15', 'equip', '장비', 15)
    pos['ship'] = top(pos['equip'][1] + 3, '호선별 TOP 10', 'ship', '호선', 10)
    pos['vendor'] = top(pos['ship'][1] + 3, '업체별', 'vendor', '업체', 50)

    # 4) 선종별 (유상 포함)
    r0 = pos['vendor'][1] + 3
    title(agg, r0, 1, f'{ly}년 선종별 (유상 포함)')
    header(agg, r0 + 1, 1, ['선종', '공급가'])
    for i, s in enumerate(SEGS):
        rr = r0 + 2 + i
        agg.cell(rr, 1, s)
        agg.cell(rr, 2, sumifs(('연', ly), ('선종', q(s)))).number_format = WON
    pos['seg'] = (r0 + 2, r0 + 4)

    # 5) 행 수 대조
    r0 = pos['seg'][1] + 3
    title(agg, r0, 1, '행 수 대조')
    agg.cell(r0 + 1, 1, 'RAW_통합 행 수')
    agg.cell(r0 + 1, 2, '=COUNTA(RAW_통합!$A:$A)-1')
    agg.cell(r0 + 2, 1, '제외된 행 수')
    agg.cell(r0 + 2, 2, '=COUNTA(제외내역!$A:$A)-1')
    agg.freeze_panes = 'A4'

    # ---------------- 대시보드 ----------------
    dash.sheet_view.showGridLines = False
    dash['A1'] = f'외주인건비 현황 ({ly}년 {lm}월 기준)'
    dash['A1'].font = Font(bold=True, size=18)
    dash['A2'] = '유상 제외 기준 · 금액은 집계 시트(수식) 참조 · 웹 대시보드와 같은 규칙'
    dash['A2'].font = Font(size=9, color='7A8599')
    yrow = pos['year'][0] + years.index(ly)
    prow = yrow - 1 if ly - 1 in years else None
    kpis = [('올해 누계', f'=집계!D{yrow}', WON),
            ('연간 목표', tgt if tgt else '-', WON),
            ('목표 대비', f'=IF(ISNUMBER(C5),A5/C5,"-")', '0%'),
            ('전년 동기', f'=집계!F{pos["month"][0] + lm - 1}' if prow else '-', WON),
            ('전년 동기 대비', f'=IF(ISNUMBER(G5),A5/G5-1,"-")' if prow else '-', '+0.0%;-0.0%'),
            ('매출 대비', f'=집계!H{yrow}', '0.00%')]
    for i, (t, f, fmt) in enumerate(kpis):
        c = 1 + i * 2
        dash.cell(4, c, t).font = Font(size=10, color='7A8599')
        v = dash.cell(5, c, f)
        v.font = Font(bold=True, size=14)
        v.alignment = Alignment(horizontal='left')
        v.number_format = fmt
        for cc in (c, c + 1):
            dash.cell(4, cc).fill = SUB
            dash.cell(5, cc).fill = SUB
        dash.merge_cells(start_row=4, start_column=c, end_row=4, end_column=c + 1)
        dash.merge_cells(start_row=5, start_column=c, end_row=5, end_column=c + 1)
    for c in range(1, 13):
        dash.column_dimensions[L(c)].width = 11

    def bar(title_, cats, vals_cols, anchor, stacked=False, horizontal=False, w=17, h=8.5):
        ch = BarChart()
        ch.type = 'bar' if horizontal else 'col'
        ch.title = title_
        ch.height, ch.width = h, w
        ch.legend.position = 'b'
        for col, first, last, hdr in vals_cols:
            ch.add_data(Reference(agg, min_col=col, min_row=hdr, max_row=last), titles_from_data=True)
        ch.set_categories(cats)
        if stacked:
            ch.grouping, ch.overlap = 'stacked', 100
        if len(vals_cols) == 1:
            ch.legend = None
        if horizontal:
            ch.x_axis.scaling.orientation = 'maxMin'
        ch.y_axis.numFmt = '#,##0,,"백만"'
        ch.y_axis.majorGridlines = None
        ch.x_axis.delete = False
        ch.y_axis.delete = False
        dash.add_chart(ch, anchor)

    y0, y1 = pos['year']
    bar('연도별 외주인건비 (유상 제외)', Reference(agg, min_col=1, min_row=y0, max_row=y1),
        [(2, y0, y1, y0 - 1), (3, y0, y1, y0 - 1)], 'A7', stacked=True)
    m0, m1 = pos['month']
    bar(f'{ly}년 월별 외주인건비', Reference(agg, min_col=1, min_row=m0, max_row=m1),
        [(3, m0, m1, m0 - 1), (2, m0, m1, m0 - 1)], 'G7', stacked=True)
    ln = LineChart()
    ln.title, ln.height, ln.width = f'누적 추이 ({ly}년 vs {ly - 1}년)', 8.5, 17
    ln.add_data(Reference(agg, min_col=5, min_row=m0 - 1, max_row=m1), titles_from_data=True)
    ln.add_data(Reference(agg, min_col=6, min_row=m0 - 1, max_row=m1), titles_from_data=True)
    ln.set_categories(Reference(agg, min_col=1, min_row=m0, max_row=m1))
    ln.y_axis.numFmt = '#,##0,,"백만"'
    ln.y_axis.majorGridlines = None
    ln.x_axis.delete = ln.y_axis.delete = False
    ln.legend.position = 'b'
    dash.add_chart(ln, 'M7')
    for key, t, anchor in [('equip', '장비별 TOP 15', 'A25'), ('ship', '호선별 TOP 10', 'G25'), ('vendor', '업체별', 'M25')]:
        a, b = pos[key]
        if b >= a:
            bar(f'{ly}년 {t}', Reference(agg, min_col=1, min_row=a, max_row=b), [(2, a, b, a - 1)],
                anchor, horizontal=True, h=10)
    s0, s1 = pos['seg']
    pie = PieChart()
    pie.title, pie.height, pie.width = f'{ly}년 선종별 (유상 포함)', 10, 12
    pie.add_data(Reference(agg, min_col=2, min_row=s0 - 1, max_row=s1), titles_from_data=True)
    pie.set_categories(Reference(agg, min_col=1, min_row=s0, max_row=s1))
    pie.dataLabels = DataLabelList()
    pie.dataLabels.showPercent = True
    pie.legend.position = 'b'
    dash.add_chart(pie, 'S25')

    wb.calculation.fullCalcOnLoad = True
    out_dir = os.path.join(root, '엑셀보고서')
    os.makedirs(out_dir, exist_ok=True)
    path = os.path.join(out_dir, f'외주인건비_보고서_{ly}.{lm:02d}.xlsx')
    wb.save(path)
    return path
