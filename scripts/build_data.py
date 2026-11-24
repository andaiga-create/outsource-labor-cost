# -*- coding: utf-8 -*-
"""외주인건비 원본(xls) → docs/data.json 변환

엑셀 '월간 인건비 보고서'의 파워쿼리 규칙을 그대로 재현한다.

 1단계 (폴더별 쿼리)
   - legacy (~2024 폴더)   : 거래처 가 비어 있지 않은 행 (합계행 제거)
   - AS 2025~              : 거래처명 이 비어 있지 않고 품목번호 = CP 또는 CQ
   - 김해공장 2025~        : 품목번호 = CP 또는 CQ
 2단계 (보고서 RAW_통합)
   - 품목번호에 CP 또는 CQ 포함
   - 프로젝트코드에 'LWN-' 미포함 (프로젝트코드가 빈 행도 제외)
 파생열 (RAW_통합 의 엑셀 수식과 동일)
   - 고객 / 호선 / 장비 : 프로젝트코드를 '-' 로 나눈 값
   - 선종 : 코드 끝 2자리 C1·C2 → 상선, S1 → 특수선, 그 외 → 유상

열은 순서가 아니라 '이름'으로 찾으므로 열이 추가·이동돼도 동작한다.
"""
import json, os, sys, glob, re, datetime
from collections import defaultdict

try:
    import xlrd
except ImportError:
    sys.exit('[ERROR] xlrd 가 없습니다:  pip install xlrd')

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, 'docs', 'data.json')

# (폴더, 구분, 규칙)  — 새 폴더가 생기면 여기에 한 줄 추가
SOURCES = [
    ('AS/2024년까지',       'AS',     'legacy'),
    ('김해공장/2024년까지', '김해공장', 'legacy'),
    ('AS (2025~)',          'AS',     'as2025'),
    ('김해공장 (2025~)',    '김해공장', 'gh2025'),
]
REQUIRED = ['입고일자', '거래처명', '품목번호', '공급가', '프로젝트코드']
ALIASES = {'규격/재질': '규격'}


def cell(v):
    """빈 문자열 → None, 나머지는 그대로"""
    if v is None:
        return None
    if isinstance(v, str):
        return v if v != '' else None
    return v


def num(v):
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).replace(',', '').strip()
    try:
        return float(s)
    except ValueError:
        return None


def to_int(v):
    """파워쿼리 Int64.Type 변환과 같게: 반올림(짝수 쪽)"""
    x = num(v)
    return None if x is None else int(round(x))


def to_date(v, datemode):
    if v is None:
        return None
    if isinstance(v, float):
        return xlrd.xldate.xldate_as_datetime(v, datemode).date()
    m = re.match(r'\s*(\d{4})[/.-](\d{1,2})[/.-](\d{1,2})', str(v))
    return datetime.date(int(m[1]), int(m[2]), int(m[3])) if m else None


def read_file(path):
    book = xlrd.open_workbook(path, logfile=open(os.devnull, 'w'))
    names = book.sheet_names()
    if 'Sheet1' in names:
        sh = book.sheet_by_name('Sheet1')
    else:
        print(f'  [주의] Sheet1 없음 → 첫 시트 "{names[0]}" 사용: {path}')
        sh = book.sheet_by_index(0)
    head = [ALIASES.get(str(h).strip(), str(h).strip()) for h in sh.row_values(0)]
    miss = [c for c in REQUIRED if c not in head]
    if miss:
        raise SystemExit(f'[ERROR] 필수 열 {miss} 이 없습니다: {path}\n        현재 열: {head}')
    idx = {h: i for i, h in enumerate(head) if h and h not in ('',)}
    rows = []
    for r in range(1, sh.nrows):
        vals = sh.row_values(r)
        rows.append({h: cell(vals[i]) if i < len(vals) else None for h, i in idx.items()})
    return rows, book.datemode, head


def stage1(rows, rule):
    out = []
    for r in rows:
        pn = r.get('품목번호')
        if rule == 'legacy':
            ok = r.get('거래처') is not None
        elif rule == 'as2025':
            ok = r.get('거래처명') is not None and pn in ('CP', 'CQ')
        else:  # gh2025
            ok = pn in ('CP', 'CQ')
        if ok:
            out.append(r)
    return out


def stage2(r):
    pn, pc = r.get('품목번호'), r.get('프로젝트코드')
    if pn is None or pc is None:
        return False
    pn, pc = str(pn), str(pc)
    return ('CP' in pn or 'CQ' in pn) and 'LWN-' not in pc


def split_code(code):
    """RAW_통합 의 엑셀 수식 재현 (오류가 나는 경우 '' 반환)
       고객 = LEFT(I, FIND("-",I)-1)
       호선 = MID(I, FIND("-",I)+1, SEARCH("-",I,LEN(고객)+2)-(FIND("-",I)+1))
       장비 = MID(I, LEN(고객)+LEN(호선)+3, LEN(I)-LEN(고객)-LEN(호선)-5)"""
    p1 = code.find('-')
    if p1 < 0:
        return '', '', ''
    cust = code[:p1]
    p2 = code.find('-', p1 + 1)
    if p2 < 0:
        return cust, '', ''
    ship = code[p1 + 1:p2]
    n = len(code) - len(cust) - len(ship) - 5
    equip = code[len(cust) + len(ship) + 2:len(cust) + len(ship) + 2 + n] if n >= 0 else ''
    return cust, ship, equip


def segment(code):
    t = code[-2:]
    return '상선' if t in ('C1', 'C2') else '특수선' if t == 'S1' else '유상'


def main():
    records, dupkeys, nfiles = [], defaultdict(set), 0
    print('원본 읽는 중...')
    for folder, kind, rule in SOURCES:
        d = os.path.join(ROOT, folder)
        if not os.path.isdir(d):
            print(f'  [주의] 폴더 없음, 건너뜀: {folder}')
            continue
        files = sorted(f for f in glob.glob(os.path.join(d, '*.xls*'))
                       if not os.path.basename(f).startswith('~$'))
        cnt1 = cnt2 = 0
        for f in files:
            nfiles += 1
            rows, datemode, head = read_file(f)
            s1 = stage1(rows, rule)
            cnt1 += len(s1)
            fn = os.path.relpath(f, ROOT)
            for r in s1:
                if not stage2(r):
                    continue
                cnt2 += 1
                dt = to_date(r.get('입고일자'), datemode)
                if dt is None:
                    print(f'  [주의] 입고일자 인식 불가 → 제외: {fn} {r.get("입고일자")}')
                    continue
                code = str(r['프로젝트코드'])
                cust, ship, equip = split_code(code)
                amt = to_int(r.get('공급가')) or 0
                key = (r.get('전표번호'), r.get('순번'), r.get('품목번호'), code, amt)
                if r.get('전표번호') is not None:
                    dupkeys[key].add(fn)
                records.append(dict(y=dt.year, m=dt.month, d=dt.isoformat(), kind=kind,
                                    vendor=str(r.get('거래처명') or ''), cust=cust, ship=ship,
                                    equip=equip, seg=segment(code), amt=amt))
        print(f'  {folder:<18} 파일 {len(files):>3}개 · 1단계 {cnt1:>5}행 · 최종 {cnt2:>5}행')

    if not records:
        sys.exit('[ERROR] 읽은 데이터가 없습니다.')

    dups = {k: v for k, v in dupkeys.items() if len(v) > 1}
    if dups:
        print(f'\n  [주의] 서로 다른 파일에 같은 전표가 {len(dups)}건 있습니다 (중복 투입 의심):')
        for k, v in list(dups.items())[:10]:
            print(f'     전표 {k[0]} 순번 {k[1]} {k[3]} {k[4]:,}원 ← {", ".join(sorted(v))}')

    dims = ['연', '월', '구분', '업체', '고객', '호선', '장비', '선종']
    agg = defaultdict(lambda: [0, 0])
    for r in records:
        k = (r['y'], r['m'], r['kind'], r['vendor'], r['cust'], r['ship'], r['equip'], r['seg'])
        agg[k][0] += r['amt']
        agg[k][1] += 1
    latest = max(r['d'] for r in records)
    data = {
        'generated': datetime.datetime.now().strftime('%Y-%m-%d %H:%M'),
        'latest': latest[:7],
        'latest_date': latest,
        'files': nfiles,
        'records': len(records),
        'dims': dims,
        'rows': [list(k) + v for k, v in sorted(agg.items())],
    }
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    with open(OUT, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, separators=(',', ':'))

    # ---- 엑셀 보고서(피벗)와 대조용 요약 ----
    print(f'\n총 {len(records):,}행 (엑셀 RAW_통합 행 수와 같아야 함) · 최신 입고일 {latest}')
    print('\n연도별 외주인건비 (피벗_연도 = 유상 제외)')
    print(f'  {"연도":<6}{"유상 제외":>16}{"유상 포함":>16}')
    by = defaultdict(lambda: [0, 0])
    for r in records:
        by[r['y']][1] += r['amt']
        if r['seg'] != '유상':
            by[r['y']][0] += r['amt']
    for y in sorted(by):
        print(f'  {y:<6}{by[y][0]:>16,}{by[y][1]:>16,}')
    ly = max(by)
    print(f'\n{ly}년 월별 (피벗_월별, 유상 제외)')
    print(f'  {"월":<4}{"김해공장":>14}{"AS":>14}{"합계":>14}')
    bm = defaultdict(lambda: defaultdict(int))
    for r in records:
        if r['y'] == ly and r['seg'] != '유상':
            bm[r['m']][r['kind']] += r['amt']
    for m in sorted(bm):
        g, a = bm[m]['김해공장'], bm[m]['AS']
        print(f'  {m:02d}월{g:>14,}{a:>14,}{g + a:>14,}')
    print(f'\n→ {os.path.relpath(OUT, ROOT)} 저장 완료')


if __name__ == '__main__':
    main()
