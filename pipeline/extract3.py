import re, json, sys
sys.path.insert(0, __import__('os').path.dirname(__import__('os').path.abspath(__file__)))
import extract as X
import extract2 as E
from extract import rows_of, wb

OUT = '/home/claude/work/up/'

# ---------- M82 SEPO (RPM): regroup the 18 model groups into the same 6 groups the RFP model uses ----------
RPM_GROUP = {
    'O_personnel': 'การดำเนินงานด้านบุคลากร', 'O_utility': 'การดำเนินงานด้านสาธารณูปโภค', 'O_admin': 'การดำเนินงานด้านบริหารจัดการ',
    'M_system': 'การดำเนินงานด้านงานระบบ', 'M_civil': 'การดำเนินงานด้านงานโยธา', 'M_building': 'การดำเนินงานด้านงานอาคาร',
}
def rpm_regroup(it):
    c = it['code']; n = it['name']
    if it['sec'] == 'O':
        if c.startswith('OU'): return RPM_GROUP['O_utility']
        if re.search(r'บุคลากร|เบี้ยเลี้ยง|ผู้บริหาร', n): return RPM_GROUP['O_personnel']
        return RPM_GROUP['O_admin']
    if c.startswith(('MT', 'MW', 'MI')) or c == 'MB2': return RPM_GROUP['M_system']
    if c == 'MB1': return RPM_GROUP['M_building']
    return RPM_GROUP['M_civil']

# ---------- rest-area models (TRA, RAP): the model report's own cost headings, by code prefix ----------
RA_GROUP = [('OM', 'ค่าใช้จ่ายบุคลากรบริหารโครงการ (Management Cost)'), ('OH', 'ค่าใช้จ่ายบุคลากรบริหารโครงการ (Management Cost)'),
            ('OD', 'ค่าใช้จ่ายทางตรง (Direct Cost)'), ('OI', 'ค่าใช้จ่ายจิปาถะ (Indirect Cost)'),
            ('MR', 'ค่าใช้จ่ายในการบำรุงรักษาประจำปี (Routine Maintenance)'), ('MP', 'ค่าบูรณะสิ่งปลูกสร้าง (Periodic Maintenance)')]
def ra_group(code):
    return next((g for p, g in RA_GROUP if code.startswith(p)), None)

def sum_sheets(f, sheets, off):
    """Parse several summary sheets with the same layout and add them up by code."""
    acc = {}; order = []; years = None; legend = {}
    for sh in sheets:
        y, its, lg = E.parse_std(f, sh, off)
        years = years or y; legend.update(lg)
        for it in its:
            k = (it['code'], it['sec'])
            if k not in acc: acc[k] = dict(it); acc[k]['v'] = list(it['v']); order.append(k)
            else: acc[k]['v'] = [a + b for a, b in zip(acc[k]['v'], it['v'])]
    items = [acc[k] for k in order]
    for it in items: it['group'] = ra_group(it['code']) or it['group']
    return years, items, legend

# ---------- per-item calculation detail: the calculation block on each item's own sheet ----------
def sheet_for(sheetnames, code):
    code = code.replace(' ', '')
    if not code: return []
    hits = []
    for s in sheetnames:
        base = re.split(r'[_ ]', s.strip())[0].replace(' ', '')
        if base == code or (base.startswith(code + '.') and not code.count('.') >= base.count('.')):
            hits.append(s)
    return hits[:3]

def fmtcell(v):
    if isinstance(v, bool) or v is None: return ''
    if isinstance(v, (int, float)):
        if abs(v) >= 1000: return f'{v:,.0f}'
        if v == int(v): return str(int(v))
        return f'{v:,.4f}'.rstrip('0').rstrip('.')
    s = str(v).strip().replace('\n', ' ')
    return s[:70]

def calc_block(f, sh, maxrows=70, maxcols=20):
    rows = rows_of(f, sh, 200)
    if not rows: return []
    ncol = max(len(r) for r in rows)
    # columns that hold a value in most rows (the yearly series on the left of the sheet)
    yearly = set()
    for c in range(ncol):
        nn = sum(1 for r in rows[1:40] if c < len(r) and isinstance(r[c], (int, float)))
        if nn >= 18: yearly.add(c)
    start = None
    for r in rows[1:]:
        for c, v in enumerate(r):
            if c < 2 or c in yearly: continue
            if isinstance(v, str) and v.strip() and not re.match(r'^\(?[A-Z]{1,3}\d', v.strip()):
                start = c if start is None else min(start, c)
                break
    if start is None: return []
    out = []
    for r in rows:
        cells = [fmtcell(v) for v in r[start:start + maxcols]]
        while cells and cells[-1] == '': cells.pop()
        if any(cells): out.append(cells)
        if len(out) >= maxrows: break
    return out

def item_calcs(proj, stage, f, codes):
    names = wb(f).sheetnames
    out = []
    for code in codes:
        for sh in sheet_for(names, code):
            try: blk = calc_block(f, sh)
            except Exception as e: blk = []
            if blk: out.append(dict(key=f'{proj}|{stage}|{code}|{sh}', project=proj, stage=stage, code=code, sheet=sh, rows=blk))
    return out

def parse_et1(sheet='SUM PPP', nyears=30, first=2572):
    rows = rows_of('ET1.xlsx', sheet, 40)
    hdr = rows[0]
    cols = [j for j, v in enumerate(hdr) if isinstance(v, int) and first <= v < first + nyears]
    years = [hdr[j] for j in cols]
    items = []; group = None
    for r in rows[1:]:
        a = str(r[0] or '').strip()
        if not a or a.startswith('รวม'): continue
        m = re.match(r'^(\d)\.\s*(.+)$', a)
        if m and not re.match(r'^\d\.\d', a): group = m.group(2).strip(); continue
        m = re.match(r'^(\d\.\d)\s+(.+)$', a)
        if not m: continue
        name = re.sub(r'\s*กรณี (PSC|PPP)( ปีที่ 1)?\s*$', '', m.group(2)).strip()
        sec = 'M' if m.group(1).startswith('3') else 'O'
        items.append(dict(code=m.group(1), name=name, group=group, sec=sec, v=[(r[j] or 0) / 1e6 if isinstance(r[j], (int, float)) else 0.0 for j in cols]))
    return years, items

def build(with_calcs=True):
    items, A = E.build()
    # M82 SEPO regrouping
    rpm_items = {}
    for k, op, m in [('PSC', '1. Sum Operation_PSC', '1_Sum_PSC (M Flow Adjust)'), ('PPP (Gross Cost)', '2. Sum Operation PPP_Gross', '2_Sum_PPP (M Flow Adjust)'),
                     ('PPP (Net Cost)', '3. Sum Operation PPP_Net', '2_Sum_PPP (M Flow Adjust)'), ('PPP (รัฐ)', '4. Sum Operation PPP_รัฐ', '3_Sum_PPP_รัฐ')]:
        pass
    for r in items:
        r['grp_model'] = r['grp']
        if r['project'] == 'M82' and r['stage'] == 'SEPO':
            r['grp'] = rpm_regroup(dict(code=r['code'], name=r['item'], sec='O' if r['sec'] == 'Operation' else 'M'))
        r['nyears'] = 30

    n = len(items)
    def add(proj, stage, case, yrs, its, lg, src, nyears=30):
        nonlocal n
        for seq, i in enumerate(its):
            v = i['v'][:nyears]
            r = dict(item_id=f'{proj}-{stage}-{n:04d}', project=proj, stage=stage, scenario=case, seq=seq,
                     sec='Operation' if i['sec'] == 'O' else 'Maintenance', grp=i['group'], grp_model=i['group'],
                     code=i['code'] or '', item=i['name'], escalation=E.esc_text(i.get('esc', ''), lg) if i.get('esc') else '',
                     vat=i.get('vat', ''), first_year=yrs[0], source_file=src, total_30y=round(sum(v), 4), nyears=nyears)
            for k in range(30): r[f'y{k + 1:02d}'] = round(v[k], 4) if k < nyears else None
            items.append(r); n += 1
    # TRA: two rest stops (Nam Phong, Tha Sae), phases 1 and 2, added up
    # each site kept apart (phases 1 and 2 added up), as cases like the RAP contracts
    for case, off in [('PSC', 0), ('PPP', 38)]:
        for site, k in [('น้ำพอง', 'NP'), ('ท่าแซะ', 'TS')]:
            y, its, lg = sum_sheets('TRA.xlsx', [f'Summary_O_M_{k}_P1', f'Summary_O_M_{k}_P2'], off)
            add('TRA', 'SEPO', case + ' ' + site, y, its, lg, 'TRA_RA_Model_O_M_Cost_081926.xlsx')
    # RAP: two contracts (C2.1, C2.2), 12-year contract period
    for case, sh, off in [('PSC สัญญาที่ 1', 'Summary_O_M_C2.1', 0), ('PSC สัญญาที่ 2', 'Summary_O_M_C2.2', 0),
                          ('PPP สัญญาที่ 1', 'Summary_O_M_C2.1', 38), ('PPP สัญญาที่ 2', 'Summary_O_M_C2.2', 38)]:
        y, its, lg = sum_sheets('RAP.xlsx', [sh], off)
        add('RAP', 'SEPO', case, y, its, lg, 'RAP_Model_O_M_Cost_DOH_12_Years_5_Area_051326.xlsx', 12)

    # ET1: Double Deck expressway, PPP case only (the PSC summary sheet is empty), 30 years 2572-2601 as the model total
    y, its = parse_et1()
    add('ET1', 'SEPO', 'PPP', y, its, {}, 'EXAT_DD_MTC_ETC_OM_Cost_Cap_Salary_RD_Welfare_050124_Intregation.xlsx', 30)

    # assumption sheets for the new models
    for proj, f, sheets in [('TRA', 'TRA.xlsx', ['Assumption_O', 'Assumption_M']), ('RAP', 'RAP.xlsx', ['Assumption_O', 'Assumption_M']), ('ET1', 'ET1.xlsx', ['Assumption_Data'])]:
        for k, a in enumerate(E.parse_assumptions(f, sheets)):
            A.append(dict(row_id=f'{proj}-SEPO-{k:04d}', project=proj, stage='SEPO', **a))

    if not with_calcs: return items, A, []
    # calculation details per item code
    files = {('M82', 'RFP'): 'M82.xlsx', ('M82', 'SEPO'): 'RPM.xlsx', ('PPK', 'SEPO'): 'PPK.xlsx', ('SSP', 'SEPO'): 'SSP.xlsx',
             ('PPK_RA', 'SEPO'): 'PPK_RA.xlsx', ('PBK', 'SEPO'): 'PBK.xlsx', ('BSP', 'SEPO'): 'BSP.xlsx', ('TRA', 'SEPO'): 'TRA.xlsx', ('RAP', 'SEPO'): 'RAP.xlsx', ('ET1', 'SEPO'): 'ET1.xlsx'}
    calcs = []
    for (proj, stage), f in files.items():
        codes = list(dict.fromkeys(r['code'] for r in items if r['project'] == proj and r['stage'] == stage and r['code']))
        c = item_calcs(proj, stage, f, codes)
        print(proj, stage, 'codes', len(codes), 'with calc sheet', len({x['code'] for x in c}))
        calcs += c
    return items, A, calcs

if __name__ == '__main__':
    items, A, calcs = build()
    json.dump(items, open(OUT + 'om_items3.json', 'w'), ensure_ascii=False, separators=(',', ':'))
    json.dump(A, open(OUT + 'assumptions3.json', 'w'), ensure_ascii=False, separators=(',', ':'))
    json.dump(calcs, open(OUT + 'item_calcs.json', 'w'), ensure_ascii=False, separators=(',', ':'))
    from collections import defaultdict
    t = defaultdict(lambda: [0, 0, 0])
    for r in items:
        k = (r['project'], r['stage'], r['scenario']); t[k][0 if r['sec'] == 'Operation' else 1] += r['total_30y']; t[k][2] += 1
    for k, v in t.items():
        if k[0] in ('TRA', 'RAP', 'ET1'): print(k, round(v[0], 2), round(v[1], 2), round(v[0] + v[1], 2), 'items', v[2])
    import os
    for fn in ['om_items3.json', 'assumptions3.json', 'item_calcs.json']: print(fn, os.path.getsize(OUT + fn))
