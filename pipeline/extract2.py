import re, json, sys, datetime
sys.path.insert(0, __import__('os').path.dirname(__import__('os').path.abspath(__file__)))
import extract as X
from extract import CODE, clean, rows_of, num

# ---------- standard EXAT / DOH summary sheet, keeping escalation group and VAT ----------
def parse_std(f, sh, off=0, stop=('รวมค่าดำเนินการและบำรุงรักษา', 'รวมทั้งหมด', 'รวมค่าใช้จ่ายงาน Operation'), skip=('Start up',)):
    rows = rows_of(f, sh)
    for i, r in enumerate(rows):
        ys = [j for j, v in enumerate(r) if j >= off and isinstance(v, (int, float)) and 2500 < v < 2700]
        if len(ys) >= 30:
            hdr = i; ycols = ys[:30]; years = [int(r[j]) for j in ycols]; break
    items = []; group = None; sec = 'O'; lastcode = None
    for r in rows[hdr + 1:]:
        a = r[off]; b = r[off + 1]
        a = str(a).strip() if a not in (None, '', 0) else ''
        vals = [num(r[j]) / 1e6 for j in ycols]
        hasnum = any(isinstance(r[j], (int, float)) for j in ycols)
        if a.startswith('สมมติฐาน'): break
        if a.startswith('รวม'):
            if any(s in a for s in stop) or ('Operation' in a and 'Maintenance' in a): break
            sec = 'M'; continue
        if not hasnum or b in (None, ''):
            if a and not CODE.match(a): group = a
            continue
        if any(s in str(b) for s in skip): continue
        if CODE.match(a): lastcode = a.strip('() ').replace(' ', '')
        esc = str(r[off + 2] or '').strip(); vat = str(r[off + 3] or '').strip()
        part = 'สวัสดิการ' if 'สวัสดิการ' in str(b) else ('เงินเดือน' if 'เงินเดือน' in str(b) else '')
        ex = [it for it in items if it['code'] == lastcode and it['group'] == group]
        if ex:
            it = ex[0]; it['v'] = [x + y for x, y in zip(it['v'], vals)]
        else:
            it = dict(code=lastcode, name=clean(b), group=group, sec=sec, v=vals, esc=[], vat=[])
            items.append(it)
        if esc: it['esc'].append((part + ': ' if part else '') + esc)
        if vat: it['vat'].append(vat)
    # escalation legend: rows like  [.. 'อัตรา...' , 'กลุ่ม X', 'รายปี', rates...]
    legend = {}
    for r in rows:
        c = str(r[off + 2] or '').strip(); d = str(r[off + 3] or '').strip()
        if re.match(r'^กลุ่ม\s*[A-Z]$', c) and d == 'รายปี' and r[off + 1]:
            rates = [r[j] for j in ycols if isinstance(r[j], (int, float))]
            nz = sorted({round(x * 100, 2) for x in rates if x})
            legend[c] = dict(desc=str(r[off + 1]).strip(), rates=nz, n=sum(1 for x in rates if x), yrs=len(rates))
    for it in items:
        it['esc'] = ' / '.join(dict.fromkeys(it['esc']))
        it['vat'] = '/'.join(dict.fromkeys(v.upper() for v in it['vat']))
    return years, items, legend

def esc_text(esc, legend):
    out = []
    for part in esc.split(' / '):
        m = re.search(r'กลุ่ม\s*[A-Z]', part)
        if not m: out.append(part); continue
        g = legend.get(m.group(0).replace(' ', ' '))
        if not g: out.append(part); continue
        r = g['rates']
        if not r: rate = 'ไม่มีการปรับขึ้น'
        else:
            rr = f"{r[0]:g}%" if len(r) == 1 else f"{r[0]:g}–{r[-1]:g}%"
            rate = f"ปรับ {rr} ทุกปี" if g["n"] >= g["yrs"] - 2 else f"ปรับ {rr} จำนวน {g['n']} ครั้งใน 30 ปี"
        desc = re.sub(r'\s*\(กลุ่ม\s*[A-Z]\)', '', g['desc'])
        lab = (part.split(':')[0] + ': ') if ':' in part else ''
        out.append(f"{lab}{desc} ({m.group(0)}) {rate}")
    return ' / '.join(out)

# ---------- BSP: the model's own grouping (SUM sheet rows 33-41) ----------
BSP_GROUPS = [('ค่าใช้จ่ายบุคลากร', ('OS1', 'OS2', 'OS3', 'OD')), ('ค่าสาธารณูปโภค', ('OU1', 'OU2', 'OU3', 'OU4')),
              ('ค่าไฟฟ้าขบวนรถไฟฟ้า', ('OR',)), ('ค่าทักษะความเชี่ยวชาญและค่าใช้จ่ายอื่น ๆ', ('OA',)),
              ('ค่าบำรุงรักษาขบวนรถไฟฟ้า', ('MR',)), ('ค่าบำรุงรักษาชิ้นส่วนอุปกรณ์หลักของระบบการเดินรถ', tuple(f'MS{i}' for i in range(1, 12))),
              ('ค่าบำรุงรักษางานโครงสร้างพื้นฐาน', ('MI',))]
def parse_bsp(sh):
    years, items = X.parse_bsp(sh)
    for it in items:
        it['group'] = next(g for g, cs in BSP_GROUPS if it['code'] in cs)
    return years, items

# ---------- GLP: no codes in the model; keep its own names, grouped as Total Oper / Total Maintenance ----------
def parse_glp(f):
    _, items = X.parse_glp(f)
    for it in items:
        it['code'] = ''
        it['group'] = 'Total Oper' if it['sec'] == 'O' else 'Total Maintenance'
    return None, items

# ---------- assumption sheets: every row as written, with code where the row has one ----------
ACODE = re.compile(r'^\(?([A-Z]{1,3})\s?(\d+(?:\.\d+)*)\)?$')
def parse_assumptions(f, sheets, maxcol=12):
    out = []
    for sh in sheets:
        rows = rows_of(f, sh, 700)
        section = ''; code = ''
        for r in rows:
            cells = [c for c in r[:maxcol]]
            texts = [(j, str(c).strip()) for j, c in enumerate(cells) if isinstance(c, str) and str(c).strip()]
            nums = [(j, c) for j, c in enumerate(cells) if isinstance(c, (int, float)) and not isinstance(c, bool)]
            if not texts and not nums: continue
            rowcode = ''
            for j, s in texts[:2]:
                m = ACODE.match(s)
                if m: rowcode = m.group(1) + m.group(2)
            labels = [s for j, s in texts if not ACODE.match(s)]
            if not labels or labels[0] in ('ลำดับที่', 'ลำดับ', 'Code', 'รายการ', 'รหัส'): continue
            if not nums and len(labels) == 1 and (r[0] not in (None, '') or rowcode):
                if rowcode: code = rowcode
                section = labels[0]; continue
            if rowcode: code = rowcode
            label = labels[0]
            first_label_col = next(j for j, s in texts if s == label)
            vals = [c for j, c in nums if j > first_label_col]
            value = vals[0] if vals else None
            rest = [s for s in labels[1:]]
            out.append(dict(sheet=sh, section=section, code=rowcode or (code if code and section else ''), item=label[:300],
                            value=(round(value, 6) if isinstance(value, float) else value), unit=(rest[0][:120] if rest else ''),
                            note=(' · '.join(rest[1:])[:300] if len(rest) > 1 else '')))
    return out

# ---------- catalogue ----------
CASE = {'PSC': 'PSC', 'PPP': 'PPP', 'PPP (Private)': 'PPP', 'PPP (State)': 'PPP (รัฐ)', 'PPP (Gross Cost)': 'PPP (Gross Cost)', 'PPP (Net Cost)': 'PPP (Net Cost)'}
def build():
    P = []   # (project, stage, case, years, items, legend, file)
    y, it, lg = parse_std('M82.xlsx', 'Summary_O&M'); P.append(('M82', 'RFP', 'PPP', y, it, lg, 'M82_Model_O_M_Cost_HDM4_Inflation_2_031125_Case_1_Final.xlsx'))
    for k, op, m in [('PSC', '1. Sum Operation_PSC', '1_Sum_PSC (M Flow Adjust)'), ('PPP (Gross Cost)', '2. Sum Operation PPP_Gross', '2_Sum_PPP (M Flow Adjust)'),
                     ('PPP (Net Cost)', '3. Sum Operation PPP_Net', '2_Sum_PPP (M Flow Adjust)'), ('PPP (รัฐ)', '4. Sum Operation PPP_รัฐ', '3_Sum_PPP_รัฐ')]:
        y, it = X.parse_rpm(op, m); P.append(('M82', 'SEPO', k, y, it, {}, 'RPM_Cost_Structure_M-Flow__68__030722_CPI_2_5.xlsx'))
    for k, s in [('PSC', 'Summary_O&M_PSC'), ('PPP', 'Summary_O&M_PPP (เอกชน)'), ('PPP (รัฐ)', 'Summary_O&M_PPP (EXAT)')]:
        y, it, lg = parse_std('PPK.xlsx', s); P.append(('PPK', 'SEPO', k, y, it, lg, '6_PPK_Model_O_M_Cost_091526_Leadline_insurance_New_BOQ_Vat.xlsx'))
    for k, off in [('PSC', 0), ('PPP', 40)]:
        y, it, lg = parse_std('SSP.xlsx', 'Summary_O&M', off); P.append(('SSP', 'SEPO', k, y, it, lg, 'SSP_Model_O_M_Cost_090926_Vat.xlsx'))
    for k, s in [('PSC', 'Summary_O&M_PSC'), ('PPP', 'Summary_O&M_PPP (เอกชน)'), ('PPP (รัฐ)', 'Summary_O&M_PPP (รัฐ)')]:
        y, it, lg = parse_std('PPK_RA.xlsx', s); P.append(('PPK_RA', 'SEPO', k, y, it, lg, 'PPK_RA_Model_O_M_Cost_062025.xlsx'))
    for k, s in [('PSC', 'Sum PSC (รวม)'), ('PPP', 'Sum PPP (รวม)')]:
        y, it = X.parse_pbk(s); P.append(('PBK', 'SEPO', k, y, it, {}, 'PBK_RA_Expense_Revenue_M81_121321.xlsx'))
    for k, s in [('PSC', 'SUM_PSC'), ('PPP', 'SUM_PPP'), ('PPP (รัฐ)', 'SUM_PPP รัฐ')]:
        y, it = parse_bsp(s); P.append(('BSP', 'SEPO', k, y, it, {}, 'BSP_OperationMaintenance_Model_CPI2.5_040222.xlsx'))
    for k, f in [('PSC', 'GLP_PSC.xlsx'), ('PPP', 'GLP_PPP.xlsx')]:
        y, it = parse_glp(f); P.append(('GLP', 'SEPO', k, y, it, {}, f'GLP_Summary_Cost_{k}_CPI2.5_101222.xlsx'))

    items = []; n = 0
    for proj, stage, case, yrs, its, lg, src in P:
        for seq, i in enumerate(its):
            r = dict(item_id=f'{proj}-{stage}-{n:04d}', project=proj, stage=stage, scenario=case, seq=seq,
                     sec='Operation' if i['sec'] == 'O' else 'Maintenance', grp=i['group'] or ('Operation' if i['sec'] == 'O' else 'Maintenance'),
                     code=i['code'] or '', item=i['name'], escalation=esc_text(i.get('esc', ''), lg) if i.get('esc') else '',
                     vat=i.get('vat', ''), first_year=yrs[0] if yrs else None, source_file=src,
                     total_30y=round(sum(i['v']), 4))
            for k, x in enumerate(i['v'], 1): r[f'y{k:02d}'] = round(x, 4)
            items.append(r); n += 1

    # assumptions
    A = []
    for proj, stage, f, sheets in [('M82', 'RFP', 'M82.xlsx', ['Assumption_O', 'Assumption_M']), ('M82', 'SEPO', 'RPM.xlsx', ['Fact_Operation', 'Assumption_Maintenance']),
                                   ('PPK', 'SEPO', 'PPK.xlsx', ['Assumption_Operation', 'Assumption_Maintenance']), ('SSP', 'SEPO', 'SSP.xlsx', ['Assumption_O', 'Assumption_M']),
                                   ('PPK_RA', 'SEPO', 'PPK_RA.xlsx', ['Assumption_Operation', 'Assumption_Maintenance']), ('PBK', 'SEPO', 'PBK.xlsx', ['Assumption Expense']),
                                   ('BSP', 'SEPO', 'BSP.xlsx', ['Assumption'])]:
        for k, a in enumerate(parse_assumptions(f, sheets, 6 if proj == 'BSP' else 12)):
            A.append(dict(row_id=f'{proj}-{stage}-{k:04d}', project=proj, stage=stage, **a))
    return items, A

if __name__ == '__main__':
    items, A = build()
    json.dump(items, open('/home/claude/work/up/om_items2.json', 'w'), ensure_ascii=False, separators=(',', ':'))
    json.dump(A, open('/home/claude/work/up/assumptions.json', 'w'), ensure_ascii=False, separators=(',', ':'))
    from collections import defaultdict
    t = defaultdict(lambda: [0, 0, 0])
    for r in items:
        k = (r['project'], r['stage'], r['scenario']); t[k][0 if r['sec'] == 'Operation' else 1] += r['total_30y']; t[k][2] += 1
    for k, v in t.items(): print(k, round(v[0], 2), round(v[1], 2), round(v[0] + v[1], 2), 'items', v[2])
    print('assumption rows', len(A), {p: sum(1 for a in A if a['project'] == p) for p in dict.fromkeys(a['project'] for a in A)})
    print('with code', sum(1 for a in A if a['code']))
