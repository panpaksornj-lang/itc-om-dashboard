import openpyxl, re, json, sys
sys.path.insert(0, __import__('os').path.dirname(__import__('os').path.abspath(__file__)))
import os
D = os.environ.get('MODEL_DIR', '/home/claude/data/').rstrip('/') + '/'
CODE = re.compile(r'^\(?\s*[A-Z]{1,3}\s?\d')

def clean(n):
    n = str(n).strip()
    n = re.sub(r'\s*\((เงินเดือน[^)]*|สวัสดิการ)\)\s*$', '', n)
    n = re.sub(r'\s*\((M[SCB]\d[\d.]*|OM\d+)\)', '', n)
    return re.sub(r'\s+', ' ', n).strip()

_wb = {}
def wb(f):
    if f not in _wb:
        _wb[f] = openpyxl.load_workbook(D + f, read_only=True, data_only=True)
    return _wb[f]

def rows_of(f, sh, maxr=400):
    return list(wb(f)[sh].iter_rows(max_row=maxr, values_only=True))

def num(v):
    return float(v) if isinstance(v, (int, float)) and not isinstance(v, bool) else 0.0

# ---------- EXAT / DOH standard summary (A code, B item, years E..AH) ----------
def parse_std(f, sh, off=0, stop=('รวมค่าดำเนินการและบำรุงรักษา', 'รวมทั้งหมด', 'รวมค่าใช้จ่ายงาน Operation'), scale=1e6, skip=('Start up',)):
    rows = rows_of(f, sh)
    for i, r in enumerate(rows):
        ys = [j for j, v in enumerate(r) if j >= off and isinstance(v, (int, float)) and 2500 < v < 2700]
        if len(ys) >= 30:
            hdr = i; ycols = ys[:30]; years = [int(r[j]) for j in ycols]; break
    items = []; group = None; sec = 'O'; lastcode = None
    for r in rows[hdr + 1:]:
        a = r[off]; b = r[off + 1]
        a = str(a).strip() if a not in (None, '', 0) else ''
        vals = [num(r[j]) / scale for j in ycols]
        hasnum = any(isinstance(r[j], (int, float)) for j in ycols)
        if a.startswith('รวม'):
            if any(s in a for s in stop): break
            sec = 'M'
            continue
        if not hasnum or b in (None, ''):
            if a and not CODE.match(a): group = a
            continue
        if any(s in str(b) for s in skip): continue
        if CODE.match(a): lastcode = a.strip('() ').replace(' ', '')
        nm = clean(b)
        ex = [it for it in items if it['code'] == lastcode and it['group'] == group]
        if ex: ex[0]['v'] = [x + y for x, y in zip(ex[0]['v'], vals)]
        else: items.append(dict(code=lastcode, name=nm, group=group, sec=sec, v=vals))
    return years, items

# ---------- BSP (code+name in col A, years from col B) ----------
def parse_bsp(sh):
    rows = rows_of('BSP.xlsx', sh, 32)
    years = [int(v) for v in rows[0][1:31]]
    items = []; sec = 'O'
    grp = {'O': 'Operation Cost', 'M': 'Maintenance Cost'}
    for r in rows[3:]:
        a = str(r[0] or '').strip()
        if not a: continue
        if a.startswith('รวม'):
            if 'Operation' in a: sec = 'M'; continue
            break
        if a.startswith('Mainten'): sec = 'M'; continue
        m = re.match(r'^([A-Z]{1,3}\d*)\s+(.*)$', a)
        if not m: continue
        items.append(dict(code=m.group(1), name=clean(m.group(2)), group=grp[sec], sec=sec,
                          v=[num(x) / 1e6 for x in r[1:31]]))
    return years, items

# ---------- GLP (Grey Line monorail, single sheet, million THB, years 0..29 = G..AJ) ----------
GLP_O = {45: 'OS1', 46: 'OS2', 47: 'OS3', 48: 'OD', 49: 'OU1', 50: 'OU2', 51: 'OU3', 52: 'OU4', 53: 'OR', 54: 'OA'}
GLP_M = {55: 'MR1', 56: 'MR2', 57: 'MR3', 58: 'MS', 59: 'MR', 60: 'MI'}
def parse_glp(f):
    rows = rows_of(f, 'วัชรพล', 62)
    items = []
    for rn, code in {**GLP_O, **GLP_M}.items():
        r = rows[rn - 1]
        sec = 'O' if rn in GLP_O else 'M'
        items.append(dict(code=code, name=str(r[1]).strip().replace('Stick', 'Stock').replace('Builing', 'Building').replace('Overhall', 'Overhaul'),
                          group='Operation Cost' if sec == 'O' else 'Maintenance Cost', sec=sec,
                          v=[num(x) for x in r[6:36]]))
    return None, items

# ---------- PBK (Rest Areas M81; A code, B item, years G..AJ = 2567..2596) ----------
def parse_pbk(sh):
    rows = rows_of('PBK.xlsx', sh, 36)
    years = [int(v) for v in rows[0][6:36]]
    items = []; group = None
    for r in rows[2:35]:
        a = str(r[0] or '').strip(); b = str(r[1] or '').strip()
        if not a:
            if b and re.match(r'^\d\.', b): group = re.sub(r'^\d\.\s*', '', b)
            continue
        if not CODE.match(a): continue
        if a in ('OD2',): continue  # header row "ค่าใช้จ่ายอื่นๆ"
        sec = 'M' if a.startswith('M') else 'O'
        items.append(dict(code=a, name=clean(b), group=group, sec=sec, v=[num(x) / 1e6 for x in r[6:36]]))
    return years, items

# ---------- RPM (M-Flow cost structure) ----------
def rpm_years(hdr):
    return {j: 2500 + int(v) for j, v in enumerate(hdr) if isinstance(v, (int, float)) and 50 < v < 120}

def parse_rpm(op_sh, m_sh, y0=2568):
    yrs = list(range(y0, y0 + 30))
    items = []
    # operation (baht)
    rows = rows_of('RPM.xlsx', op_sh, 90)
    ym = rpm_years(rows[0]); inv = {v: k for k, v in ym.items()}
    started = False; group = None
    for r in rows[1:]:
        a = str(r[0] or '').strip(); b = str(r[1] or '').strip()
        if a == 'O': started = True; continue
        if not started: continue
        if a.startswith('รวม'): break
        if re.match(r'^O[A-Z]$', a): group = re.sub(r'\s*@.*$', '', b).strip(); continue
        if re.match(r'^O[A-Z]+\s?\d+', a):
            items.append(dict(code=a.replace(' ', ''), name=clean(b), group=group, sec='O',
                              v=[num(r[inv[y]]) / 1e6 if y in inv else 0.0 for y in yrs]))
    # maintenance (million baht)
    rows = rows_of('RPM.xlsx', m_sh, 60)
    ym = rpm_years(rows[0]); inv = {v: k for k, v in ym.items()}
    group = None
    for r in rows[2:]:
        a = r[0]; b = str(r[1] or '').strip()
        if isinstance(a, (int, float)): group = re.sub(r'\s*@.*$', '', b).strip(); continue
        a = str(a or '').strip()
        if not a and not b: break
        if re.match(r'^M[A-Z]+\s?\d+', a):
            items.append(dict(code=a.replace(' ', ''), name=clean(b), group=group, sec='M',
                              v=[num(r[inv[y]]) if y in inv else 0.0 for y in yrs]))
        elif b.startswith('ต้นทุน'): break
    return yrs, items

# ---------- SEPO harmonised categories ----------
CATS = [
    ('O1', 'Personnel', 'ค่าบุคลากร'),
    ('O2', 'Utilities & Energy', 'ค่าสาธารณูปโภคและพลังงาน'),
    ('O3', 'Management, Admin & Services', 'ค่าบริหารจัดการและจ้างเหมาบริการ'),
    ('M1', 'Systems & M&E Maintenance', 'บำรุงรักษางานระบบ'),
    ('M2', 'Civil, Pavement & Structure', 'บำรุงรักษางานโยธา ผิวทาง และโครงสร้าง'),
    ('M3', 'Building & Facilities', 'บำรุงรักษางานอาคาร'),
    ('M4', 'Major Renewal & Refurbishment', 'บูรณะ/เปลี่ยนทดแทนครั้งใหญ่'),
]
def categorise(it, ptype):
    n = it['name']; c = it['code'] or ''; nl = n.lower()
    if it['sec'] == 'O':
        if re.search(r'ค่าไฟ|ไฟฟ้า|น้ำประปา|ค่าน้ำ|โทรศัพท์|อินเตอร์|ไปรษณีย์|ขยะ|สิ่งปฏิกูล|electric|water|telephone|postage|utility|วัสดุสิ้นเปลือง', nl): return 'O2'
        if re.search(r'ภาษี|ประกัน|กำกับดูแลสัญญา|นายหน้า|ซ่อมบำรุง ?$|อาคารสถานที่|ค่าใช้จ่ายสำนักงาน', nl): return 'O3'
        if re.search(r'บุคลากร|staff|เงินเดือน|พนักงาน|ผู้บริหาร|ฝ่ายงาน|แผนก|เบี้ยเลี้ยง|รักษาความปลอดภัย|ทำความสะอาด|ดูแลสวน', nl): return 'O1'
        if c.startswith(('OD', 'OS')): return 'O1'
        return 'O3'
    if re.search(r'บูรณะ|refurb|overhaul|replacement of m&e', nl): return 'M4'
    if re.search(r'อาคาร|building|ประจำปี', nl): return 'M3'
    if ptype == 'Railway':
        return 'M2' if re.search(r'infrastructure|โครงสร้าง', nl) else 'M1'
    if re.search(r'ระบบ(?!ระบาย|บำบัด)|software|system|ไฟฟ้า|generator|ตู้ไฟ|สายไฟ|maintenance online', nl) or c.startswith(('MT', 'MI', 'MW')):
        return 'M1'
    return 'M2'

# ---------- project catalogue ----------
P = []
def add(pid, name, nameTh, ptype, owner, src, specs, scen, note=''):
    P.append(dict(id=pid, name=name, nameTh=nameTh, type=ptype, owner=owner, source=src, specs=specs, scen=scen, note=note))

def build():
    # Motorway
    y, it = parse_std('M82.xlsx', 'Summary_O&M')
    add('M82', 'M82 Motorway (Bang Khun Thian–Ban Phaeo)', 'ทางหลวงพิเศษหมายเลข 82 สายบางขุนเทียน–บ้านแพ้ว', 'Motorway', 'DOH',
        'M82_Model_O_M_Cost_HDM4_Inflation_2_031125_Case_1_Final.xlsx',
        dict(distance=24.7, tolls=6, lanes=6, opening=y[0], extra='Ramps 2 lanes; 5 TCB buildings'),
        {'PSC': (y, it)}, 'Case 1 Final (HDM-4, inflation 2%). Single cost case in model.')
    sc = {}
    for k, op, m in [('PSC', '1. Sum Operation_PSC', '1_Sum_PSC (M Flow Adjust)'),
                     ('PPP (Gross Cost)', '2. Sum Operation PPP_Gross', '2_Sum_PPP (M Flow Adjust)'),
                     ('PPP (Net Cost)', '3. Sum Operation PPP_Net', '2_Sum_PPP (M Flow Adjust)'),
                     ('PPP (State)', '4. Sum Operation PPP_รัฐ', '3_Sum_PPP_รัฐ')]:
        sc[k] = parse_rpm(op, m)
    add('RPM', 'RPM M-Flow Toll System (Rama II, M82)', 'ระบบเก็บค่าผ่านทาง M-Flow ถนนพระราม 2 (M82)', 'Motorway', 'DOH',
        'RPM_Cost_Structure_M-Flow__68__030722_CPI_2_5.xlsx',
        dict(distance=24.69, tolls=6, lanes=6, opening=2568, extra='M-Flow Adjust case, CPI 2.5%'),
        sc, 'M-Flow Adjust case. PPP Net/Gross share maintenance sheet.')
    # Expressway
    sc = {k: parse_std('PPK.xlsx', s) for k, s in [('PSC', 'Summary_O&M_PSC'), ('PPP (Private)', 'Summary_O&M_PPP (เอกชน)'), ('PPP (State)', 'Summary_O&M_PPP (EXAT)')]}
    add('PPK', 'Phuket Expressway (Kathu–Patong & Mueang Mai–Ko Kaeo)', 'ทางพิเศษสายกะทู้–ป่าตอง และเมืองใหม่–เกาะแก้ว', 'Expressway', 'EXAT',
        '6_PPK_Model_O_M_Cost_091526_Leadline_insurance_New_BOQ_Vat.xlsx',
        dict(distance=34.6, tolls=8, lanes=4, opening=sc['PSC'][0][0], extra='Ramp 2 lanes; Phase 1 3.98 km + Phase 2 30.62 km'),
        sc)
    sc = {'PSC': parse_std('SSP.xlsx', 'Summary_O&M', 0), 'PPP': parse_std('SSP.xlsx', 'Summary_O&M', 40)}
    add('SSP', 'SSP Expressway', 'ทางพิเศษ SSP', 'Expressway', 'EXAT',
        'SSP_Model_O_M_Cost_090926_Vat.xlsx',
        dict(distance=15.8, tolls=2, lanes=4, opening=sc['PSC'][0][0], extra='2 TSB buildings, 2 sub-rescue'),
        sc, 'Start-up cost (4.39 MB) excluded, as in the model total.')
    # Rest area
    sc = {k: parse_std('PPK_RA.xlsx', s) for k, s in [('PSC', 'Summary_O&M_PSC'), ('PPP (Private)', 'Summary_O&M_PPP (เอกชน)'), ('PPP (State)', 'Summary_O&M_PPP (รัฐ)')]}
    add('PPK_RA', 'Phuket Expressway Rest Area', 'ที่พักริมทางพิเศษภูเก็ต', 'Rest Area', 'EXAT',
        'PPK_RA_Model_O_M_Cost_062025.xlsx',
        dict(distance=34.6, tolls=8, lanes=4, opening=sc['PSC'][0][0], sites=1, area_rai=37, extra='1 service area, 37 rai, incl. fuel station'),
        sc, 'Host route: Phuket Expressway (34.6 km, 4 lanes, 8 toll plazas).')
    sc = {'PSC': parse_pbk('Sum PSC (รวม)'), 'PPP': parse_pbk('Sum PPP (รวม)')}
    add('PBK', 'M81 Rest Areas (Bang Yai–Kanchanaburi)', 'ที่พักริมทาง M81 บางใหญ่–กาญจนบุรี', 'Rest Area', 'DOH',
        'PBK_RA_Expense_Revenue_M81_121321.xlsx',
        dict(distance=96.4, tolls=8, lanes=6, opening=sc['PSC'][0][0], sites=3, extra='3 sites: km 19+500, 47+500, 70+900; 6 lanes 47.5 km / 4 lanes 48.9 km'),
        sc, 'Host route: M81 (96.4 km, 8 toll plazas).')
    # Railway
    sc = {k: parse_bsp(s) for k, s in [('PSC', 'SUM_PSC'), ('PPP', 'SUM_PPP'), ('PPP (State)', 'SUM_PPP รัฐ')]}
    add('BSP', 'BSP Light Rail (Bang Na–Suvarnabhumi)', 'รถไฟฟ้ารางเบา บางนา–สุวรรณภูมิ', 'Railway', 'BMA',
        'BSP_OperationMaintenance_Model_CPI2.5_040222.xlsx',
        dict(distance=14.9, stations=12, depots=1, opening=sc['PSC'][0][0], extra='12 stations, 1 depot, 1 OCC'),
        sc)
    sc = {'PSC': parse_glp('GLP_PSC.xlsx'), 'PPP': parse_glp('GLP_PPP.xlsx')}
    add('GLP', 'Grey Line Monorail (Watcharaphon–Thong Lo)', 'รถไฟฟ้าสายสีเทา ช่วงวัชรพล–ทองหล่อ', 'Railway', 'BMA',
        'GLP_Summary_Cost_PSC/PPP_CPI2.5_101222.xlsx',
        dict(distance=None, stations=None, depots=1, opening=None, extra='Depot at Watcharaphon; distance & stations not stated in model'),
        sc, 'Model is in operation-year index (Year 1–30); calendar year not stated.')

    out = []
    for p in P:
        scen = {}
        for k, (yrs, items) in p['scen'].items():
            for x in items:
                x['cat'] = categorise(x, p['type'])
                x['v'] = [round(v, 4) for v in x['v']]
            scen[k] = dict(years=yrs, items=items)
        q = {kk: vv for kk, vv in p.items() if kk != 'scen'}
        q['scenarios'] = scen
        out.append(q)
    return out

if __name__ == '__main__':
    data = build()
    cats = [dict(id=a, en=b, th=c) for a, b, c in CATS]
    json.dump(dict(projects=data, cats=cats), open('/home/claude/work/data.json', 'w'), ensure_ascii=False)
    for p in data:
        for k, s in p['scenarios'].items():
            O = sum(sum(i['v']) for i in s['items'] if i['sec'] == 'O')
            M = sum(sum(i['v']) for i in s['items'] if i['sec'] == 'M')
            cs = {}
            for i in s['items']: cs[i['cat']] = cs.get(i['cat'], 0) + sum(i['v'])
            print(f"{p['id']:7s} {k:18s} n={len(s['items']):3d} yrs={s['years'][0] if s['years'] else '-'} O={O:10.2f} M={M:10.2f} T={O+M:10.2f}", {c: round(v) for c, v in sorted(cs.items())})
