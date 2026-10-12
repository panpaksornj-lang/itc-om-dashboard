"""O&M cost of each rest-area site, as each model reports it site by site (million THB over the model period)."""
import re
import extract as X

def _tot(items):
    o = sum(sum(i['v']) for i in items if i['sec'] == 'O'); m = sum(sum(i['v']) for i in items if i['sec'] == 'M')
    return o, m

def _pbk(sheet):
    y, it = X.parse_pbk(sheet); return _tot(it)

def _rap_sites(ny=12):
    rows = X.rows_of('RAP.xlsx', 'Summary_O_M_Seperate_Inflate', 172)
    out = {}; site = None
    for r in rows:
        a = r[0]
        if isinstance(a, str) and a.startswith('จุดพักรถ'): site = a.strip(); out[site] = [0.0, 0.0]; continue
        if site and isinstance(a, str) and re.match(r'^[OM][A-Z]', a):
            v = sum(x for x in r[4:4 + ny] if isinstance(x, (int, float))) / 1e6
            out[site][0 if a[0] == 'O' else 1] += v
    return out

def site_costs(items, projects):
    """Attach cost = {PSC: {o, m, ny, basis}, PPP: {...}} to each rest-area site, in place."""
    def case_tot(proj, case):
        rs = [r for r in items if r['project'] == proj and r['scenario'] == case]
        return (sum(r['total_30y'] for r in rs if r['sec'] == 'Operation'), sum(r['total_30y'] for r in rs if r['sec'] == 'Maintenance'), rs[0]['nyears'] if rs else None) if rs else None
    for p in projects:
        if p.get('type') != 'Rest Area' or not p.get('sites'): continue
        pj = p['project']
        for s in p['sites']: s['cost'] = {}
        if pj == 'PPK_RA':
            for k in ('PSC', 'PPP'):
                t = case_tot(pj, k)
                if t: p['sites'][0]['cost'][k] = dict(o=t[0], m=t[1], ny=t[2], basis='โครงการมีแห่งเดียว')
        elif pj == 'TRA':
            for s in p['sites']:
                nm = 'น้ำพอง' if 'น้ำพอง' in s['name'] else 'ท่าแซะ'
                for k in ('PSC', 'PPP'):
                    t = case_tot(pj, k + ' ' + nm)
                    if t: s['cost'][k] = dict(o=t[0], m=t[1], ny=t[2], basis='แบบจำลองแยกรายแห่ง')
        elif pj == 'PBK':
            for k, sp in (('PSC', 'Sum PSC'), ('PPP', 'Sum PPP')):
                tail = '' if k == 'PSC' else ' '
                a19 = _pbk(f'{sp} กม.19 แยก{tail}'); a47 = _pbk(f'{sp} กม.47 แยก{tail}'); a4770 = _pbk(f'{sp} กม.47+70{tail}')
                a70 = (a4770[0] - a47[0], a4770[1] - a47[1])
                for s, v in zip(p['sites'], (a19, a47, a70)):
                    s['cost'][k] = dict(o=v[0], m=v[1], ny=30, basis='แบบจำลองแยกรายแห่ง (กม.70+900 = กม.47+70 หัก กม.47)')
        elif pj == 'RAP':
            per = _rap_sites(p.get('nyears_SEPO') or 12)
            for s in p['sites']:
                key = next((k for k in per if s['name'].replace('จุดพักรถ', '') in k), None)
                if key: s['cost']['PSC'] = dict(o=per[key][0], m=per[key][1], ny=p.get('nyears_SEPO') or 12, basis='แบบจำลองแยกรายจุด (เฉพาะ PSC)')
    return projects
