"""Build the dashboard data from the encrypted model files: decrypt -> extract -> apply admin edits -> encrypt -> _site/."""
import os, sys, json, shutil, tempfile, datetime, base64
HERE = os.path.dirname(os.path.abspath(__file__)); ROOT = os.path.dirname(HERE)
sys.path.insert(0, HERE)
import crypt
PW = os.environ.get('DATA_PASSWORD') or sys.exit('DATA_PASSWORD secret is not set')
SLOTS = ['M82', 'RPM', 'PPK', 'SSP', 'ET1', 'PPK_RA', 'PBK', 'TRA', 'RAP', 'BSP', 'GLP_PSC', 'GLP_PPP']

tmp = tempfile.mkdtemp()
for s in SLOTS:
    b = crypt.read_enc(os.path.join(ROOT, 'models', s + '.enc'), PW)
    if b is None: sys.exit(f'model file for {s} is missing')
    open(os.path.join(tmp, s + '.xlsx'), 'wb').write(b)
os.environ['MODEL_DIR'] = tmp
import extract3, sitecost

items, _A, _c = extract3.build(with_calcs=False)
base = json.loads(crypt.dec(open(os.path.join(HERE, 'base.enc'), 'rb').read(), PW))
st_path = os.path.join(ROOT, 'data', 'state.enc')
state = json.loads(crypt.dec(open(st_path, 'rb').read(), PW)) if os.path.exists(st_path) else {}
projects = base['projects']
# admin edits: project fields, whole site lists, layout plans, model file names
for p in projects:
    for k, v in (state.get('projects', {}).get(p['project']) or {}).items():
        if k not in ('project', 'sites', 'cost'): p[k] = v
    if p['project'] in state.get('sites', {}): p['sites'] = state['sites'][p['project']]
    if p.get('sites'): p['rest_area_sites'] = len(p['sites']); p['area_rai'] = round(sum(float(x.get('area_rai') or 0) for x in p['sites']), 2) or None
sitecost.site_costs(items, projects)
# checks: every project and case in the profiles must come out of the models
tot = {}
for r in items: tot[(r['project'], r['stage'], r['scenario'])] = tot.get((r['project'], r['stage'], r['scenario']), 0) + r['total_30y']
bad = []
for p in projects:
    for st in p['stages'].split('|'):
        for c in (p.get('cases_' + st) or '').split('|'):
            if c and not tot.get((p['project'], st, c)): bad.append(f"{p['project']} {st} {c}")
for k in sorted(tot): print('%-8s %-5s %-22s %12.2f' % (k[0], k[1], k[2], tot[k]))
if bad: sys.exit('no figures extracted for: ' + ', '.join(bad))

bundle = dict(om_items=items, projects=projects, item_notes=base['item_notes'], models=state.get('models', {}),
              built_at=datetime.datetime.now(datetime.timezone.utc).isoformat(timespec='seconds'))
blob = crypt.enc(json.dumps(bundle, ensure_ascii=False, separators=(',', ':')).encode(), PW)
salt, iv, ct = blob[4:20], blob[20:32], blob[32:]
out = os.path.join(ROOT, '_site'); shutil.rmtree(out, ignore_errors=True); shutil.copytree(os.path.join(ROOT, 'site'), out)
b64 = lambda b: base64.b64encode(b).decode()
json.dump({'v': 1, 'iter': crypt.ITER, 'salt': b64(salt), 'iv': b64(iv), 'ct': b64(ct)}, open(os.path.join(out, 'data.enc.json'), 'w'))
if os.path.isdir(os.path.join(ROOT, 'plans')): shutil.copytree(os.path.join(ROOT, 'plans'), os.path.join(out, 'plans'))
open(os.path.join(out, '.nojekyll'), 'w').write('')
open(os.path.join(out, 'robots.txt'), 'w').write('User-agent: *\nDisallow: /\n')
print('built', len(items), 'items')
