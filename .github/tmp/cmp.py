import json, re, sys
def D(p):
    t = open(p, encoding="utf-8").read()
    return json.loads(re.search(r'<script id="data"[^>]*>(.*?)</script>', t, re.S).group(1))
old, new = D(sys.argv[1]), D(sys.argv[2])
print('built', old['meta']['built'], '->', new['meta']['built'], '| latest', old['meta']['latest'], new['meta']['latest'])
for st in ('sold', 'all'):
    o, n = old['study'][st], new['study'][st]
    print(f"\n== study[{st}]: events {o['n_events']} -> {n['n_events']}, homes {o['n_homes']} -> {n['n_homes']}, n_mgmt -> {n.get('n_mgmt')}")
    for m in ('overall', 'staffing', 'hours', 'rn', 'turnover'):
        for k in ('12', '24'):
            a = (o['metrics'][m]['summary'] or {}).get(k); b = (n['metrics'][m]['summary'] or {}).get(k)
            f = lambda v: f"n={v['n']} {v['before']}->{v['after']} sold {v['sold']} similar {v['ctrl']} w/s/b {v['worse']}/{v['same']}/{v['better']}" if v else None
            print(f'  {m:9} {k}mo  old: {f(a)}\n               new: {f(b)}')
        pa = [p for p in o['metrics'][m]['points'] if p['months'] == -12]; pb = [p for p in n['metrics'][m]['points'] if p['months'] == -12]
        if m == 'overall': print('  year before sale old', pa, '\n                   new', pb)
print('\n== levels (same homes panel)')
for m in ('overall', 'hours', 'rn'):
    print(m, 'old', old['levels'].get(m), '\n   new', new['levels'].get(m))
print('\n== head tiles'); print({k: (old['head'].get(k), new['head'].get(k)) for k in new['head'] if k not in ('per_year',) and old['head'].get(k) != new['head'].get(k)})
print('\n== sales whose numbers changed')
ko = {s['cec']: s for s in old['sales']}
for s in new['sales']:
    o = ko.get(s['cec'])
    if not o: print('new row', s['cec']); continue
    if o.get('before_date') != s.get('before_date') or o['m']['overall'] != s['m']['overall'] or o['ser'] != s['ser']:
        print(s['cec'], s['name'][:30], s['buyer'], '| before', o.get('before_date'), '->', s.get('before_date'), '| stars', o['m']['overall'], '->', s['m']['overall'], '| harm', o['ser'], '->', s['ser'])
print('\n== owner groups that changed')
go = {g['group']: g for g in old['groups']}
for g in new['groups']:
    a = go.get(g['group'])
    if a is None or json.dumps(a, sort_keys=True) != json.dumps(g, sort_keys=True):
        print(g['group'], '\n  old', json.dumps({k: v for k, v in (a or {}).items() if k != 'homes'}), '\n  new', json.dumps({k: v for k, v in g.items() if k != 'homes'}))
print('\nother top-level keys changed:', [k for k in new if k not in ('meta', 'study', 'levels', 'head', 'sales', 'groups') and json.dumps(old.get(k), sort_keys=True) != json.dumps(new.get(k), sort_keys=True)])
