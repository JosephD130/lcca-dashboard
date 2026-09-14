#!/usr/bin/env python3
"""Turn a worked-example run into a baseline the Windows harness can hold Excel to.

run_example.py recalculates the workbook with LibreOffice and cross-checks the arithmetic in
Python. This writes the numbers it read out to baseline.json, cell by cell, so Run-ExcelChecks.ps1
can open the same workbook in real Excel, force a full rebuild, and prove the two engines agree.

usage: python3 verification/win/make_baseline.py <run dir> [out.json]
"""
import json, os, sys

RUN = sys.argv[1]
OUT = sys.argv[2] if len(sys.argv) > 2 else os.path.join(os.path.dirname(__file__), 'baseline.json')
r = json.load(open(os.path.join(RUN, 'results.json')))

cells = []
add = lambda sheet, ref, want, tol=0.5: cells.append(
    {'sheet': sheet, 'cell': ref, 'expect': want, 'tol': tol})

# the Summary results table, row by row
for i, row in enumerate(r['summary']['results']):
    if row.get('G') in (None, ''): continue
    for col in ('J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q'):
        if isinstance(row.get(col), (int, float)):
            add('Summary', '%s%d' % (col, 12 + i), row[col])
    add('Summary', 'R%d' % (12 + i), row['R'], 1e-6)

# the six dashboard tiles, as the text they display
for k, (ref, val) in enumerate(zip(
        ['G4', 'K4', 'O4', 'G8', 'K8', 'O8'],
        [r['summary']['kpi'][j] for j in (3, 4, 5, 12, 13, 14)])):
    cells.append({'sheet': 'Summary', 'cell': ref, 'expect': val, 'tol': None})

# the RealCost comparison block
for i, row in enumerate(r['summary'].get('comparison', [])):
    if not isinstance(row.get('H'), (int, float)): continue
    for col in ('H', 'I', 'J', 'K', 'L', 'M'):
        add('Summary', '%s%d' % (col, 22 + i), row[col])

# each alternative worksheet: the initial construction total and the net present worth
for a in r['alternatives']:
    add(a['sheet'], 'TOTAL', a['initial'])
    add(a['sheet'], 'NPW', a['NPW'])

base = {'example': r['example'], 'scenario': r['scenario'],
        'workbook': [f for f in os.listdir(RUN) if f.endswith('.xlsx')][0],
        'cells': cells}
json.dump(base, open(OUT, 'w'), indent=1)
print('wrote %s: %d cells from %s' % (OUT, len(cells), RUN))
