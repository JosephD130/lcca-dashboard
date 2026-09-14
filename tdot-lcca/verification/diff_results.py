"""Deep-compare two run_example results.json files: every leaf, numbers to the cent."""
import json, sys

def walk(o, path=''):
    if isinstance(o, dict):
        for k, v in o.items():
            yield from walk(v, '%s.%s' % (path, k) if path else str(k))
    elif isinstance(o, list):
        for i, v in enumerate(o):
            yield from walk(v, '%s[%d]' % (path, i))
    else:
        yield path, o

a = dict(walk(json.load(open(sys.argv[1]))))
b = dict(walk(json.load(open(sys.argv[2]))))
label = sys.argv[3] if len(sys.argv) > 3 else ''
only_a = sorted(set(a) - set(b))
only_b = sorted(set(b) - set(a))
diff = []
for k in sorted(set(a) & set(b)):
    x, y = a[k], b[k]
    if isinstance(x, (int, float)) and isinstance(y, (int, float)) and not isinstance(x, bool):
        if abs(x - y) > 0.005: diff.append((k, x, y))
    elif x != y:
        diff.append((k, x, y))
print('%s: %d leaves compared' % (label, len(set(a) & set(b))))
if only_a: print('  only in baseline (%d): %s' % (len(only_a), ', '.join(only_a[:8])))
if only_b: print('  only in new run (%d): %s' % (len(only_b), ', '.join(only_b[:8])))
if not diff:
    print('  no value differs')
else:
    print('  %d values differ:' % len(diff))
    for k, x, y in diff[:40]:
        print('    %-52s %r -> %r' % (k, x, y))
sys.exit(1 if diff else 0)
