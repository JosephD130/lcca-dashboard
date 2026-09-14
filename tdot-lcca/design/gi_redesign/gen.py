import json, os
D = os.path.dirname(os.path.abspath(__file__))
css = open(D + '/_css.txt').read()
C = json.load(open(D + '/_content.json'))
HDR = ('<!doctype html>\n<html>\n<head>\n  <meta charset="utf-8">\n  <script src="./support.js"></script>\n'
       '</head>\n<body>\n<x-dc>\n<helmet>\n  <style>\n' + css + '  </style>\n</helmet>\n')
END = '</x-dc>\n</body>\n</html>\n'
COLS = 'ABCDEFGHIJ'
TABS = ('  <div class="tabs">'
        '<div class="tab">Overview</div><div class="tab">Instructions</div>'
        '<div class="tab on">General Information</div><div class="tab">Pay_Items</div>'
        '<div class="tab">Maintenance Policies</div><div class="tab">Summary</div>'
        '<div class="tab">Typical Values</div><div class="tab">Method</div></div>\n')


def build(W, rows):
    out = ['  <div class="colstrip"><div class="rh"></div>'
           + ''.join('<div class="c" style="width:%dpx;justify-content:center;border-right:1px solid #E3E6EA;'
                     'color:#77808C">%s</div>' % (W[k], k) for k in COLS) + '</div>\n', '  <div class="grid">\n']
    total = 18
    for h, cells in rows:
        seen = ''.join(sp for sp, *_ in cells)
        assert seen == COLS, 'row must cover A..J exactly, got %r' % seen
        body = ''.join('<div class="c %s" style="width:%dpx;%s">%s</div>'
                       % (cls, sum(W[k] for k in sp), style, text) for sp, text, cls, style in cells)
        out.append('    <div class="r" style="height:%dpx"><div class="rh"></div>%s</div>\n' % (h, body))
        total += h
    out.append('  </div>\n')
    return '<div class="win">\n' + ''.join(out) + TABS + '</div>\n', total + 32


def sheet(W, card_h, card_text, card_clip, status_cls, field_cls, gap_after_d):
    """One General Information board. card_clip renders the card at a fixed height that cuts it off."""
    E = ('E' if gap_after_d else '')
    lab, inp = 'C', 'D'
    R = []
    R.append((38, [('AB', '+  New Study', 'btn new nw', ''),
                   ('CD' + E, 'STEP 2 of 5: the project and the LCCA parameters. Fill in the grey cells below.',
                    'note nw', ''),
                   ('FGH' if gap_after_d else 'EFGH', '', '', ''),
                   ('IJ', '<div class="mark">TN</div><div><div class="t1">TDOT</div>'
                          '<div class="t2">Department of Transportation</div></div>', 'logo', '')]))
    R.append((6, [('ABCD' + E, '', '', ''), ('FGHIJ' if gap_after_d else 'FGHIJ', '', '', '')]
                 if gap_after_d else [('ABCDE', '', '', ''), ('FGHIJ', '', '', '')]))
    R.append((19, [('ABCD' + E if gap_after_d else 'ABCDE', '', '', ''),
                   ('FGHIJ', 'HOW TO USE THIS WORKBOOK', 'cardt', '')]))
    body_style = 'align-self:flex-start;height:%dpx;%s' % (card_h, 'overflow:hidden' if card_clip else '')
    R.append((card_h, [('ABCD' + E if gap_after_d else 'ABCDE', '', '', ''),
                       ('FGHIJ', card_text, 'cardb', body_style)]))
    R.append((8, [('ABCDE', '', '', ''), ('FGHIJ', '', '', '')]))
    R.append((19, [('ABCDE', 'Airport Information:', 'band nw', ''), ('FGHIJ', '', '', '')]))
    for i, (l, v, kind) in enumerate(C['AIRPORT']):
        cells = [('AB', '', '', ''), ('C', l, 'nw', ''),
                 ('D', v, 'grey' if kind == 'calc' else field_cls, '')]
        cells.append(('E', '', '', ''))
        if i == 0:
            cells.append(('FGHIJ', C['STATUS'], status_cls,
                          'align-self:flex-start;height:26px;' + ('overflow:hidden' if card_clip else '')))
        else:
            cells.append(('FGHIJ', '', '', ''))
        R.append((26 if i == 0 else 19, cells))
    R.append((8, [('ABCDE', '', '', ''), ('FGHIJ', '', '', '')]))
    R.append((19, [('ABCDE', 'Project Information:', 'band nw', ''), ('FGHIJ', '', '', '')]))
    for l, v, kind in C['PROJECT']:
        R.append((19, [('AB', '', '', ''), ('C', l, 'nw', ''), ('D', v, field_cls, ''),
                       ('E', '', '', ''), ('FGHIJ', '', '', '')]))
    R.append((8, [('ABCDE', '', '', ''), ('FGHIJ', '', '', '')]))
    R.append((19, [('ABCDE', 'LCCA Parameters:', 'band nw', ''), ('FGHIJ', '', '', '')]))
    for l, v, hint in C['PARAMS']:
        R.append((19, [('AB', '', '', ''), ('C', l, 'nw', ''), ('D', v, field_cls + ' right', ''),
                       ('E', '', '', ''), ('FGHIJ', hint, 'note nw', '')]))
    return R
