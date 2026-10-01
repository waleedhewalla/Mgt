"""Static Arabic diagrams for the print edition (matplotlib + Amiri; matplotlib shapes Arabic itself).
Kinds: flow, cycle, levels, matrix, hub, bars. One hue family (the book's greens), text in ink tokens."""
import math, os, re
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib import font_manager as fm
from matplotlib.patches import FancyBboxPatch, Circle, FancyArrowPatch, Polygon

HERE = os.path.dirname(os.path.abspath(__file__))
FONT = os.path.join(HERE, '..', 'fonts', 'Amiri-1.000', 'Amiri-Regular.ttf')
FONT_B = os.path.join(HERE, '..', 'fonts', 'Amiri-1.000', 'Amiri-Bold.ttf')
fm.fontManager.addfont(FONT)
fm.fontManager.addfont(FONT_B)
plt.rcParams['font.family'] = 'Amiri'
plt.rcParams['text.parse_math'] = False

INK, INK2 = '#1F2A24', '#4A5A50'
RAMP = ['#E3F1E6', '#C2E1C8', '#94C9A0', '#5FAA72', '#2E7D42', '#1A5C2A']     # light -> dark
GOLD = '#8B6914'
SURF = '#FFFFFF'


def wrap(s, n=14):
    words, lines, cur = s.split(), [], ''
    for w in words:
        if len(cur) + len(w) + 1 > n and cur:
            lines.append(cur)
            cur = w
        else:
            cur = (cur + ' ' + w).strip()
    if cur:
        lines.append(cur)
    return '\n'.join(lines)


def fill_for(i, n):
    """Sequential steps across the ramp (skip the palest for legibility of borders)."""
    if n == 1:
        return RAMP[3]
    idx = 1 + round(i * (len(RAMP) - 2) / (n - 1))
    return RAMP[idx]


def ink_on(hexcol):
    r, g, b = (int(hexcol[i:i + 2], 16) for i in (1, 3, 5))
    return '#FFFFFF' if (0.299 * r + 0.587 * g + 0.114 * b) < 140 else INK


def _box(ax, x, y, w, h, fc, text, sub=None, fs=13):
    ax.add_patch(FancyBboxPatch((x - w / 2, y - h / 2), w, h, boxstyle='round,pad=0.02,rounding_size=0.06',
                                fc=fc, ec=RAMP[3] if fc in RAMP[:2] else SURF, lw=1.2))
    col = ink_on(fc)
    t = wrap(text, 20)
    if sub:
        two = '\n' in t
        ax.text(x, y + h * (0.22 if two else 0.16), t, ha='center', va='center', fontsize=fs - (1 if two else 0), color=col,
                fontweight='bold', linespacing=0.95)
        ax.text(x, y - h * (0.30 if two else 0.22), wrap(sub, 22), ha='center', va='center', fontsize=fs - 3, color=col)
    else:
        ax.text(x, y, wrap(text, 16), ha='center', va='center', fontsize=fs, color=col, fontweight='bold')


def _arrow(ax, p, q, col=INK2):
    ax.add_patch(FancyArrowPatch(p, q, arrowstyle='-|>', mutation_scale=14, lw=1.6, color=col, shrinkA=2, shrinkB=2))


def _finish(fig, ax, title, path):
    if title:
        fig.suptitle(title, fontsize=17, color=INK, fontweight='bold', y=0.98)
    ax.set_aspect('equal')
    ax.axis('off')
    fig.savefig(path, dpi=200, bbox_inches='tight', facecolor=SURF)
    plt.close(fig)


def items_norm(items):
    out = []
    for it in items:
        if isinstance(it, (list, tuple)):
            out.append((str(it[0]), str(it[1]) if len(it) > 1 and it[1] else None))
        else:
            out.append((str(it), None))
    return out


def flow(title, items, path):
    items = items_norm(items)
    n = len(items)
    per_row = n if n <= 3 else (2 if n == 4 else 3)
    rows = math.ceil(n / per_row)
    lines = max(wrap(t, 20).count('\n') + (wrap(s_, 22).count('\n') + 1 if s_ else 0) for t, s_ in items) + 1
    w, h, gap, vgap = 2.6, 1.0 + 0.32 * lines, 0.6, 0.9
    fig, ax = plt.subplots(figsize=(3.3 * per_row, (h + vgap) * rows * 1.05 + 0.7))
    pos = []
    for i in range(n):
        r, c = divmod(i, per_row)
        if r % 2:                                   # snake: odd rows run left-to-right under the previous row's end
            c = per_row - 1 - c
        pos.append((-(c * (w + gap)), -r * (h + vgap)))
    for i in range(n - 1):
        (x1, y1), (x2, y2) = pos[i], pos[i + 1]
        if y1 == y2:
            d = 1 if x2 < x1 else -1
            _arrow(ax, (x1 - d * w / 2, y1), (x2 + d * w / 2, y2))
        else:
            _arrow(ax, (x1, y1 - h / 2), (x2, y2 + h / 2))
    for i, (x, y) in enumerate(pos):
        _box(ax, x, y, w, h, fill_for(i, n), items[i][0], items[i][1])
        ax.text(x + w / 2 - 0.1, y + h / 2 - 0.12, str(i + 1), ha='right', va='top', fontsize=11, color=ink_on(fill_for(i, n)))
    ax.set_xlim(min(p[0] for p in pos) - w / 2 - 0.2, max(p[0] for p in pos) + w / 2 + 0.2)
    ax.set_ylim(min(p[1] for p in pos) - h / 2 - 0.2, h / 2 + 0.2)
    _finish(fig, ax, title, path)


def cycle(title, items, path):
    items = items_norm(items)
    n = len(items)
    R = max(3.0, 0.62 * n)
    fig, ax = plt.subplots(figsize=(8.5, 8.5))
    pts = []
    for i in range(n):
        a = math.pi / 2 - 2 * math.pi * i / n      # start at top, run clockwise (reading direction for RTL)
        pts.append((R * math.cos(a), R * math.sin(a)))
    for i in range(n):
        p, q = pts[i], pts[(i + 1) % n]
        d = 1.45 / R
        a1 = math.atan2(p[1], p[0]) - d
        a2 = math.atan2(q[1], q[0]) + d
        Rr = R + 0.05
        ax.add_patch(FancyArrowPatch((Rr * math.cos(a1), Rr * math.sin(a1)), (Rr * math.cos(a2), Rr * math.sin(a2)),
                                     connectionstyle='arc3,rad=-0.12', arrowstyle='-|>', mutation_scale=15, lw=1.6,
                                     color=INK2, shrinkA=0, shrinkB=0, zorder=0))
    for i, (x, y) in enumerate(pts):
        _box(ax, x, y, 2.3, 1.3, fill_for(i, n), items[i][0], items[i][1], fs=12)
    ax.set_xlim(-R - 1.6, R + 1.6)
    ax.set_ylim(-R - 1.2, R + 1.2)
    _finish(fig, ax, title, path)


def levels(title, items, path):
    """items from base (first) to top (last), drawn as a stepped pyramid."""
    items = items_norm(items)
    n = len(items)
    fig, ax = plt.subplots(figsize=(9, 0.95 * n + 0.4))
    H = 1.0
    step = 4.8 / max(n, 1)
    for i in range(n):
        wb = 9.2 - i * step
        wt = wb - step
        y = i * H
        ax.add_patch(Polygon([(-wb / 2, y), (wb / 2, y), (wt / 2, y + H * 0.92), (-wt / 2, y + H * 0.92)],
                             closed=True, fc=fill_for(i, n), ec=SURF, lw=2))
        label = items[i][0] + (('  |  ' + items[i][1]) if items[i][1] else '')
        ax.text(0, y + H * 0.46, label, ha='center', va='center', fontsize=12.5 if len(label) < 34 else 11,
                color=ink_on(fill_for(i, n)), fontweight='bold')
    ax.set_xlim(-4.8, 4.8)
    ax.set_ylim(-0.1, n * H)
    if title:
        fig.suptitle(title, fontsize=17, color=INK, fontweight='bold')
    ax.axis('off')
    fig.savefig(path, dpi=200, bbox_inches='tight', facecolor=SURF)
    plt.close(fig)


def hub(title, items, path, center=None):
    items = items_norm(items)
    if center is None:
        center, items = items[0][0], items[1:]
    n = len(items)
    R = 3.2
    fig, ax = plt.subplots(figsize=(8.5, 8.5))
    ax.add_patch(Circle((0, 0), 1.35, fc=RAMP[5], ec=SURF, lw=2))
    ax.text(0, 0, wrap(center, 12), ha='center', va='center', fontsize=15, color='#FFFFFF', fontweight='bold')
    for i, (t, s) in enumerate(items):
        a = math.pi / 2 - 2 * math.pi * i / n
        x, y = R * math.cos(a), R * math.sin(a)
        ax.plot([1.35 * math.cos(a), (R - 1.0) * math.cos(a)], [1.35 * math.sin(a), (R - 1.0) * math.sin(a)], color=INK2, lw=1.4)
        _box(ax, x, y, 2.3, 1.25, fill_for(i, max(n, 2)) if n > 1 else RAMP[3], t, s, fs=12)
    ax.set_xlim(-R - 1.5, R + 1.5)
    ax.set_ylim(-R - 1.1, R + 1.1)
    _finish(fig, ax, title, path)


def matrix(title, items, path):
    """items: first row = column headers (first cell = corner); following rows = [row label, cells...]."""
    rows = [list(map(str, r)) for r in items]
    head, body = rows[0], rows[1:]
    nc, nr = len(head), len(body)
    cw, ch = 2.6, 0.95
    fig, ax = plt.subplots(figsize=(min(cw * nc * 1.05, 14), ch * (nr + 1) * 1.1 + 0.8))
    for c in range(nc):                                    # columns laid right-to-left
        x = -c * cw
        ax.add_patch(FancyBboxPatch((x - cw / 2 + 0.04, -ch / 2 + 0.04), cw - 0.08, ch - 0.08, boxstyle='round,pad=0,rounding_size=0.05',
                                    fc=RAMP[5] if c else RAMP[4], ec=SURF))
        ax.text(x, 0, wrap(head[c], 18), ha='center', va='center', fontsize=12, color='#FFFFFF', fontweight='bold')
    for r, row in enumerate(body, 1):
        y = -r * ch
        for c in range(nc):
            x = -c * cw
            fc = RAMP[2] if c == 0 else (RAMP[0] if r % 2 else RAMP[1])
            ax.add_patch(FancyBboxPatch((x - cw / 2 + 0.04, y - ch / 2 + 0.04), cw - 0.08, ch - 0.08,
                                        boxstyle='round,pad=0,rounding_size=0.05', fc=fc, ec=SURF))
            txt = row[c] if c < len(row) else ''
            ax.text(x, y, wrap(txt, 20), ha='center', va='center', fontsize=11, color=INK, fontweight='bold' if c == 0 else 'normal')
    ax.set_xlim(-(nc - 0.5) * cw, cw / 2)
    ax.set_ylim(-(nr + 0.5) * ch, ch / 2)
    _finish(fig, ax, title, path)


def bars(title, items, path):
    """items: [[label, value], ...] — only numbers stated in the text."""
    labs = [str(a) for a, _ in items]
    vals = [float(str(b).replace('٫', '.').translate(str.maketrans('٠١٢٣٤٥٦٧٨٩', '0123456789')).rstrip('%')) for _, b in items]
    fig, ax = plt.subplots(figsize=(8, 0.75 * len(items) + 1.4))
    y = range(len(items))
    ax.barh(list(y), vals, color=RAMP[4], height=0.55, edgecolor=SURF, linewidth=2)
    ax.set_yticks(list(y))
    ax.set_yticklabels([wrap(l, 22) for l in labs], fontsize=12, color=INK)
    ax.invert_yaxis()
    ax.yaxis.tick_right()
    ax.invert_xaxis()
    for i, v in enumerate(vals):
        ax.text(v, i, f' {v:g} ', va='center', ha='right', fontsize=11, color=INK2)
    for s in ('top', 'left', 'bottom'):
        ax.spines[s].set_visible(False)
    ax.spines['right'].set_color('#BBBBBB')
    ax.tick_params(axis='x', colors=INK2)
    ax.grid(axis='x', color='#E5E5E5', lw=0.8)
    ax.set_axisbelow(True)
    fig.suptitle(title, fontsize=16, color=INK, fontweight='bold')
    fig.savefig(path, dpi=200, bbox_inches='tight', facecolor=SURF)
    plt.close(fig)


KINDS = {'flow': flow, 'cycle': cycle, 'levels': levels, 'matrix': matrix, 'hub': hub, 'bars': bars}


CIRC = {chr(0x2460 + i): str(i + 1) for i in range(20)}
CIRC.update({chr(0x2776 + i): str(i + 1) for i in range(10)})


def clean(x):
    if isinstance(x, str):
        x = ''.join(CIRC.get(ch, ch) for ch in x)
        x = re.sub('[\u2600-\u27BF\u2B00-\u2BFF\uFE0F\U0001F000-\U0001FFFF\u25A0-\u25FF]', '', x)
        return re.sub(r'\s{2,}', ' ', x).strip()
    if isinstance(x, (list, tuple)):
        return [clean(y) for y in x]
    if isinstance(x, dict):
        return {k: clean(v) for k, v in x.items()}
    return x


def render(spec, path):
    kind = spec.get('kind', 'flow')
    fn = KINDS.get(kind, flow)
    fn(clean(spec.get('title', '')), clean(spec.get('items', [])), path)
    return path


if __name__ == '__main__':
    os.makedirs('figs_test', exist_ok=True)
    render({'kind': 'cycle', 'title': 'الدورة السباعية للفعل الإداري', 'items': [['الاستشعار', 'إدراك الموقف'], ['التأطير الغائي', 'ماذا يخدم القرار؟'], ['الشورى', 'تداول قبل الحسم'], ['التثبت', 'تحقق من المعلومات'], ['العزم', 'الحسم بعد التداول'], ['التوكل والتنفيذ', 'التزام وأخذ بالأسباب'], ['المتابعة والمحاسبة', 'رصد الأثر']]}, 'figs_test/cycle.png')
    render({'kind': 'flow', 'title': 'مراحل التحول المؤسسي الخمس', 'items': [['اليقظة', 'إدراك الفجوة'], ['العزم', 'قرار التغيير'], ['البناء', 'الأنظمة'], ['الترسيخ', 'الثقافة'], ['الاستدامة', 'التجدد']]}, 'figs_test/flow.png')
    render({'kind': 'levels', 'title': 'مستويات قياس الأثر', 'items': [['الفرد', 'السلوك'], ['الفريق', 'التعاون'], ['المؤسسة', 'الأنظمة'], ['المجتمع', 'العمران'], ['الأجيال', 'الأثر الممتد']]}, 'figs_test/levels.png')
    render({'kind': 'hub', 'title': 'الحوكمة القرآنية', 'items': ['الاستخلاف', ['الشورى', 'قرار جماعي'], ['المساءلة', 'لا استثناء'], ['الشفافية', 'إفصاح صادق'], ['الأمانة', 'رد الحقوق']]}, 'figs_test/hub.png')
    render({'kind': 'matrix', 'title': 'الأسلمة الشكلية والتحول الحقيقي', 'items': [['المعيار', 'الأسلمة الشكلية', 'التحول الحقيقي'], ['الشورى', 'إخطار بعد القرار', 'تغير القرار'], ['المساءلة', 'انتقائية', 'تشمل القمة'], ['القياس', 'مالي فقط', 'ماذا بنينا وأفسدنا']]}, 'figs_test/matrix.png')
    render({'kind': 'bars', 'title': 'مثال أعمدة', 'items': [['الحوكمة', 6], ['الإنسان', 7], ['الفعل', 6], ['التحول', 8]]}, 'figs_test/bars.png')
    print('ok')
