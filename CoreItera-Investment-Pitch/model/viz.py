"""Visuals for the proposal page: every chart is drawn to scale from model numbers. Pure Python, returns HTML/SVG strings."""
import math


def quarter_chart(qs):
    W, H, L, R, T, B = 720, 300, 60, 20, 20, 50
    vals = [q["revenue"] for q in qs] + [q["costs"] for q in qs] + [q["cash_end"] for q in qs]
    top = max(vals)
    top = math.ceil(top / 50) * 50
    bottom = min(0, math.floor(min(vals) / 10) * 10)
    span = top - bottom
    def yv(v):
        return T + (top - v) / span * (H - T - B)
    n = len(qs)
    step = (W - L - R) / n
    bw = step * 0.28
    g = []
    for t in range(int(bottom), int(top) + 1, 50):
        g.append(f'<line x1="{L}" x2="{W - R}" y1="{yv(t):.1f}" y2="{yv(t):.1f}" class="grid"/>'
                 f'<text x="{L - 8}" y="{yv(t) + 4:.1f}" class="ax" text-anchor="end">{t}</text>')
    g.append(f'<line x1="{L}" x2="{W - R}" y1="{yv(0):.1f}" y2="{yv(0):.1f}" class="zero"/>')
    pts = []
    for i, q in enumerate(qs):
        x0 = L + i * step + step * 0.18
        for j, (key, c) in enumerate((("revenue", "c-rev"), ("costs", "c-cost"))):
            v = q[key]
            x = x0 + j * (bw + 3)
            g.append(f'<rect x="{x:.1f}" y="{yv(v):.1f}" width="{bw:.1f}" height="{max(0.5, yv(0) - yv(v)):.1f}" class="{c}"/>')
        cx = L + i * step + step / 2
        pts.append((cx, yv(q["cash_end"])))
        g.append(f'<text x="{cx:.1f}" y="{H - B + 20}" class="ax" text-anchor="middle">{q["quarter"]}</text>')
    g.append('<polyline fill="none" class="c-cash" points="' + " ".join(f"{x:.1f},{y:.1f}" for x, y in pts) + '"/>')
    for (x, y), q in zip(pts, qs):
        g.append(f'<circle cx="{x:.1f}" cy="{y:.1f}" r="4" class="c-cashdot"/>'
                 f'<text x="{x + 8:.1f}" y="{y - 9:.1f}" class="lab" text-anchor="start">{q["cash_end"]:.0f}</text>')
    return (f'<div class="chart"><svg viewBox="0 0 {W} {H}" role="img" aria-label="Quarterly revenue, costs and cash">'
            + "".join(g) + "</svg>"
            '<div class="legend"><span><i class="c-rev"></i>Revenue</span><span><i class="c-cost"></i>Costs</span>'
            '<span><i class="c-cashline"></i>Cash at the end of the quarter</span><span>Rp million</span></div></div>')


def revenue_chart(years):
    W, H, L, R, T, B = 720, 280, 70, 20, 20, 40
    top = math.ceil(max(v["revenue"] for v in years.values()) / 2000) * 2000
    def yv(v):
        return T + (top - v) / top * (H - T - B)
    step = (W - L - R) / len(years)
    bw = step * 0.5
    g = []
    for t in range(0, int(top) + 1, 2000):
        g.append(f'<line x1="{L}" x2="{W - R}" y1="{yv(t):.1f}" y2="{yv(t):.1f}" class="grid"/>'
                 f'<text x="{L - 8}" y="{yv(t) + 4:.1f}" class="ax" text-anchor="end">{t / 1000:.0f}</text>')
    for i, (y, v) in enumerate(years.items()):
        x = L + i * step + (step - bw) / 2
        base = 0.0
        for key, c in (("plan", "c-plan"), ("fees", "c-fee"), ("copilot", "c-cop")):
            h = v[key]
            g.append(f'<rect x="{x:.1f}" y="{yv(base + h):.1f}" width="{bw:.1f}" height="{max(0.5, yv(base) - yv(base + h)):.1f}" class="{c}"/>')
            base += h
        g.append(f'<text x="{x + bw / 2:.1f}" y="{yv(base) - 7:.1f}" class="lab" text-anchor="middle">{base / 1000:.1f}</text>'
                 f'<text x="{x + bw / 2:.1f}" y="{H - B + 20}" class="ax" text-anchor="middle">{y}</text>')
    return (f'<div class="chart"><svg viewBox="0 0 {W} {H}" role="img" aria-label="Revenue by year and source">' + "".join(g) + "</svg>"
            '<div class="legend"><span><i class="c-plan"></i>Monthly plans</span><span><i class="c-fee"></i>Rp5,000 payment fees</span>'
            '<span><i class="c-cop"></i>Copilot seats</span><span>Rp billion a year</span></div></div>')


def positioning_map():
    W, H, L, B = 640, 420, 70, 60
    pts = [("CoreItera", 86, 86, "us"),
           ("WhatsApp, spreadsheets and Word", 8, 62, ""),
           ("Global practice software (SimplePractice, Carepatron)", 26, 14, ""),
           ("Clinic record software (Klinik Pintar)", 8, 40, ""),
           ("Counselling apps (Riliv, Satu Persen)", 32, 74, ""),
           ("Online test sites (KantorKu)", 48, 52, "")]
    def xy(x, y):
        return L + x / 100 * (W - L - 20), 20 + (100 - y) / 100 * (H - 20 - B)
    g = [f'<rect x="{L}" y="20" width="{W - L - 20}" height="{H - 20 - B}" class="frame"/>',
         f'<line x1="{xy(50, 0)[0]:.0f}" x2="{xy(50, 0)[0]:.0f}" y1="20" y2="{H - B}" class="grid"/>',
         f'<line x1="{L}" x2="{W - 20}" y1="{xy(0, 50)[1]:.0f}" y2="{xy(0, 50)[1]:.0f}" class="grid"/>',
         f'<text x="{(L + W - 20) / 2:.0f}" y="{H - 22}" class="ax" text-anchor="middle">Handles assessment and signed reports →</text>',
         f'<text x="18" y="{(20 + H - B) / 2:.0f}" class="ax" text-anchor="middle" transform="rotate(-90 18 {(20 + H - B) / 2:.0f})">Fits how Indonesian bureaus work →</text>']
    for name, x, y, cls in pts:
        cx, cy = xy(x, y)
        anchor = "end" if x > 60 else "start"
        dx = -10 if anchor == "end" else 10
        g.append(f'<circle cx="{cx:.0f}" cy="{cy:.0f}" r="{9 if cls else 6}" class="{"dot-us" if cls else "dot"}"/>'
                 f'<text x="{cx + dx:.0f}" y="{cy + 4:.0f}" class="{"lab-us" if cls else "lab"}" text-anchor="{anchor}">{name}</text>')
    return f'<div class="chart"><svg viewBox="0 0 {W} {H}" role="img" aria-label="Positioning map">' + "".join(g) + "</svg></div>"




# ---------------------------------------------------------------- pyramid
def pyramid(answer, pillars):
    """answer: governing thought (html). pillars: [(letter, headline, [(anchor, point), ...]), ...]"""
    cols = []
    for letter, head, points in pillars:
        items = "".join(f'<li><a href="#{a}">{p}</a></li>' for a, p in points)
        cols.append(f'<div class="pyr-col"><div class="pyr-head"><span class="pyr-letter">{letter}</span>{head}</div><ul>{items}</ul></div>')
    return (f'<div class="pyr" role="group" aria-label="The argument in one picture">'
            f'<div class="pyr-top">{answer}</div><div class="pyr-stem" aria-hidden="true"></div>'
            f'<div class="pyr-row">{"".join(cols)}</div></div>')


# ---------------------------------------------------------------- horizontal bars to scale
def hbars(rows, unit="", note=""):
    """rows: [(label, value, css_class, value_text)]; bars share one scale."""
    top = max(v for _, v, _, _ in rows)
    out = []
    for label, v, cls, txt in rows:
        w = max(0.6, 100 * v / top)
        out.append(f'<div class="hb-row"><div class="hb-label">{label}</div>'
                   f'<div class="hb-track"><div class="hb-bar {cls}" style="width:{w:.2f}%"></div>'
                   f'<span class="hb-val">{txt}</span></div></div>')
    n = f'<p class="note">{note}</p>' if note else ""
    return f'<div class="hb">{"".join(out)}{n}</div>'


# ---------------------------------------------------------------- 24-hour band
def day_band(admin_from, admin_to, after_share):
    W, L, R = 720, 120, 16
    span = W - L - R
    def x(h):
        return L + h / 24 * span
    g = []
    for h in range(0, 25, 3):
        g.append(f'<line x1="{x(h):.1f}" x2="{x(h):.1f}" y1="18" y2="150" class="grid"/>'
                 f'<text x="{x(h):.1f}" y="172" class="ax" text-anchor="middle">{h:02d}:00</text>')
    g.append(f'<rect x="{x(0):.1f}" y="18" width="{x(admin_from) - x(0):.1f}" height="132" class="night"/>')
    g.append(f'<rect x="{x(admin_to):.1f}" y="18" width="{x(24) - x(admin_to):.1f}" height="132" class="night"/>')
    g.append(f'<text x="{L - 10}" y="58" class="lab" text-anchor="end">Admin</text>')
    g.append(f'<rect x="{x(admin_from):.1f}" y="40" width="{x(admin_to) - x(admin_from):.1f}" height="28" rx="4" class="c-cost"/>')
    g.append(f'<text x="{(x(admin_from) + x(admin_to)) / 2:.1f}" y="59" class="inbar" text-anchor="middle">answers and closes</text>')
    g.append(f'<text x="{L - 10}" y="112" class="lab" text-anchor="end">Sales agent</text>')
    g.append(f'<rect x="{x(0):.1f}" y="94" width="{span:.1f}" height="28" rx="4" class="c-rev"/>')
    g.append(f'<text x="{x(3.5):.1f}" y="113" class="inbar" text-anchor="middle">answers, sorts, books</text>')
    g.append(f'<text x="{(x(admin_from) + x(admin_to)) / 2:.1f}" y="113" class="inbar" text-anchor="middle">hands hot leads to the admin</text>')
    g.append(f'<text x="{x(21):.1f}" y="113" class="inbar" text-anchor="middle">answers, sorts, books</text>')
    g.append(f'<text x="{x(21):.1f}" y="36" class="lab" text-anchor="middle">{after_share:.0%} of chats arrive here</text>')
    return (f'<div class="chart"><svg viewBox="0 0 {W} 182" role="img" aria-label="Who answers a chat at each hour">' + "".join(g) + "</svg>"
            '<div class="legend"><span><i class="night-key"></i>Outside admin hours</span><span><i class="c-cost"></i>Admin</span><span><i class="c-rev"></i>Sales agent</span></div></div>')


# ---------------------------------------------------------------- funnel, today vs with the agent
def compare_bars(rows, left="Today", right="With CoreItera"):
    """rows: [(label, today, with, today_txt, with_txt)] on one scale."""
    top = max(max(a, b) for _, a, b, _, _ in rows)
    out = [f'<div class="cb-legend"><span><i class="c-cost"></i>{left}</span><span><i class="c-rev"></i>{right}</span></div>']
    for label, a, b, ta, tb in rows:
        out.append(f'<div class="cb-row"><div class="hb-label">{label}</div><div class="cb-pair">'
                   f'<div class="hb-track"><div class="hb-bar c-cost" style="width:{max(0.6, 100 * a / top):.1f}%"></div><span class="hb-val">{ta}</span></div>'
                   f'<div class="hb-track"><div class="hb-bar c-rev" style="width:{max(0.6, 100 * b / top):.1f}%"></div><span class="hb-val">{tb}</span></div>'
                   f'</div></div>')
    return f'<div class="hb">{"".join(out)}</div>'


# ---------------------------------------------------------------- product stack
def product_stack(layers):
    """layers: [(tag, css, title, items:[(name, status_html)])]"""
    out = []
    for tag, css, title, items in layers:
        li = "".join(f'<div class="ps-item"><b>{n}</b><span>{s}</span></div>' for n, s in items)
        out.append(f'<div class="ps-layer {css}"><div class="ps-tag">{tag}</div><div class="ps-title">{title}</div><div class="ps-items">{li}</div></div>')
    return f'<div class="ps">{"".join(out)}</div>'


# ---------------------------------------------------------------- roadmap with proof gates
def roadmap(phases, lanes):
    """phases: [{name, when, gate, cells:{lane: html}}]; lanes: [lane names]"""
    head = "".join(f'<div class="rm-ph rm-p{i}"><b>{p["name"]}</b><span>{p["when"]}</span></div>' for i, p in enumerate(phases))
    rows = ""
    for lane in lanes:
        rows += f'<div class="rm-lane">{lane}</div>' + "".join(f'<div class="rm-cell">{p["cells"].get(lane, "")}</div>' for p in phases)
    gates = "".join(f'<div class="rm-gate"><span class="diamond" aria-hidden="true"></span>{p["gate"]}</div>' for p in phases)
    n = len(phases)
    return (f'<div class="rm-wrap"><div class="rm" style="grid-template-columns:120px repeat({n},minmax(150px,1fr))">'
            f'<div class="rm-corner"></div>{head}{rows}<div class="rm-lane">Proof gate</div>{gates}</div></div>')


# ---------------------------------------------------------------- engineering plan (Gantt)
def gantt(items, start, months):
    """items: [(build, owner, first_month_index, length, css)], month 0 = start (y, m)."""
    W, L, T, rowh = 780, 280, 26, 30
    H = T + rowh * len(items) + 10
    span = W - L - 10
    def x(i):
        return L + i / months * span
    names = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
    g = []
    y0, m0 = start
    for i in range(months + 1):
        g.append(f'<line x1="{x(i):.1f}" x2="{x(i):.1f}" y1="{T - 6}" y2="{H - 8}" class="grid"/>')
        if i < months:
            mm = (m0 - 1 + i) % 12
            yy = y0 + (m0 - 1 + i) // 12
            lab = names[mm] + (f" {str(yy)[2:]}" if mm in (0,) or i == 0 else "")
            g.append(f'<text x="{(x(i) + x(i + 1)) / 2:.1f}" y="{T - 10}" class="ax" text-anchor="middle">{lab}</text>')
    for r, (build, owner, a, n, css) in enumerate(items):
        y = T + r * rowh
        g.append(f'<text x="{L - 10}" y="{y + 19}" class="lab" text-anchor="end">{build}</text>')
        w = x(a + n) - x(a) - 4
        g.append(f'<rect x="{x(a) + 2:.1f}" y="{y + 5}" width="{w:.1f}" height="20" rx="4" class="{css}"/>')
        if w >= 6.5 * len(owner):
            g.append(f'<text x="{x(a) + 8:.1f}" y="{y + 19}" class="inbar">{owner}</text>')
        else:
            g.append(f'<text x="{x(a + n) + 6:.1f}" y="{y + 19}" class="lab">{owner}</text>')
    return f'<div class="chart"><svg viewBox="0 0 {W} {H}" role="img" aria-label="Engineering plan by month">' + "".join(g) + "</svg></div>"


# ---------------------------------------------------------------- flywheel
def flywheel(nodes, center):
    W = H = 480
    cx, cy, R = W / 2, H / 2, 160
    n = len(nodes)
    pts = [(cx + R * math.sin(2 * math.pi * i / n), cy - R * math.cos(2 * math.pi * i / n)) for i in range(n)]
    g = ['<defs><marker id="fw-arrow" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="7" markerHeight="7" orient="auto-start-reverse">'
         '<path d="M0,0 L10,5 L0,10 z" class="fw-head"/></marker></defs>']
    for i in range(n):
        a0 = 2 * math.pi * (i + 0.18) / n
        a1 = 2 * math.pi * (i + 0.82) / n
        x0, y0 = cx + R * math.sin(a0), cy - R * math.cos(a0)
        x1, y1 = cx + R * math.sin(a1), cy - R * math.cos(a1)
        g.append(f'<path d="M{x0:.1f},{y0:.1f} A{R},{R} 0 0,1 {x1:.1f},{y1:.1f}" class="fw-arc" marker-end="url(#fw-arrow)"/>')
    for (x, y), (title, sub) in zip(pts, nodes):
        g.append(f'<circle cx="{x:.1f}" cy="{y:.1f}" r="60" class="fw-node"/>'
                 f'<text x="{x:.1f}" y="{y - 4:.1f}" class="fw-t" text-anchor="middle">{title}</text>'
                 f'<text x="{x:.1f}" y="{y + 14:.1f}" class="fw-s" text-anchor="middle">{sub}</text>')
    g.append(f'<text x="{cx}" y="{cy - 4}" class="fw-c" text-anchor="middle">{center[0]}</text>'
             f'<text x="{cx}" y="{cy + 16}" class="fw-s" text-anchor="middle">{center[1]}</text>')
    return f'<div class="chart fw"><svg viewBox="0 0 {W} {H}" role="img" aria-label="Growth flywheel">' + "".join(g) + "</svg></div>"


# ---------------------------------------------------------------- ownership bars
def cap_bars(steps):
    """steps: [(label, [(share, css, text)])]"""
    rows = []
    for label, segs in steps:
        seg = "".join(f'<div class="{c}" style="width:{100 * s:.2f}%">{t}</div>' for s, c, t in segs)
        rows.append(f'<div class="bar-row"><b>{label}</b><div class="bar">{seg}</div></div>')
    return ('<div class="bars" role="img" aria-label="Ownership before and after the deal">' + "".join(rows) +
            '<div class="legend"><span><i class="seg-o"></i>Galih and the other founders</span><span><i class="seg-r"></i>Riskia</span>'
            '<span><i class="seg-t"></i>TemanJourney</span></div></div>')


# ---------------------------------------------------------------- one stacked bar (use of money)
def stacked(items, total_txt):
    tot = sum(v for _, v, _ in items)
    segs = "".join(f'<div class="{c}" style="width:{100 * v / tot:.2f}%" title="{n}"></div>' for n, v, c in items)
    leg = "".join(f'<span><i class="{c}"></i>{n}</span>' for n, v, c in items)
    return f'<div class="bars"><div class="bar tall">{segs}</div><div class="legend">{leg}<span>{total_txt}</span></div></div>'


# ---------------------------------------------------------------- positioning map (generic)
def positioning(points, xlab, ylab):
    """points: [(name, x0-100, y0-100, is_us)]"""
    W, H, L, B = 640, 420, 70, 60
    def xy(x, y):
        return L + x / 100 * (W - L - 20), 20 + (100 - y) / 100 * (H - 20 - B)
    g = [f'<rect x="{L}" y="20" width="{W - L - 20}" height="{H - 20 - B}" class="frame"/>',
         f'<line x1="{xy(50, 0)[0]:.0f}" x2="{xy(50, 0)[0]:.0f}" y1="20" y2="{H - B}" class="grid"/>',
         f'<line x1="{L}" x2="{W - 20}" y1="{xy(0, 50)[1]:.0f}" y2="{xy(0, 50)[1]:.0f}" class="grid"/>',
         f'<text x="{(L + W - 20) / 2:.0f}" y="{H - 22}" class="ax" text-anchor="middle">{xlab} →</text>',
         f'<text x="18" y="{(20 + H - B) / 2:.0f}" class="ax" text-anchor="middle" transform="rotate(-90 18 {(20 + H - B) / 2:.0f})">{ylab} →</text>']
    for name, x, y, us in points:
        cx, cy = xy(x, y)
        anchor = "end" if x > 60 else "start"
        dx = -12 if anchor == "end" else 12
        g.append(f'<circle cx="{cx:.0f}" cy="{cy:.0f}" r="{9 if us else 6}" class="{"dot-us" if us else "dot"}"/>'
                 f'<text x="{cx + dx:.0f}" y="{cy + 4:.0f}" class="{"lab-us" if us else "lab"}" text-anchor="{anchor}">{name}</text>')
    return f'<div class="chart"><svg viewBox="0 0 {W} {H}" role="img" aria-label="Positioning map">' + "".join(g) + "</svg></div>"
