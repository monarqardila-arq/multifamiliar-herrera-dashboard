"""Cartilla de acero y cálculo de concreto por elementos del foso de ascensor (plano EST-02).

Genera:
  - foso_cartilla_html (para imprimir a PDF con un navegador)
  - Foso_Ascensor_Acero_Concreto.xlsx (tablas con fórmulas)

Uso: python generar_foso.py <salida.html> <salida.xlsx>
"""
import math
import sys

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

HTML_OUT, XLSX_OUT = sys.argv[1], sys.argv[2]
PLAN_OUT = sys.argv[3] if len(sys.argv) > 3 else None

DESP_ACERO = 0.05
DESP_CONC = 0.05

# Marca, elemento, barra, separación, long. a distribuir, n fijo, multiplicador, L, forma, tramos, figura
CARTILLA = [
    ("L1", "Losa de fondo — sentido X (2 capas)", "#4", 0.20, 2.08, None, 2, 2.20,
     "Recta con gancho a 180° en ambos extremos", "1.70 + 2 × 0.20", ("H2", 1.70, 0.20)),
    ("L2", "Losa de fondo — sentido Y (2 capas)", "#4", 0.20, 1.68, None, 2, 2.60,
     "Recta con gancho a 180° en ambos extremos", "2.10 + 2 × 0.20", ("H2", 2.10, 0.20)),
    ("M1", "Muros X — horizontales (2 muros)", "#4", 0.20, 1.20, 7, 2, 2.10,
     "En U, patas a 90°", "1.68 + 2 × 0.20", ("U", 1.68, 0.20)),
    ("M2", "Muros Y — horizontales (2 muros)", "#4", 0.20, 1.20, 7, 2, 2.50,
     "En U, patas a 90°", "2.08 + 2 × 0.20", ("U", 2.08, 0.20)),
    ("M3", "Muros X — verticales (2 muros)", "#4", 0.20, 1.30, None, 2, 1.75,
     "Gancho sup. 180°, pata inf. 90°", "1.28 + 0.20 + 0.20", ("HL", 1.28, 0.20)),
    ("M4", "Muros Y — verticales (2 muros)", "#4", 0.20, 1.70, None, 2, 1.75,
     "Gancho sup. 180°, pata inf. 90°", "1.28 + 0.20 + 0.20", ("HL", 1.28, 0.20)),
    ("C1", "Columnas 40×40 — longitudinal (4 col.)", "#4", None, None, 4, 4, 1.80,
     "Gancho sup. 180°, pata inf. 90°", "1.32 + 0.20 + 0.20", ("HL", 1.32, 0.20)),
    ("E1", "Columnas 40×40 — flejes (4 col.)", "#3", 0.075, 1.20, None, 4, 1.50,
     "Fleje cerrado 0.32 × 0.32, ganchos 135°", "4 × 0.32 + ganchos", ("F", 0.32, 0.32)),
]

# Elemento, cantidad, largo, ancho, alto, descuento (m³ total), nota
CONCRETO = [
    ("Losa de fondo e = 0.20", 1, 2.20, 1.80, 0.20, 4 * 0.25 * 0.25 * 0.20,
     "Exterior de muros. Se descuentan las 4 esquinas que ocupan las columnas (0.25 × 0.25 × 0.20 c/u)"),
    ("Muros sentido X e = 0.15", 2, 1.30, 0.15, 1.20, 0, "Entre caras de columnas, sobre la losa (N-1.30 a N-0.10)"),
    ("Muros sentido Y e = 0.15", 2, 1.70, 0.15, 1.20, 0, "Entre caras de columnas, sobre la losa (N-1.30 a N-0.10)"),
    ("Columnas 40 × 40", 4, 0.40, 0.40, 1.40, 0, "Altura completa N-1.50 a N-0.10"),
]
SOLADO = ("Solado de limpieza e = 0.05 (14 MPa, recomendado)", 1, 2.40, 2.00, 0.05)


def nbarras(sep, dist, nfix):
    # Revisión de obra: sin barra adicional en los extremos (allí ya hay barra de columna o de borde)
    return nfix if nfix is not None else math.ceil(round(dist / sep, 6))


rows = []
for m, el, bar, sep, dist, nfix, mult, L, forma, tramos, figdef in CARTILLA:
    n = nbarras(sep, dist, nfix)
    tot = n * mult
    lt = tot * L
    rows.append(dict(m=m, el=el, bar=bar, sep=sep, dist=dist, n=n, mult=mult, tot=tot, L=L,
                     lt=lt, forma=forma, tramos=tramos, fig=figdef))
l4 = sum(r["lt"] for r in rows if r["bar"] == "#4")
l3 = sum(r["lt"] for r in rows if r["bar"] == "#3")

conc = []
for el, c, a, b, h, desc, nota in CONCRETO:
    vu = a * b * h
    vt = c * vu - desc
    conc.append(dict(el=el, c=c, a=a, b=b, h=h, vu=vu, desc=desc, vt=vt, nota=nota))
vtot = sum(x["vt"] for x in conc)


# ------------------------------------------------------------------ figuras
def f(v):
    return f"{v:.1f}"


def fig(kind, a, b, w=200, h=58):
    p = []
    if kind == "H2":
        x0, x1, y, r, ret = 18, w - 18, h / 2 + 8, 5, 30
        p.append(f'<path d="M {f(x0+ret)} {f(y-2*r)} L {f(x0)} {f(y-2*r)} A {r} {r} 0 0 0 {f(x0)} {f(y)} '
                 f'L {f(x1)} {f(y)} A {r} {r} 0 0 0 {f(x1)} {f(y-2*r)} L {f(x1-ret)} {f(y-2*r)}"/>')
        p.append(f'<text x="{f((x0+x1)/2)}" y="{f(y+14)}">{a:.2f}</text>')
        p.append(f'<text x="{f(x0+ret/2)}" y="{f(y-2*r-4)}">{b:.2f}</text>')
        p.append(f'<text x="{f(x1-ret/2)}" y="{f(y-2*r-4)}">{b:.2f}</text>')
    elif kind == "U":
        x0, x1, yb, hk = 18, w - 18, h - 16, 24
        p.append(f'<polyline points="{f(x0)},{f(yb-hk)} {f(x0)},{f(yb)} {f(x1)},{f(yb)} {f(x1)},{f(yb-hk)}"/>')
        p.append(f'<text x="{f((x0+x1)/2)}" y="{f(yb+13)}">{a:.2f}</text>')
        p.append(f'<text x="{f(x0+16)}" y="{f(yb-hk/2+3)}">{b:.2f}</text>')
        p.append(f'<text x="{f(x1-16)}" y="{f(yb-hk/2+3)}">{b:.2f}</text>')
    elif kind == "HL":
        x0, yt, yb, r, ret, foot = w / 2 + 6, 10, h - 10, 4, 18, 40
        p.append(f'<path d="M {f(x0-2*r)} {f(yt+ret)} L {f(x0-2*r)} {f(yt)} A {r} {r} 0 0 1 {f(x0)} {f(yt)} '
                 f'L {f(x0)} {f(yb)} L {f(x0-foot)} {f(yb)}"/>')
        p.append(f'<text x="{f(x0+20)}" y="{f((yt+yb)/2+3)}">{a:.2f}</text>')
        p.append(f'<text x="{f(x0-2*r-26)}" y="{f(yt+ret/2+3)}">{b:.2f}</text>')
        p.append(f'<text x="{f(x0-foot/2)}" y="{f(yb-4)}">{b:.2f}</text>')
    elif kind == "F":
        sz, x0, y0 = 44, w / 2 - 22, 10
        p.append(f'<rect x="{f(x0)}" y="{f(y0)}" width="{sz}" height="{sz}"/>')
        p.append(f'<polyline points="{f(x0+4)},{f(y0)} {f(x0+15)},{f(y0+11)}"/>')
        p.append(f'<polyline points="{f(x0)},{f(y0+4)} {f(x0+11)},{f(y0+15)}"/>')
        p.append(f'<text x="{f(x0+sz/2)}" y="{f(y0+sz+11)}">{a:.2f} × {b:.2f}</text>')
    return f'<svg viewBox="0 0 {w} {h}" class="fig">' + "".join(p) + "</svg>"


# ------------------------------------------------------------------ HTML / PDF
def n2(v):
    return f"{v:,.2f}"


def n3(v):
    return f"{v:,.3f}"


cart_rows = "".join(
    f'<tr><td class="c">{r["m"]}</td><td>{r["el"]}<div class="sub">{r["forma"]}</div></td>'
    f'<td class="figc">{fig(*r["fig"])}</td><td class="c">{r["bar"]}</td>'
    f'<td class="r">{"—" if r["sep"] is None else f"{r["sep"]:.3f}".rstrip("0").rstrip(".")}</td>'
    f'<td class="r">{r["n"]} × {r["mult"]}</td><td class="r b">{r["tot"]}</td>'
    f'<td class="r">{r["L"]:.2f}<div class="sub">{r["tramos"]}</div></td><td class="r b">{n2(r["lt"])}</td></tr>'
    for r in rows)

conc_rows = "".join(
    f'<tr><td>{x["el"]}</td><td class="r">{x["c"]}</td>'
    f'<td class="r">{x["a"]:.2f}</td><td class="r">{x["b"]:.2f}</td><td class="r">{x["h"]:.2f}</td>'
    f'<td class="r">{n3(x["vu"])}</td><td class="r">{"—" if not x["desc"] else n3(x["desc"])}</td>'
    f'<td class="r b">{n3(x["vt"])}</td><td class="r">{n3(x["vt"]*(1+DESP_CONC))}</td></tr>'
    for x in conc)
sol = SOLADO
sol_v = sol[1] * sol[2] * sol[3] * sol[4]


# ------------------------------------------------------------------ PLAN DE CORTE (varilla comercial de 6 m)
VARILLA = 6.00
# Patrón, barra, n° de varillas, piezas [(marca, L)], uso sugerido del retal
PLAN = [
    ("A", "#4", 9, [("M2", 2.50), ("M4", 1.75), ("M4", 1.75)], "Sin retal"),
    ("A2", "#4", 5, [("M2", 2.50), ("M3", 1.75), ("M3", 1.75)], "Sin retal"),
    ("B", "#4", 15, [("L2", 2.60), ("L1", 2.20)], "Burritos para la parrilla superior (1 por retal)"),
    ("C", "#4", 7, [("M1", 2.10), ("M1", 2.10), ("C1", 1.80)], "Sin retal"),
    ("D", "#4", 4, [("L1", 2.20), ("C1", 1.80), ("C1", 1.80)], "Chatarra"),
    ("E", "#4", 3, [("L2", 2.60), ("M3", 1.75)], "Estacas de formaleta y replanteo"),
    ("F", "#4", 1, [("L1", 2.20), ("C1", 1.80), ("M3", 1.75)], "Chatarra"),
    ("G", "#4", 1, [("L1", 2.20), ("L1", 2.20)], "Reserva para reponer una pieza"),
    ("H", "#3", 16, [("E1", 1.50)] * 4, "Sin retal"),
]
_need = {}
for m, el, bar, sep, dist, nfix, mult, L, *_ in CARTILLA:
    _need[(bar, L)] = _need.get((bar, L), 0) + nbarras(sep, dist, nfix) * mult
_got = {}
for pat, bar, nv, pcs, uso in PLAN:
    assert sum(L for _, L in pcs) <= VARILLA + 1e-9, pat
    for _, L in pcs:
        _got[(bar, L)] = _got.get((bar, L), 0) + nv
assert _got == _need, (_got, _need)
VAR = {b: sum(nv for _, bb, nv, _, _ in PLAN if bb == b) for b in ("#4", "#3")}

# ------------------------------------------------------------------ INVENTARIO DE OBRA (varilla de 1/2")
# Fila en la hoja "INVENTARIO DE OBRA", longitud en inventario (m), detalle, uso en el foso, longitud a cortar (m)
INV_USO = [
    (30, 1.05, 'Con gancho de 0.20 m', 'Burrito', 0.75),
    (31, 1.765, '—', 'M3', 1.75),
    (32, 1.765, '—', 'M3', 1.75),
    (35, 1.047, 'Con gancho de 0.20 m', 'Burrito', 0.75),
    (36, 1.062, 'Con gancho de 0.195 m', 'Burrito', 0.75),
    (39, 1.06, 'Con gancho de 0.195 m', 'Burrito', 0.75),
    (40, 0.777, '—', 'Burrito', 0.75),
    (41, 0.767, '—', 'Burrito', 0.75),
    (45, 3.55, 'Con patas de 0.20 m a cada lado', 'L2', 2.60),
    (49, 2.29, 'Con ganchos a 180° de 20.4 cm', 'M3', 1.75),
    (50, 2.265, 'Con ganchos a 160° de 20.2 cm', 'M3', 1.75),
    (54, 2.04, '—', 'C1', 1.80),
    (55, 2.705, '—', 'L2', 2.60),
    (56, 2.352, 'Con pata de 20 cm a un lado', 'M1', 2.10),
    (57, 2.415, '—', 'L1', 2.20),
    (58, 2.33, '—', 'L1', 2.20),
    (59, 1.943, '—', 'C1', 1.80),
    (61, 1.874, '—', 'C1', 1.80),
    (62, 2.267, '—', 'L1', 2.20),
    (63, 2.393, '—', 'L1', 2.20),
    (64, 2.322, '—', 'L1', 2.20),
    (65, 2.724, '—', 'L2', 2.60),
    (66, 2.705, '—', 'L2', 2.60),
    (67, 2.387, '—', 'L1', 2.20),
    (68, 2.371, '—', 'L1', 2.20),
    (69, 2.335, '—', 'L1', 2.20),
    (70, 1.878, '—', 'C1', 1.80),
    (71, 2.335, '—', 'L1', 2.20),
    (72, 2.328, '—', 'L1', 2.20),
    (73, 1.87, '—', 'C1', 1.80),
    (78, 0.932, '—', 'Burrito', 0.75),
    (80, 0.943, '—', 'Burrito', 0.75),
    (81, 2.392, '—', 'L1', 2.20),
    (82, 2.372, '—', 'L1', 2.20),
    (83, 2.467, '—', 'L1', 2.20),
    (84, 2.582, '—', 'M2', 2.50),
    (85, 2.52, '—', 'M2', 2.50),
    (88, 3.095, 'Con ganchos a 180° de 20 cm', 'M2', 2.50),
    (89, 3.11, 'Con ganchos a 180° de 20 cm', 'L2', 2.60),
    (90, 2.39, '—', 'L1', 2.20),
]
# Varillas nuevas de 6 m para lo que no sale del inventario (mínimo exacto: 31 #4)
PLAN_INV = [
    ("A", "#4", 9, [("M2", 2.50), ("M4", 1.75), ("M4", 1.75)], "Sin retal"),
    ("A2", "#4", 2, [("M2", 2.50), ("M3", 1.75), ("M3", 1.75)], "Sin retal"),
    ("B", "#4", 6, [("M1", 2.10), ("M1", 2.10), ("C1", 1.80)], "Sin retal"),
    ("C", "#4", 4, [("L2", 2.60), ("L2", 2.60)], "Estacas de formaleta"),
    ("D", "#4", 3, [("L2", 2.60), ("L1", 2.20)], "Estacas de formaleta y replanteo"),
    ("E", "#4", 2, [("L1", 2.20), ("M3", 1.75), ("M3", 1.75)], "Chatarra"),
    ("F", "#4", 2, [("L1", 2.20), ("C1", 1.80), ("C1", 1.80)], "Chatarra"),
    ("G", "#4", 1, [("L2", 2.60), ("M3", 1.75)], "Reserva para reponer una pieza"),
    ("H", "#4", 1, [("L2", 2.60), ("M1", 2.10)], "Estacas de formaleta"),
    ("I", "#4", 1, [("L1", 2.20), ("C1", 1.80), ("M3", 1.75)], "Chatarra"),
    ("J", "#3", 16, [("E1", 1.50)] * 4, "Sin retal"),
]
_got = {}
for _, L, _, uso, Lc in INV_USO:
    assert Lc <= L + 1e-9
    if uso != "Burrito":
        _got[("#4", Lc)] = _got.get(("#4", Lc), 0) + 1
for pat, bar, nv, pcs, uso in PLAN_INV:
    assert sum(L for _, L in pcs) <= VARILLA + 1e-9, pat
    for _, L in pcs:
        _got[(bar, L)] = _got.get((bar, L), 0) + nv
assert _got == _need, (_got, _need)
# Letras de patrón consecutivas (A, B, C…)
PLAN = [(chr(65 + i),) + t[1:] for i, t in enumerate(PLAN)]
PLAN_INV = [(chr(65 + i),) + t[1:] for i, t in enumerate(PLAN_INV)]
_need_m = {}
for m, el, bar, sep, dist, nfix, mult, L, *_ in CARTILLA:
    _need_m[m] = nbarras(sep, dist, nfix) * mult
for plan_ in (PLAN, PLAN_INV):
    _got_m = {}
    if plan_ is PLAN_INV:
        for _, L, _, uso, Lc in INV_USO:
            if uso != "Burrito":
                _got_m[uso] = _got_m.get(uso, 0) + 1
    for pat, bar, nv, pcs, uso in plan_:
        for mm, _ in pcs:
            _got_m[mm] = _got_m.get(mm, 0) + nv
    assert _got_m == _need_m, (_got_m, _need_m)
VAR_INV = {b: sum(nv for _, bb, nv, _, _ in PLAN_INV if bb == b) for b in ("#4", "#3")}
N_SIL = sum(1 for x in INV_USO if x[3] == "Burrito")
INV_PZ = [x for x in INV_USO if x[3] != "Burrito"]
INV_ML = sum(x[4] for x in INV_USO)


def barra_svg(pcs, w=330, h=22):
    sc = (w - 2) / VARILLA
    x, out = 1, []
    for m, L in pcs:
        ww = L * sc
        out.append(f'<rect x="{x:.1f}" y="2" width="{ww:.1f}" height="{h-4}" class="pz"/>'
                   f'<text x="{x+ww/2:.1f}" y="{h/2+4:.1f}">{m} {L:.2f}</text>')
        x += ww
    rest = VARILLA - sum(L for _, L in pcs)
    if rest > 0.001:
        ww = rest * sc
        out.append(f'<rect x="{x:.1f}" y="2" width="{ww:.1f}" height="{h-4}" class="rt"/>')
        if ww > 26:
            out.append(f'<text x="{x+ww/2:.1f}" y="{h/2+4:.1f}" class="rtt">{rest:.2f}</text>')
    return f'<svg viewBox="0 0 {w} {h}" class="bar6">' + "".join(out) + "</svg>"


def plan_rows_html(plan):
    return "".join(
        f'<tr><td class="c b">{pat}</td><td class="c">{bar}</td><td class="r b">{nv}</td><td>{barra_svg(pcs)}</td>'
        f'<td class="r">{VARILLA - sum(L for _, L in pcs):.2f}</td><td class="r">{nv*(VARILLA - sum(L for _, L in pcs)):.2f}</td><td>{uso}</td></tr>'
        for pat, bar, nv, pcs, uso in plan)


USO_TXT = {
    "L1": "L1 · Losa, sentido X",
    "L2": "L2 · Losa, sentido Y",
    "M1": "M1 · Muro lado X (1.30 m), horizontal",
    "M2": "M2 · Muro lado Y (1.70 m), horizontal",
    "M3": "M3 · Muro lado X (1.30 m), vertical",
    "M4": "M4 · Muro lado Y (1.70 m), vertical",
    "C1": "C1 · Columna, longitudinal",
    "E1": "E1 · Columna, fleje",
    "Burrito": "Burrito · Losa, parrilla superior",
}


def inv_rows_html(items):
    return "".join(
        f'<tr><td class="c">{fila}</td><td class="r">{L:.3f}</td><td>{d}</td><td class="b">{USO_TXT[uso]}</td><td class="r b">{Lc:.2f}</td></tr>'
        for fila, L, d, uso, Lc in items)


half = (len(INV_USO) + 1) // 2
INV_HEAD = ('<thead><tr><th>Fila</th><th class="r">Long. (m)</th><th>Detalle en inventario</th>'
            '<th>Uso</th><th class="r">Cortar (m)</th></tr></thead>')
ret_inv = VAR_INV["#4"] * VARILLA - sum(nv * sum(L for _, L in pcs) for _, b, nv, pcs, _ in PLAN_INV if b == "#4")
PLAN_HTML = f"""
<h2 class="c2">3. Aprovechamiento del inventario de varilla de 1/2"</h2>
<div class="res">
  <div><b>{len(INV_PZ)} piezas</b><span>del foso salen del inventario de obra, más {N_SIL} burritos</span></div>
  <div><b>{n2(INV_ML)} m</b><span>de varilla de 1/2" del inventario que se usan</span></div>
  <div><b>{VAR['#4'] - VAR_INV['#4']} varillas</b><span>#4 de 6 m que se dejan de comprar (de {VAR['#4']} a {VAR_INV['#4']})</span></div>
  <div><b>0</b><span>Varilla #3 en inventario: los flejes se compran completos</span></div>
</div>
<div style="display:grid;grid-template-columns:minmax(0,1fr) minmax(0,1fr);gap:10px;margin-top:8px">
<table>{INV_HEAD}<tbody>{inv_rows_html(INV_USO[:half])}</tbody></table>
<table>{INV_HEAD}<tbody>{inv_rows_html(INV_USO[half:])}</tbody></table>
</div>
<ul class="notas">
  <li>"Fila" es la fila de la hoja INVENTARIO DE OBRA, donde estas varillas están marcadas en naranja.</li>
  <li>En los trozos con patas o ganchos se corta el doblez 5 cm más allá y se usa solo el tramo recto. Las varillas curvadas no se usan: no se deben enderezar.</li>
  <li>No sirven del inventario: la varilla de 1/4", la de 5/8" (salvo aprobación del calculista) ni los flejes de 31.5 × 25 cm y 34 × 31 cm, porque la columna lleva fleje de 32 × 32 cm.</li>
</ul>

<h2 class="c2">4. Plan de corte de las varillas nuevas de 6.00 m</h2>
<div class="res">
  <div><b>{VAR_INV['#4']} varillas</b><span>#4 (1/2") de 6 m para lo que no sale del inventario</span></div>
  <div><b>{VAR_INV['#3']} + 1 varillas</b><span>#3 (3/8") de 6 m, más 1 de reserva</span></div>
  <div><b>{n2(ret_inv)} m</b><span>Retal total de #4 ({ret_inv/(VAR_INV['#4']*VARILLA)*100:.1f} % de lo comprado)</span></div>
  <div><b>{VAR_INV['#4'] + VAR_INV['#3'] + 1} varillas</b><span>Total a pedir</span></div>
</div>
<table style="margin-top:8px">
<thead><tr><th>Patrón</th><th>Barra</th><th class="r">Varillas</th><th>Cortes en cada varilla de 6.00 m (marca y longitud)</th>
<th class="r">Retal c/u (m)</th><th class="r">Retal total (m)</th><th>Uso del retal</th></tr></thead>
<tbody>{plan_rows_html(PLAN_INV)}</tbody>
</table>
<ul class="notas">
  <li>{VAR_INV['#4']} varillas #4 es el mínimo exacto para las piezas que faltan. Sin usar el inventario serían {VAR['#4']}.</li>
  <li>Los burritos ya salen del inventario, así que los retales de 0.80 y 1.20 m quedan para estacas de formaleta y replanteo.</li>
  <li>Cortar primero las piezas largas de cada varilla y marcar cada pieza con su marca (L1, M2…) antes de figurar.</li>
</ul>
"""

HTML = f"""<!doctype html><html lang="es"><head><meta charset="utf-8">
<title>Foso de ascensor — cartilla de acero y concreto</title>
<style>
@page {{ size: Letter landscape; margin: 12mm 12mm 14mm; }}
* {{ box-sizing: border-box; }}
body {{ font: 9pt/1.35 Arial, Helvetica, sans-serif; color: #1a1a1a; margin: 0; padding: 0 2px; }}
h1 {{ font-size: 15pt; margin: 0 0 2px; letter-spacing: .3px; }}
h2 {{ font-size: 11pt; margin: 10px 0 5px; padding-bottom: 3px; border-bottom: 1.5px solid #1a1a1a; text-transform: uppercase; letter-spacing: .4px; }}
.head {{ display: flex; justify-content: space-between; gap: 16px; border-bottom: 2px solid #1a1a1a; padding-bottom: 8px; }}
.head .meta {{ font-size: 8pt; text-align: right; color: #444; line-height: 1.5; }}
.lead {{ color: #444; margin: 4px 0 0; }}
table {{ width: 100%; border-collapse: collapse; }}
th {{ font-size: 7.2pt; text-transform: uppercase; letter-spacing: .3px; text-align: left; background: #e8ebe6; border: 0.6pt solid #9aa09a; padding: 4px 5px; }}
td {{ border: 0.6pt solid #b8bdb7; padding: 1px 5px; vertical-align: middle; }}
td.r, th.r {{ text-align: right; font-variant-numeric: tabular-nums; white-space: nowrap; }}
td.c {{ text-align: center; }}
td.b {{ font-weight: bold; }}
.sub {{ font-size: 7pt; color: #666; }}
tr.tot td {{ font-weight: bold; background: #f3f4f1; }}
tr {{ break-inside: avoid; }}
col.el {{ width: 230px; }}
.figc {{ width: 190px; padding: 2px 6px; }}
svg.fig {{ width: 140px; height: auto; display: block; }}
svg.fig path, svg.fig polyline, svg.fig rect {{ fill: none; stroke: #b3261e; stroke-width: 2.2; stroke-linejoin: round; stroke-linecap: round; }}
svg.fig text {{ font: 9px Arial, sans-serif; fill: #333; text-anchor: middle; }}
svg.bar6 {{ width: 280px; height: auto; display: block; }}
svg.bar6 .pz {{ fill: #fbe3df; stroke: #b3261e; stroke-width: 1; }}
svg.bar6 .rt {{ fill: #e6e6e6; stroke: #999; stroke-width: 1; stroke-dasharray: 3 2; }}
svg.bar6 text {{ font: 8.5px Arial, sans-serif; fill: #1a1a1a; text-anchor: middle; }}
svg.bar6 .rtt {{ fill: #666; }}
.res {{ display: grid; grid-template-columns: repeat(4, 1fr); gap: 6px; margin-top: 8px; }}
h2.c2 {{ break-before: page; }}
.res div {{ border: 0.8pt solid #9aa09a; padding: 6px 8px; }}
.res b {{ display: block; font-size: 13pt; }}
.res span {{ font-size: 7.5pt; color: #555; }}
.notas {{ font-size: 8pt; color: #333; margin: 6px 0 0; padding-left: 16px; }}
.notas li {{ margin-bottom: 2px; }}
.geo {{ font-size: 8.5pt; margin: 2px 0 8px; }}
.avoid {{ break-inside: avoid; }}
</style></head><body>
<h1>FOSO DE ASCENSOR — CARTILLA DE ACERO Y CONCRETO</h1>

<h2>1. Cartilla de despiece del acero de refuerzo</h2>
<table>
<colgroup><col style="width:44px"><col class="el"></colgroup>
<thead><tr><th>Marca</th><th>Elemento / forma</th><th>Figura</th><th>Barra</th><th class="r">Sep. (m)</th>
<th class="r">N° × elem.</th><th class="r">Cant.</th><th class="r">L corte (m)</th><th class="r">Long. total (ml)</th></tr></thead>
<tbody>{cart_rows}
<tr class="tot"><td colspan="8">Barra #4 (1/2") — neto / con 5 % de desperdicio</td><td class="r">{n2(l4)} / {n2(l4*(1+DESP_ACERO))}</td></tr>
<tr class="tot"><td colspan="8">Barra #3 (3/8") — neto / con 5 % de desperdicio</td><td class="r">{n2(l3)} / {n2(l3*(1+DESP_ACERO))}</td></tr>
<tr class="tot"><td colspan="8">Total #4 + #3 — neto / con 5 % de desperdicio</td><td class="r">{n2(l4+l3)} / {n2((l4+l3)*(1+DESP_ACERO))}</td></tr>
</tbody>
</table>

<h2>2. Cálculo de concreto por elementos</h2>
<table>
<colgroup><col style="width:290px"></colgroup>
<thead><tr><th>Elemento</th><th class="r">Cant.</th><th class="r">Largo (m)</th><th class="r">Ancho (m)</th><th class="r">Alto / esp. (m)</th>
<th class="r">Vol. unitario (m³)</th><th class="r">Descuento (m³)</th><th class="r">Vol. neto (m³)</th><th class="r">Con 5 % (m³)</th></tr></thead>
<tbody>{conc_rows}
<tr class="tot"><td colspan="7">Total concreto estructural 21 MPa</td><td class="r">{n3(vtot)}</td><td class="r">{n3(vtot*(1+DESP_CONC))}</td></tr>
<tr><td>{sol[0]}</td><td class="r">{sol[1]}</td><td class="r">{sol[2]:.2f}</td><td class="r">{sol[3]:.2f}</td><td class="r">{sol[4]:.2f}</td>
<td class="r">{n3(sol_v)}</td><td class="r">—</td><td class="r">{n3(sol_v)}</td><td class="r">{n3(sol_v*(1+DESP_CONC))}</td></tr>
</tbody>
</table>
</body></html>"""
open(HTML_OUT, "w").write(HTML)

# Plan de corte en documento aparte: solo tablas
_head = HTML[:HTML.index("</style></head><body>")].replace(
    "<title>Foso de ascensor — cartilla de acero y concreto</title>", "<title>Foso de ascensor — plan de corte</title>")
PLAN_DOC = _head + f"""</style></head><body>
<h1>FOSO DE ASCENSOR — PLAN DE CORTE DE VARILLAS</h1>

<h2>1. Cortes de varillas de 1/2" del inventario de obra</h2>
<div style="display:grid;grid-template-columns:minmax(0,1fr) minmax(0,1fr);gap:10px">
<table>{INV_HEAD}<tbody>{inv_rows_html(INV_USO[:half])}</tbody></table>
<table>{INV_HEAD}<tbody>{inv_rows_html(INV_USO[half:])}</tbody></table>
</div>
<table style="margin-top:6px"><tbody>
<tr class="tot"><td>Total del inventario: {len(INV_PZ)} piezas + {N_SIL} burritos</td><td class="r">{n2(INV_ML)} m</td></tr>
</tbody></table>

<h2 class="c2">2. Cortes de varillas nuevas de 6.00 m</h2>
<table>
<thead><tr><th>Patrón</th><th>Barra</th><th class="r">Varillas</th><th>Cortes en cada varilla de 6.00 m (marca y longitud)</th>
<th class="r">Retal c/u (m)</th><th class="r">Retal total (m)</th><th>Uso del retal</th></tr></thead>
<tbody>{plan_rows_html(PLAN_INV)}
<tr class="tot"><td colspan="2">Total #4 (1/2")</td><td class="r">{VAR_INV['#4']}</td><td colspan="4"></td></tr>
<tr class="tot"><td colspan="2">Total #3 (3/8")</td><td class="r">{VAR_INV['#3']}</td><td colspan="4"></td></tr>
</tbody>
</table>
</body></html>"""
if PLAN_OUT:
    open(PLAN_OUT, "w").write(PLAN_DOC)

# ------------------------------------------------------------------ EXCEL
F = "Arial"
BLUE = Font(name=F, size=10, color="0000FF")
BLK = Font(name=F, size=10)
BOLD = Font(name=F, size=10, bold=True)
HF = Font(name=F, size=10, bold=True, color="FFFFFF")
TIT = Font(name=F, size=14, bold=True)
HFILL = PatternFill("solid", fgColor="1F3A5F")
TOT = PatternFill("solid", fgColor="F2F2F2")
YEL = PatternFill("solid", fgColor="FFFF00")
th = Side(style="thin", color="BFBFBF")
BR = Border(left=th, right=th, top=th, bottom=th)
WR = Alignment(wrap_text=True, vertical="center")
CE = Alignment(horizontal="center", vertical="center", wrap_text=True)


def cell(ws, ref, v, font=BLK, fmt=None, fill=None, al=WR):
    c = ws[ref]
    c.value, c.font, c.border, c.alignment = v, font, BR, al
    if fmt:
        c.number_format = fmt
    if fill:
        c.fill = fill
    return c


def header(ws, r, labels, widths):
    for i, (l, w) in enumerate(zip(labels, widths), 1):
        c = ws.cell(row=r, column=i, value=l)
        c.font, c.fill, c.alignment, c.border = HF, HFILL, CE, BR
        ws.column_dimensions[get_column_letter(i)].width = w


wb = Workbook()

# Parámetros
P = wb.active
P.title = "Parámetros"
P["A1"] = "PARÁMETROS"; P["A1"].font = TIT
P["A2"] = "Celdas en azul sobre amarillo: editables. Todas las tablas se recalculan con ellas."
P["A2"].font = Font(name=F, size=9, italic=True)
header(P, 4, ["Parámetro", "Valor", "Und", "Fuente"], [44, 12, 8, 50])
PAR = [("Desperdicio acero", 0.05, "", "Criterio de obra"), ("Desperdicio concreto", 0.05, "", "Criterio de obra"),
       ("f'c concreto estructural", 21, "MPa", "Supuesto; EST-02 no lo indica"),
       ("fy acero de refuerzo", 420, "MPa", "Supuesto; EST-02 no lo indica")]
for i, (n, v, u, src) in enumerate(PAR, 5):
    cell(P, f"A{i}", n); cell(P, f"B{i}", v, BLUE, "0%" if "Desperdicio" in n else "0.000", YEL)
    cell(P, f"C{i}", u); cell(P, f"D{i}", src)
DA, DC = "Parámetros!$B$5", "Parámetros!$B$6"

# Acero: todas las tablas de varillas en una sola hoja, con el formato de "CÁLCULO DE CANTIDADES DE OBRA"
A = wb.create_sheet("ACERO")
CAL = "Calibri"
def bs(l="dotted", r="dotted", t="dotted", b="dotted"):
    return Border(left=Side(style=l), right=Side(style=r), top=Side(style=t), bottom=Side(style=b))
G_TIT = PatternFill("solid", fgColor="FF6AA84F")
G_SUB = PatternFill("solid", fgColor="FFB6D7A8")
G_TOT = PatternFill("solid", fgColor="FF00FF00")
ORG = PatternFill("solid", fgColor="FFFFC000")
for col, w in zip("ABCDEFGHIJ", [4, 4, 52, 24, 14, 26, 13, 9, 9, 9]):
    A.column_dimensions[col].width = w
A.sheet_view.showGridLines = True


def a(ref, v, bold=False, fmt=None, fill=None, border=None, al=None, color=None):
    c = A[ref]
    c.value = v
    c.font = Font(name=CAL, size=11, bold=bold, color=color)
    if fmt: c.number_format = fmt
    if fill: c.fill = fill
    c.border = border or bs()
    c.alignment = al or Alignment(vertical="center", wrap_text=True)
    return c


def titulo(r, txt):
    a(f"C{r}", txt, True, fill=G_TIT, border=bs("thick", "thick", "thick", "double"),
      al=Alignment(horizontal="center", vertical="center"))
    for col in "DEFGHIJ":
        a(f"{col}{r}", None, fill=G_TIT, border=bs("thin", "thick" if col == "J" else "thin", "thick", "double"))
    A.merge_cells(f"C{r}:J{r}")


def encabezado(r, labels):
    for col, l in zip("CDEFGH", labels):
        a(f"{col}{r}", l, True, border=bs("thick" if col == "C" else "thin", "thin", "medium", "medium"),
          al=Alignment(horizontal="center", vertical="center", wrap_text=True))
    for col in "IJ":
        a(f"{col}{r}", None, border=bs("thin", "thick" if col == "J" else "thin", "medium", "medium"))
    A.merge_cells(f"H{r}:J{r}")


def fila(r, vals, fmts, fill=None, bold=False, borde=("dotted", "dotted"), inputs=()):
    for k, (col, v) in enumerate(zip("CDEFGH", vals)):
        a(f"{col}{r}", v, bold, fmts[k] if k < len(fmts) else None, fill,
          bs("thick" if col == "C" else "dotted", "dotted", *borde),
          Alignment(horizontal="left" if col == "C" else "center", vertical="center", wrap_text=True),
          color="0000FF" if col in inputs else None)
    for col in "IJ":
        a(f"{col}{r}", None, fill=fill, border=bs("dotted", "thick" if col == "J" else "dotted", *borde))
    A.merge_cells(f"H{r}:J{r}")


def total(r, txt, frm, fmt="0.00", fill=G_SUB, extra=None):
    a(f"C{r}", txt, True, fill=fill, border=bs("thick", "thin", "double", "thin" if fill is G_SUB else "thick"))
    for col in "DEFG":
        v = extra.get(col) if extra else None
        a(f"{col}{r}", v, True, fill=fill, border=bs("thin", "thin", "double", "thin" if fill is G_SUB else "thick"),
          fmt=fmt, al=Alignment(horizontal="center"))
    a(f"H{r}", frm, True, fmt, fill, bs("thin", "thin", "double", "thin" if fill is G_SUB else "thick"),
      Alignment(horizontal="center"))
    for col in "IJ":
        a(f"{col}{r}", None, fill=fill, border=bs("thin", "thick" if col == "J" else "thin", "double",
                                                     "thin" if fill is G_SUB else "thick"))
    A.merge_cells(f"H{r}:J{r}")


a("C2", "FOSO DE ASCENSOR — CANTIDADES DE ACERO DE REFUERZO (PLANO EST-02)", True, border=Border(),
  al=Alignment(vertical="center"))
A["C2"].font = Font(name=CAL, size=14, bold=True)
a("C3", "Multifamiliar Herrera · Acero fy 420 MPa · Varilla comercial de 6.00 m · Valores en azul: datos del plano", border=Border())
A["C3"].font = Font(name=CAL, size=10, italic=True)

r = 5
REFS = {}
for barra, nombre in (("#4", 'ML DE ACERO #4 (1/2") — CARTILLA'), ("#3", 'ML DE ACERO #3 (3/8") — CARTILLA')):
    titulo(r, nombre); r += 1
    encabezado(r, ["ELEMENTO", "FORMA (TRAMOS m)", "SEPARACIÓN (m)", "CANTIDAD", "L CORTE (m)", "ML"]); r += 1
    r0 = r
    for m, el, bar, sep, dist, nfix, mult, L, forma, tramos, _ in CARTILLA:
        if bar != barra: continue
        cant = f"=ROUNDUP({dist}/{sep},0)*{mult}" if nfix is None else f"={nfix}*{mult}"
        fila(r, [f"{m} — {el}", tramos, sep if sep else "—", cant, L, f"=F{r}*G{r}"],
             [None, None, "0.000", "0", "0.00", "0.00"], inputs="EG")
        r += 1
    total(r, "TOTAL", f"=SUM(H{r0}:H{r-1})", extra={"F": f"=SUM(F{r0}:F{r-1})"})
    A[f"F{r}"].number_format = "0"
    REFS[barra] = f"H{r}"
    r += 3

titulo(r, "RESUMEN DE METROS LINEALES DE ACERO"); r += 1
fila(r, ['Acero #4 (1/2")', None, None, None, None, f"={REFS['#4']}"], [None]*5 + ["0.00"]); r += 1
fila(r, ['Acero #3 (3/8")', None, None, None, None, f"={REFS['#3']}"], [None]*5 + ["0.00"]); r += 1
total(r, "SUB TOTAL DE ML DE ACERO", f"=H{r-2}+H{r-1}"); sub = r; r += 1
fila(r, ["DESPERDICIO", None, None, None, f"={DA}", f"=H{sub}*G{r}"], [None, None, None, None, "0%", "0.00"]); des = r; r += 1
total(r, "TOTAL DE ML DE ACERO CON DESPERDICIO", f"=H{sub}+H{des}", fill=G_TOT)
REFS["bruto"] = f"H{r}"
r += 3

titulo(r, 'VARILLAS DEL INVENTARIO DE 1/2" QUE SE USAN (MARCADAS EN NARANJA EN EL INVENTARIO)'); r += 1
encabezado(r, ["FILA EN INVENTARIO · DETALLE", "LONGITUD (m)", "USO", "L A CORTAR (m)", "SOBRANTE (m)", "ML A UTILIZAR"]); r += 1
i0 = r
for fila_inv, L, d, uso, Lc in INV_USO:
    fila(r, [f"Fila {fila_inv} · {d}", L, USO_TXT[uso], Lc, f"=D{r}-F{r}", f"=F{r}"],
         [None, "0.000", None, "0.00", "0.00", "0.00"], fill=ORG, inputs="DF")
    r += 1
i1 = r - 1
total(r, "TOTAL A UTILIZAR DEL INVENTARIO", f"=SUM(H{i0}:H{i1})")
r += 3

titulo(r, "PIEZAS #4: CUÁNTAS SALEN DEL INVENTARIO Y CUÁNTAS HAY QUE CORTAR DE VARILLA NUEVA"); r += 1
encabezado(r, ["PIEZA", "L CORTE (m)", "REQUERIDAS", "DEL INVENTARIO", "FALTAN", "ML FALTANTES"]); r += 1
p0 = r
DEM4 = {}
for m, el, bar, sep, dist, nfix, mult, L, *_ in CARTILLA:
    if bar == "#4":
        k = m
        DEM4.setdefault(k, [L, 0, el])
        DEM4[k][1] += nbarras(sep, dist, nfix) * mult
for k, (L, n, el) in DEM4.items():
    nombre = USO_TXT[k]
    fila(r, [nombre, L, n, f'=COUNTIF($E${i0}:$E${i1},"{k} ·*")', f"=E{r}-F{r}", f"=G{r}*D{r}"],
         [None, "0.00", "0", "0", "0", "0.00"])
    r += 1
fila(r, ["Burritos para la parrilla superior de la losa", 0.75, N_SIL, f'=COUNTIF($E${i0}:$E${i1},"Burrito ·*")',
         f"=E{r}-F{r}", f"=G{r}*D{r}"], [None, "0.00", "0", "0", "0", "0.00"])
r += 1
total(r, "TOTAL", f"=SUM(H{p0}:H{r-1})",
      extra={"E": f"=SUM(E{p0}:E{r-1})", "F": f"=SUM(F{p0}:F{r-1})", "G": f"=SUM(G{p0}:G{r-1})"})
for col in "EFG": A[f"{col}{r}"].number_format = "0"
r += 3


def bloque_plan(r, txt, plan, barra):
    titulo(r, txt); r += 1
    encabezado(r, ["PATRÓN (PIEZAS EN CADA VARILLA DE 6.00 m)", "USADO (m)", "RETAL C/U (m)", "USO DEL RETAL",
                   "RETAL TOTAL (m)", "VARILLAS"]); r += 1
    q0 = r
    for pat, bar, nv, pcs, uso in plan:
        if bar != barra: continue
        txt_p = f"{pat}: " + " + ".join(f"{mm} {LL:.2f}" for mm, LL in pcs)
        fila(r, [txt_p, "=" + "+".join(f"{LL:.2f}" for _, LL in pcs), f"={VARILLA:.2f}-D{r}", uso, f"=E{r}*H{r}", nv],
             [None, "0.00", "0.00", None, "0.00", "0"], inputs="H")
        r += 1
    total(r, "TOTAL VARILLAS", f"=SUM(H{q0}:H{r-1})", extra={"G": f"=SUM(G{q0}:G{r-1})"})
    A[f"H{r}"].number_format = "0"
    return r


r = bloque_plan(r, "PLAN DE CORTE — VARILLAS NUEVAS #4 DE 6.00 m (USANDO EL INVENTARIO)", PLAN_INV, "#4")
REFS["v4"] = f"H{r}"; r += 3
r = bloque_plan(r, "PLAN DE CORTE — VARILLAS NUEVAS #3 DE 6.00 m", PLAN_INV, "#3")
REFS["v3"] = f"H{r}"; r += 3

titulo(r, "RESUMEN DE COMPRA DE VARILLAS DE 6.00 m"); r += 1
fila(r, ['Varilla #4 (1/2")', None, None, None, None, f"={REFS['v4']}"], [None]*5 + ["0"]); c4 = r; r += 1
fila(r, ['Varilla #3 (3/8")', None, None, None, None, f"={REFS['v3']}"], [None]*5 + ["0"]); r += 1
fila(r, ['Reserva varilla #3 (3/8")', None, None, None, None, 1], [None]*5 + ["0"], inputs="H"); r += 1
total(r, "TOTAL VARILLAS A COMPRAR", f"=SUM(H{c4}:H{r-1})", fmt="0", fill=G_TOT)
REFS["compra"] = f"H{r}"
r += 3

r = bloque_plan(r, "REFERENCIA — PLAN DE CORTE SIN USAR EL INVENTARIO (VARILLAS #4)", PLAN, "#4")
r += 2
for n_ in ["Las patas y ganchos se figuran después de cortar cada pieza a su L de corte.",
           "En los trozos del inventario con dobleces se corta el doblez 5 cm más allá y se usa solo el tramo recto.",
           "Las varillas curvadas del inventario no se usan: no se deben enderezar."]:
    A[f"C{r}"] = "• " + n_; A[f"C{r}"].font = Font(name=CAL, size=10); r += 1

# Concreto
K = wb.create_sheet("Concreto por elementos")
K["A1"] = "CÁLCULO DE CONCRETO POR ELEMENTOS — FOSO DE ASCENSOR (EST-02)"; K["A1"].font = TIT
K["A2"] = ("Foso: luz libre 1.50 × 1.90 m · profundidad útil 1.20 m (N-0.10 a N-1.30) · exterior de muros 1.80 × 2.20 m · "
           "con columnas 2.10 × 2.50 m · altura total 1.40 m (N-0.10 a N-1.50).")
K["A2"].font = Font(name=F, size=9, italic=True)
header(K, 4, ["Elemento", "Cant.", "Largo (m)", "Ancho (m)", "Alto / espesor (m)", "Vol. unitario (m³)",
              "Descuento (m³)", "Vol. neto (m³)", "Con desperdicio (m³)", "Nota"],
       [34, 7, 10, 10, 12, 13, 12, 12, 14, 60])
k0 = 5
for i, (el, c, a, b, h, desc, nota) in enumerate(CONCRETO):
    r = k0 + i
    cell(K, f"A{r}", el); cell(K, f"B{r}", c, BLUE, "0")
    cell(K, f"C{r}", a, BLUE, "0.00"); cell(K, f"D{r}", b, BLUE, "0.00"); cell(K, f"E{r}", h, BLUE, "0.00")
    cell(K, f"F{r}", f"=C{r}*D{r}*E{r}", fmt="0.000")
    cell(K, f"G{r}", "=4*0.25*0.25*E5" if desc else 0, fmt='0.000;-0.000;"-"')
    cell(K, f"H{r}", f"=B{r}*F{r}-G{r}", BOLD, "0.000")
    cell(K, f"I{r}", f"=H{r}*(1+{DC})", fmt="0.000")
    cell(K, f"J{r}", nota)
k1 = k0 + len(CONCRETO) - 1
r = k1 + 1
cell(K, f"A{r}", "TOTAL CONCRETO ESTRUCTURAL", BOLD, fill=TOT)
for col in "BCDEFG":
    cell(K, f"{col}{r}", None, fill=TOT)
cell(K, f"H{r}", f"=SUM(H{k0}:H{k1})", BOLD, "0.000", TOT)
cell(K, f"I{r}", f"=SUM(I{k0}:I{k1})", BOLD, "0.000", TOT)
cell(K, f"J{r}", "Concreto f'c según Parámetros", fill=TOT)
KT = r
r += 2
cell(K, f"A{r}", SOLADO[0]); cell(K, f"B{r}", SOLADO[1], BLUE, "0")
cell(K, f"C{r}", SOLADO[2], BLUE, "0.00"); cell(K, f"D{r}", SOLADO[3], BLUE, "0.00"); cell(K, f"E{r}", SOLADO[4], BLUE, "0.00")
cell(K, f"F{r}", f"=C{r}*D{r}*E{r}", fmt="0.000"); cell(K, f"G{r}", 0, fmt='0.000;-0.000;"-"')
cell(K, f"H{r}", f"=B{r}*F{r}", BOLD, "0.000"); cell(K, f"I{r}", f"=H{r}*(1+{DC})", fmt="0.000")
cell(K, f"J{r}", "No dibujado en EST-02. Losa + 0.10 m por lado. Va aparte del concreto estructural.")

# Resumen
R = wb.create_sheet("Resumen", 0)
R["A1"] = "FOSO DE ASCENSOR — RESUMEN DE ACERO Y CONCRETO"; R["A1"].font = TIT
R["A2"] = "Multifamiliar Herrera · Plano EST-02 v.01 (19/06/2026). Valores enlazados a las otras hojas."
R["A2"].font = Font(name=F, size=9, italic=True)
header(R, 4, ["Concepto", "Neto", "Con desperdicio / total", "Und"], [50, 14, 18, 8])
GRN = Font(name=F, size=10, color="008000")
RES = [("Acero #4 (1/2\")", f"=ACERO!{REFS['#4']}", f"=B5*(1+{DA})", "ml"),
       ("Acero #3 (3/8\")", f"=ACERO!{REFS['#3']}", f"=B6*(1+{DA})", "ml"),
       ("Acero total", "=B5+B6", "=C5+C6", "ml"),
       ("Concreto estructural 21 MPa", f"='Concreto por elementos'!H{KT}", f"='Concreto por elementos'!I{KT}", "m³"),
       ("Varillas #4 de 6 m a comprar (usando inventario)", None, f"=ACERO!{REFS['v4']}", "und"),
       ("Varillas #3 de 6 m a comprar (incluye 1 de reserva)", None, f"=ACERO!{REFS['v3']}+1", "und")]
for i, (n, a_, b, u) in enumerate(RES, 5):
    cell(R, f"A{i}", n, BOLD if "total" in n.lower() else BLK)
    cell(R, f"B{i}", a_, GRN, "#,##0.00" if u == "ml" else "0.000")
    cell(R, f"C{i}", b, BOLD, "#,##0.00" if u == "ml" else ("0" if u == "und" else "0.000"))
    cell(R, f"D{i}", u, al=CE)

for ws in wb.worksheets:
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
wb.save(XLSX_OUT)
print(f"varillas {VAR} / con inventario {VAR_INV} | #4 {l4:.2f} ml | #3 {l3:.2f} ml | neto {l4+l3:.2f} | bruto {(l4+l3)*1.05:.2f} | conc {vtot:.3f} / {vtot*1.05:.3f}")
