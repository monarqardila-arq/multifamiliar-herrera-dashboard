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

PESO = {"#3": 0.560, "#4": 0.994}
DESP_ACERO = 0.05
DESP_CONC = 0.05

# Marca, elemento, barra, separación, long. a distribuir, n fijo, multiplicador, L, forma, tramos, figura
CARTILLA = [
    ("L1", "Losa de fondo — sentido X (2 capas)", "#4", 0.20, 2.08, None, 2, 2.20,
     "Recta con gancho a 180° en ambos extremos", "1.70 + 2 × 0.20", ("H2", 1.70, 0.20)),
    ("L2", "Losa de fondo — sentido Y (2 capas)", "#4", 0.20, 1.68, None, 2, 2.60,
     "Recta con gancho a 180° en ambos extremos", "2.10 + 2 × 0.20", ("H2", 2.10, 0.20)),
    ("M1", "Muros X — horizontales (2 muros)", "#4", 0.20, 1.20, None, 2, 2.10,
     "En U, patas a 90°", "1.68 + 2 × 0.20", ("U", 1.68, 0.20)),
    ("M2", "Muros Y — horizontales (2 muros)", "#4", 0.20, 1.20, None, 2, 2.50,
     "En U, patas a 90°", "2.08 + 2 × 0.20", ("U", 2.08, 0.20)),
    ("M3", "Muros X — verticales (2 muros)", "#4", 0.20, 1.30, None, 2, 1.75,
     "Gancho sup. 180°, pata inf. 90°", "1.28 + 0.20 + 0.20", ("HL", 1.28, 0.20)),
    ("M4", "Muros Y — verticales (2 muros)", "#4", 0.20, 1.70, None, 2, 1.75,
     "Gancho sup. 180°, pata inf. 90°", "1.28 + 0.20 + 0.20", ("HL", 1.28, 0.20)),
    ("C1", "Columnas 40×40 — longitudinal (4 col.)", "#4", None, None, 4, 4, 1.80,
     "Gancho sup. 180°, pata inf. 90°", "1.32 + 0.20 + 0.20", ("HL", 1.32, 0.20)),
    ("E1", "Columnas 40×40 — flejes (4 col.)", "#3", 0.075, 1.40, None, 4, 1.50,
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
    return nfix if nfix is not None else math.ceil(round(dist / sep, 6)) + 1


rows = []
for m, el, bar, sep, dist, nfix, mult, L, forma, tramos, figdef in CARTILLA:
    n = nbarras(sep, dist, nfix)
    tot = n * mult
    lt = tot * L
    kg = lt * PESO[bar]
    rows.append(dict(m=m, el=el, bar=bar, sep=sep, dist=dist, n=n, mult=mult, tot=tot, L=L,
                     lt=lt, kg=kg, forma=forma, tramos=tramos, fig=figdef))
k4 = sum(r["kg"] for r in rows if r["bar"] == "#4")
k3 = sum(r["kg"] for r in rows if r["bar"] == "#3")
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
    f'<td class="r">{r["L"]:.2f}<div class="sub">{r["tramos"]}</div></td><td class="r">{n2(r["lt"])}</td>'
    f'<td class="r">{PESO[r["bar"]]:.3f}</td><td class="r b">{n2(r["kg"])}</td></tr>'
    for r in rows)

conc_rows = "".join(
    f'<tr><td>{x["el"]}<div class="sub">{x["nota"]}</div></td><td class="r">{x["c"]}</td>'
    f'<td class="r">{x["a"]:.2f}</td><td class="r">{x["b"]:.2f}</td><td class="r">{x["h"]:.2f}</td>'
    f'<td class="r">{n3(x["vu"])}</td><td class="r">{"—" if not x["desc"] else n3(x["desc"])}</td>'
    f'<td class="r b">{n3(x["vt"])}</td><td class="r">{n3(x["vt"]*(1+DESP_CONC))}</td></tr>'
    for x in conc)
sol = SOLADO
sol_v = sol[1] * sol[2] * sol[3] * sol[4]

HTML = f"""<!doctype html><html lang="es"><head><meta charset="utf-8">
<title>Foso de ascensor — cartilla de acero y concreto</title>
<style>
@page {{ size: Letter landscape; margin: 12mm 12mm 14mm; }}
* {{ box-sizing: border-box; }}
body {{ font: 9pt/1.35 Arial, Helvetica, sans-serif; color: #1a1a1a; margin: 0; }}
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
<div class="head">
  <div>
    <h1>FOSO DE ASCENSOR — CARTILLA DE ACERO Y CONCRETO</h1>
    <p class="lead">Multifamiliar Herrera · Puerto Berrío, Antioquia</p>
  </div>
  <div class="meta">Base: plano estructural EST-02 · v.01 · 19/06/2026<br>Concreto f'c = 21 MPa · Acero fy = 420 MPa (NTC 2289)<br>Cotas y longitudes en metros</div>
</div>

<div class="res avoid">
  <div><b>{n2(k4*(1+DESP_ACERO))} kg</b><span>Acero #4 (1/2"), con 5 % de desperdicio</span></div>
  <div><b>{n2(k3*(1+DESP_ACERO))} kg</b><span>Acero #3 (3/8"), con 5 % de desperdicio</span></div>
  <div><b>{n2((k4+k3)*(1+DESP_ACERO))} kg</b><span>Acero total, con desperdicio</span></div>
  <div><b>{n3(vtot*(1+DESP_CONC))} m³</b><span>Concreto 21 MPa, con 5 % de desperdicio</span></div>
</div>

<h2>1. Cartilla de despiece del acero de refuerzo</h2>
<table>
<colgroup><col style="width:44px"><col class="el"></colgroup>
<thead><tr><th>Marca</th><th>Elemento / forma</th><th>Figura</th><th>Barra</th><th class="r">Sep. (m)</th>
<th class="r">N° × elem.</th><th class="r">Cant.</th><th class="r">L corte (m)</th><th class="r">Long. total (m)</th>
<th class="r">kg/m</th><th class="r">Peso (kg)</th></tr></thead>
<tbody>{cart_rows}
<tr class="tot"><td colspan="8">Subtotal #4 (1/2")</td><td class="r">{n2(l4)}</td><td></td><td class="r">{n2(k4)}</td></tr>
<tr class="tot"><td colspan="8">Subtotal #3 (3/8")</td><td class="r">{n2(l3)}</td><td></td><td class="r">{n2(k3)}</td></tr>
<tr class="tot"><td colspan="10">Total neto</td><td class="r">{n2(k4+k3)}</td></tr>
<tr class="tot"><td colspan="10">Total con 5 % de desperdicio</td><td class="r">{n2((k4+k3)*(1+DESP_ACERO))}</td></tr>
</tbody>
</table>
<ul class="notas">
  <li>N° de barras = longitud a distribuir ÷ separación, redondeado hacia arriba, + 1. Losa X: 2.08 m; losa Y: 1.68 m; filas de muro: 1.20 m; flejes: 1.40 m.</li>
  <li>L corte es la longitud indicada en el plano. Es mayor que la suma de los tramos porque incluye el doblez de los ganchos.</li>
  <li>Las patas de los verticales de muro (0.20 m) entran a la losa entre las dos capas. La pata inferior de las columnas va hacia afuera.</li>
  <li>Recubrimientos según EST-02: 0.05 m en la losa, 0.062 y 0.075 m en los muros.</li>
</ul>

<div class="avoid">
<h2 class="c2">2. Cálculo de concreto por elementos</h2>
<p class="geo">Foso: luz libre de 1.50 × 1.90 m, profundidad útil de 1.20 m (N-0.10 a N-1.30). Exterior de muros: 1.80 × 2.20 m. Con columnas: 2.10 × 2.50 m. Altura total: 1.40 m (N-0.10 a N-1.50).</p>
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
<ul class="notas">
  <li>Cuantía de referencia: {(k4+k3)/vtot:,.1f} kg de acero por m³ de concreto (valores netos).</li>
  <li>El solado no está dibujado en EST-02. Se recomienda bajo la losa en contacto con el suelo y se da aparte del concreto estructural.</li>
  <li>El plano no indica f'c ni fy. Se asumieron 21 MPa y 420 MPa: confirmar con el calculista.</li>
</ul>
</div>
</body></html>"""
open(HTML_OUT, "w").write(HTML)

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
PAR = [("Peso barra #3 (3/8\")", 0.560, "kg/m", "NTC 2289"), ("Peso barra #4 (1/2\")", 0.994, "kg/m", "NTC 2289"),
       ("Desperdicio acero", 0.05, "", "Criterio de obra"), ("Desperdicio concreto", 0.05, "", "Criterio de obra"),
       ("f'c concreto estructural", 21, "MPa", "Supuesto; EST-02 no lo indica"),
       ("fy acero de refuerzo", 420, "MPa", "Supuesto; EST-02 no lo indica")]
for i, (n, v, u, src) in enumerate(PAR, 5):
    cell(P, f"A{i}", n); cell(P, f"B{i}", v, BLUE, "0%" if "Desperdicio" in n else "0.000", YEL)
    cell(P, f"C{i}", u); cell(P, f"D{i}", src)
PW3, PW4, DA, DC = "Parámetros!$B$5", "Parámetros!$B$6", "Parámetros!$B$7", "Parámetros!$B$8"

# Cartilla
C = wb.create_sheet("Cartilla acero")
C["A1"] = "CARTILLA DE DESPIECE — ACERO DE REFUERZO FOSO DE ASCENSOR (EST-02)"; C["A1"].font = TIT
C["A2"] = "Valores en azul tomados del plano EST-02 (confirmados en obra). N° de barras = REDONDEAR.MAS(distribución ÷ separación) + 1."
C["A2"].font = Font(name=F, size=9, italic=True)
header(C, 4, ["Marca", "Elemento", "Forma", "Tramos (m)", "Barra", "Separación (m)", "Long. a distribuir (m)",
              "N° por elemento", "Elementos", "Cant. total", "L corte (m)", "Long. total (m)", "kg/m", "Peso (kg)"],
       [7, 36, 30, 18, 7, 11, 12, 10, 10, 10, 10, 12, 8, 11])
r0 = 5
for i, (m, el, bar, sep, dist, nfix, mult, L, forma, tramos, _) in enumerate(CARTILLA):
    r = r0 + i
    cell(C, f"A{r}", m, BOLD, al=CE); cell(C, f"B{r}", el); cell(C, f"C{r}", forma); cell(C, f"D{r}", tramos)
    cell(C, f"E{r}", bar, al=CE)
    cell(C, f"F{r}", sep, BLUE, "0.000"); cell(C, f"G{r}", dist, BLUE, "0.00")
    cell(C, f"H{r}", nfix if nfix is not None else f"=ROUNDUP(ROUND(G{r}/F{r},6),0)+1",
         BLUE if nfix is not None else BLK, "0")
    cell(C, f"I{r}", mult, BLUE, "0")
    cell(C, f"J{r}", f"=H{r}*I{r}", BOLD, "0")
    cell(C, f"K{r}", L, BLUE, "0.00")
    cell(C, f"L{r}", f"=J{r}*K{r}", fmt="#,##0.00")
    cell(C, f"M{r}", f'=IF(E{r}="#4",{PW4},{PW3})', Font(name=F, size=10, color="008000"), "0.000")
    cell(C, f"N{r}", f"=L{r}*M{r}", BOLD, "#,##0.00")
r1 = r0 + len(CARTILLA) - 1
r = r1 + 1
TOTS = {}
for lbl, frm_l, frm_k, key in [
    ("Subtotal #4 (1/2\")", f'=SUMIF(E{r0}:E{r1},"#4",L{r0}:L{r1})', f'=SUMIF(E{r0}:E{r1},"#4",N{r0}:N{r1})', "k4"),
    ("Subtotal #3 (3/8\")", f'=SUMIF(E{r0}:E{r1},"#3",L{r0}:L{r1})', f'=SUMIF(E{r0}:E{r1},"#3",N{r0}:N{r1})', "k3"),
]:
    cell(C, f"A{r}", lbl, BOLD, fill=TOT); C.merge_cells(f"A{r}:K{r}")
    cell(C, f"L{r}", frm_l, BOLD, "#,##0.00", TOT); cell(C, f"M{r}", None, fill=TOT); cell(C, f"N{r}", frm_k, BOLD, "#,##0.00", TOT)
    TOTS[key] = r
    r += 1
cell(C, f"A{r}", "Total neto", BOLD, fill=TOT); C.merge_cells(f"A{r}:M{r}")
cell(C, f"N{r}", f"=N{TOTS['k4']}+N{TOTS['k3']}", BOLD, "#,##0.00", TOT); TOTS["neto"] = r
r += 1
cell(C, f"A{r}", "Total con desperdicio", BOLD, fill=TOT); C.merge_cells(f"A{r}:M{r}")
cell(C, f"N{r}", f"=N{TOTS['neto']}*(1+{DA})", BOLD, "#,##0.00", TOT); TOTS["bruto"] = r
C.freeze_panes = "C5"

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
r += 2
cell(K, f"A{r}", "Cuantía de referencia (kg acero / m³ concreto, neto)", BOLD)
cell(K, f"H{r}", f"='Cartilla acero'!N{TOTS['neto']}/H{KT}", BOLD, "0.0")

# Resumen
R = wb.create_sheet("Resumen", 0)
R["A1"] = "FOSO DE ASCENSOR — RESUMEN DE ACERO Y CONCRETO"; R["A1"].font = TIT
R["A2"] = "Multifamiliar Herrera · Plano EST-02 v.01 (19/06/2026). Valores enlazados a las otras hojas."
R["A2"].font = Font(name=F, size=9, italic=True)
header(R, 4, ["Concepto", "Neto", "Con desperdicio", "Und"], [44, 14, 16, 8])
GRN = Font(name=F, size=10, color="008000")
RES = [("Acero #4 (1/2\")", f"='Cartilla acero'!N{TOTS['k4']}", f"=B5*(1+{DA})", "kg"),
       ("Acero #3 (3/8\")", f"='Cartilla acero'!N{TOTS['k3']}", f"=B6*(1+{DA})", "kg"),
       ("Acero total", "=B5+B6", "=C5+C6", "kg"),
       ("Concreto estructural 21 MPa", f"='Concreto por elementos'!H{KT}", f"='Concreto por elementos'!I{KT}", "m³")]
for i, (n, a, b, u) in enumerate(RES, 5):
    cell(R, f"A{i}", n, BOLD if "total" in n.lower() else BLK)
    cell(R, f"B{i}", a, GRN, "#,##0.00" if u == "kg" else "0.000")
    cell(R, f"C{i}", b, BOLD, "#,##0.00" if u == "kg" else "0.000")
    cell(R, f"D{i}", u, al=CE)

for ws in wb.worksheets:
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
wb.save(XLSX_OUT)
print(f"#4 {k4:.2f} kg | #3 {k3:.2f} kg | neto {k4+k3:.2f} | bruto {(k4+k3)*1.05:.2f} | conc {vtot:.3f} / {vtot*1.05:.3f}")
