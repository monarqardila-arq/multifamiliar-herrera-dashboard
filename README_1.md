# Dashboard BIM — Multifamiliar Herrera

Dashboard interactivo con el resumen del modelo BIM (Revit) del proyecto **Multifamiliar Herrera**: edificio de 6 niveles + cubierta.

🔗 **Ver el dashboard en vivo:** https://monarqardila-arq.github.io/multifamiliar-herrera-dashboard/

---

## ¿Qué muestra este dashboard?

- **Planos por piso**: selector interactivo P1–P5 — muestra la planta real de ese nivel junto con un resumen sincronizado (área, muros, puertas, ventanales, ambientes, baños).
- **Resumen**: indicadores generales del modelo (niveles, área construida, ventanales, muros, aparatos sanitarios, ambientes).
- **Carpintería**: puertas y ventanales agrupados por función y medida (no por marca/familia), con desglose por piso.
- **Muros y baños**: área de muros de mampostería por piso, y mobiliario sanitario completo (sanitarios, duchas, regaderas, lavamanos, mamparas).
- **Calidad (QA)**: observaciones pendientes, y un listado de todo lo que ya se corrigió en esta revisión.

> ⚠️ Este dashboard **no incluye presupuesto** (costos ni precios unitarios) — solo cantidades y estado del modelo. El presupuesto se maneja como documento aparte.

---

## Origen de los datos

Toda la información se extrajo por **consulta directa al modelo activo en Autodesk Revit** (`MULTIF H - RVT.rvt`), no son datos estimados a mano. Fecha del último corte de datos: ver esquina superior derecha del dashboard.

---

## Cómo actualizar este dashboard

Los datos están "quemados" en el archivo `index.html` (no se conectan en vivo a Revit). Para actualizarlos:

1. Vuelve a consultar el modelo en Revit (directamente o con ayuda de un asistente BIM).
2. Genera un `index.html` nuevo con los datos actualizados.
3. En este repositorio, entra a `index.html` → ícono de lápiz (editar) o **Add file → Upload files** para reemplazarlo.
4. Confirma el cambio (**Commit changes**). GitHub Pages actualiza la página pública en menos de 1 minuto — no hay que repetir la configuración.

---

## Estructura del repositorio

```
├── index.html      → el dashboard completo (HTML + CSS + JS, sin dependencias externas)
├── imgplants/       → plantas arquitectónicas (P1.jpg – P5.jpg) usadas en la pestaña "Planos por piso"
└── README.md        → este archivo
```

> Si mueves o renombras la carpeta `imgplants`, hay que actualizar las rutas `img:'imgplants/P1.jpg'` etc. dentro del `index.html` (sección `const PLANOS`), o las imágenes dejan de cargar.

---

## Notas técnicas

- Un solo archivo HTML autocontenido: sin librerías externas, sin necesidad de servidor ni build. Funciona con solo abrirlo en cualquier navegador, incluso sin internet.
- Los gráficos son SVG generado con JavaScript nativo (sin dependencias).
- Compatible con cualquier hosting estático (GitHub Pages, Netlify, servidor propio, etc.) si en el futuro se quiere migrar.
