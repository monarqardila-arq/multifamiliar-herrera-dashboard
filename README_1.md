# Dashboard BIM — Multifamiliar Herrera

Dashboard interactivo con el resumen del modelo BIM (Revit) del proyecto **Multifamiliar Herrera**: edificio de 6 niveles + cubierta.

🔗 **Ver el dashboard en vivo:** `https://tu-usuario.github.io/multifamiliar-herrera-dashboard/`
*(reemplaza con tu URL real una vez publicado en GitHub Pages)*

---

## ¿Qué muestra este dashboard?

- **Resumen**: indicadores generales del modelo (niveles, área construida, ventanales, muros, aparatos sanitarios, ambientes).
- **Puertas y ventanas**: inventario completo por tipo, dimensión y piso.
- **Muros y mobiliario**: área de muros por piso (para presupuesto) y mobiliario sanitario (duchas, sanitarios, lavamanos).
- **Calidad (QA)**: observaciones críticas, importantes y de mejora detectadas en el modelo.

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
├── index.html   → el dashboard completo (HTML + CSS + JS, un solo archivo, sin dependencias externas)
└── README.md    → este archivo
```

---

## Notas técnicas

- Un solo archivo HTML autocontenido: sin librerías externas, sin necesidad de servidor ni build. Funciona con solo abrirlo en cualquier navegador, incluso sin internet.
- Los gráficos son SVG generado con JavaScript nativo (sin dependencias).
- Compatible con cualquier hosting estático (GitHub Pages, Netlify, servidor propio, etc.) si en el futuro se quiere migrar.
