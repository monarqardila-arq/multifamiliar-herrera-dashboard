# Cuadro de cantidades — Multifamiliar Herrera

**Archivo de trabajo:** [`Cantidades_Multifamiliar_Herrera.xlsx`](Cantidades_Multifamiliar_Herrera.xlsx). Todas las cantidades derivadas son fórmulas: si cambias un dato o un supuesto, el libro se recalcula solo.

| Hoja | Contenido |
|---|---|
| **Resumen** | Indicadores principales enlazados |
| **Cantidades** | Cuadro por capítulos: ítem, descripción, especificación técnica, unidad, cantidad P1–P5, total, fuente y observaciones |
| **Datos Revit** | Datos de entrada medidos en el modelo `MULTIF H - RVT.rvt` (en azul) |
| **Foso ascensor** | Memoria de cálculo del plano estructural EST-02: geometría, concreto, excavación, formaleta y cartilla de despiece del acero |
| **Supuestos** | Rendimientos, pesos y desperdicios editables (en amarillo) |
| **Observaciones** | Inconsistencias encontradas y datos pendientes |

**Alcance:** niveles P1–P5 terminados en el modelo y el foso del ascensor (EST-02). P6, P7 y la cubierta no se incluyen porque siguen en diseño. No incluye costos: el presupuesto se lleva aparte.

**Fuente de cada cantidad:** R = medida en Revit · E = plano EST-02 · D = derivada por fórmula · P = por confirmar.

---

## 1. Preliminares
| Ítem | Descripción | Und | Cantidad |
|---|---|---|---|
| 1.1 | Localización, trazado y replanteo (área cubierta P1–P5 + exterior) | m² | 603.90 |
| 1.2 | Replanteo del foso de ascensor (2.10 × 2.50 m) | m² | 5.25 |

## 2. Foso de ascensor (EST-02, A3 Ingeniería, v.01, 19/06/2026)
Geometría: luz libre de 1.50 × 1.90 m, muros de concreto e = 0.15 m, 4 columnas de 40×40, losa de fondo e = 0.20 m, de N-0.10 a N-1.50.

| Ítem | Descripción | Especificación | Und | Cantidad |
|---|---|---|---|---|
| 2.1 | Excavación manual | Hasta N-1.55, sobreancho de 0.30 m. Hay zapatas vecinas en los ejes A y B: no descalzarlas | m³ | 12.14 |
| 2.2 | Solado de limpieza e = 5 cm | Concreto 14 MPa (recomendado) | m³ | 0.25 |
| 2.3 | Losa de fondo e = 0.20 | Concreto 21 MPa + impermeabilizante integral | m³ | 0.78 |
| 2.4 | Muros e = 0.15, h = 1.20 | Concreto 21 MPa + impermeabilizante integral | m³ | 1.13 |
| 2.5 | Columnas 40×40, h = 1.40 | Concreto 21 MPa | m³ | 0.94 |
| 2.6 | Acero #4 | fy 420 MPa, NTC 2289 | kg | 333.8 |
| 2.7 | Acero #3 (flejes @0.075) | fy 420 MPa | kg | 70.6 |
| 2.8 | Formaleta | Muros (2 caras) + columnas | m² | 22.72 |
| 2.9 | Impermeabilización interior | Mortero cementicio + mediacaña (recomendado) | m² | 11.01 |
| 2.10 | Relleno compactado | Capas de ≤ 0.20 m al 95 % del Proctor modificado | m³ | 5.81 |
| 2.11 | Retiro de sobrantes | Volumen suelto | m³ | 8.23 |
| 2.12 | Platina base A36 300×300×16 mm | 45.2 kg en total | un | 4 |
| 2.13 | Pernos de anclaje Ø5/8" | Embebidos 314 mm, 4 por platina | un | 16 |
| 2.14 | Soldadura de filete 5/16" | E70XX | ml | 4.6 |
| 2.15 | Perfil HE 200 | Longitud pendiente (altura del ducto) | un | 4 |

Las cantidades de concreto y acero incluyen 5 % de desperdicio.

## 3–6. Mampostería, pañetes, pisos y cielos
| Ítem | Descripción | Und | P1 | P2 | P3 | P4 | P5 | Total |
|---|---|---|---|---|---|---|---|---|
| 3.1 | Muro de mampostería de arcilla (área neta) | m² | 339.40 | 251.10 | 256.42 | 256.08 | 256.15 | **1,359.15** |
| 3.2 | Unidades de mampostería, bloque 10×20×30 (ref.) | un | 5,488 | 4,060 | 4,146 | 4,141 | 4,142 | **21,977** |
| 3.3 | Mortero de pega 1:4 | m³ | 4.07 | 3.01 | 3.08 | 3.07 | 3.07 | **16.31** |
| 4.1 | Pañete, ambas caras (bruto) | m² | 678.80 | 502.20 | 512.84 | 512.16 | 512.30 | **2,718.30** |
| 4.2 | Estuco + vinilo tipo 1, 3 manos (bruto) | m² | 678.80 | 502.20 | 512.84 | 512.16 | 512.30 | **2,718.30** |
| 4.3 | Dinteles | ml | 18.07 | 14.52 | 21.41 | 21.41 | 21.41 | **96.83** |
| 5.1 | Alistado de piso e = 4 cm | m² | 128.03 | 110.53 | 109.91 | 110.04 | 108.30 | **566.81** |
| 5.2 | Mortero de alistado | m³ | 5.12 | 4.42 | 4.40 | 4.40 | 4.33 | **22.67** |
| 5.3 | Acabado de piso interior (tipo por definir) | m² | 128.03 | 110.53 | 109.91 | 110.04 | 108.30 | **566.81** |
| 5.4 | Piso exterior / andén | m² | 37.09 | – | – | – | – | **37.09** |
| 6.1 | Cielorraso (≈ área cubierta) | m² | 128.03 | 110.53 | 109.91 | 110.04 | 108.30 | **566.81** |

## 7. Puertas
| Ítem | Función | Medida (m) | P1 | P2 | P3 | P4 | P5 | Total |
|---|---|---|---|---|---|---|---|---|
| 7.1 | Puerta de garaje | 2.40 × 2.10 | 1 | – | – | – | – | 1 |
| 7.2 | Principal de acceso | 0.90 × 2.10 | 1 | – | – | – | – | 1 |
| 7.3 | Acceso secundario | 0.915 × 2.134 | 1 | – | – | – | – | 1 |
| 7.4 | Corrediza social (balcón) | 1.20 × 2.10 | 1 | – | – | – | – | 1 |
| 7.5 | Corrediza de habitación/clóset | 0.90 × 2.10 | 3 | – | – | – | – | 3 |
| 7.6 | Corrediza de baño/clóset menor | 0.70 × 2.10 | – | 3 | 2 | 2 | 2 | 9 |
| 7.7 | Batiente de habitación | 0.80 × 2.10 | 2 | 1 | – | – | – | 3 |
| 7.8 | Batiente de habitación | 0.90 × 2.10 | 2 | – | 3 | 3 | 3 | 11 |
| 7.9 | Batiente de baño/clóset | 0.70 × 2.10 | – | 1 | 1 | 1 | 1 | 4 |
| 7.10 | Principal de apartamento | 1.00 × 2.10 | – | 1 | – | – | – | 1 |
| 7.11 | Puerta-ventanal corrediza (balcón) | 1.60 × 2.30 | – | – | 1 | 1 | 1 | 3 |
| 7.12 | Corrediza de vestier (Häfele) | ≈7.73 m² de hoja | – | 1 | 1 | 1 | 1 | 4 |
| | **Total** | | 11 | 7 | 8 | 8 | 8 | **42** |

## 8. Ventanería
| Ítem | Función | Medida (m) | Especificación | P1 | P2 | P3 | P4 | P5 | Total |
|---|---|---|---|---|---|---|---|---|---|
| 8.1 | Ventana de habitación | 0.80 × 0.80 | Aluminio, vidrio de 4 mm | – | 3 | 8 | 8 | 8 | 27 |
| 8.2 | Ventana de baño/cocina | 0.91 × 0.91 | Aluminio, vidrio esmerilado de 4 mm | – | 2 | 1 | 1 | 1 | 5 |
| 8.3 | Ventanal de circulación | 0.90 × 2.41 | Vidrio de seguridad de 6 mm (NSR-10 K) | – | 1 | 1 | 1 | 1 | 4 |
| 8.4 | Ventanal de fachada (acceso) | 0.68 × 2.20 | Vidrio de seguridad de 6 mm | 2 | – | – | – | – | 2 |
| 8.5 | Área de vidrio en ventanas | m² | | 2.99 | 5.75 | 8.12 | 8.12 | 8.12 | 33.09 |
| 8.6 | Vidrio templado en puertas-ventanal | m² | | 2.52 | – | 3.68 | 3.68 | 3.68 | 13.56 |
| 8.7 | Alfajías | ml | | – | 4.22 | 7.31 | 7.31 | 7.31 | 26.15 |

## 9. Aparatos sanitarios
| Ítem | Aparato | P1 | P2 | P3 | P4 | P5 | Total |
|---|---|---|---|---|---|---|---|
| 9.1 | Sanitario de bajo consumo (≤ 4.8 L) | 2 | 3 | 3 | 3 | 3 | 14 |
| 9.2 | Cabina de ducha | – | 3 | 3 | 3 | 3 | 12 |
| 9.3 | Regadera de ducha | – | 3 | 3 | 3 | 3 | 12 |
| 9.4 | Ducha accesible con barras (NTC 6047) | 1 | – | – | – | – | 1 |
| 9.5 | Lavamanos | 2 | 3 | 3 | 3 | 3 | 14 |
| 9.6 | Mampara de vidrio templado de 8 mm | – | – | 3 | 3 | 3 | 9 |
| 9.7 | Grifería de lavamanos | 2 | 3 | 3 | 3 | 3 | 14 |
| 9.8 | Mezclador de ducha | 1 | 3 | 3 | 3 | 3 | 13 |
| 9.9 | Juego de incrustaciones | 2 | 3 | 3 | 3 | 3 | 14 |
| 9.10 | Sifón de piso de 3" | 1 | 3 | 3 | 3 | 3 | 13 |

## 10. Ascensor
| 10.1 | Ascensor de pasajeros (MRL, 6 paradas a confirmar) | gl | 1 |
|---|---|---|---|

---

## Observaciones importantes
1. **Muros:** el aviso del dashboard dice 1,377.66 m² en 130 muros, pero la tabla por piso suma **1,359.15 m² en 131 muros**. Se usó la tabla.
2. **Puertas:** el KPI dice 44, pero el detalle por función suma **42**. Hay que conciliar en Revit.
3. **Aparatos sanitarios:** el KPI dice 61 y el gráfico por nivel suma 56, pero la tabla detallada suma **62**. Se usó la tabla.
4. **Falta el plano EST-01** (EST-02 es la hoja 2 de 2). Sin él no se puede cuantificar la estructura general: cimentación, columnas, vigas, placas y escaleras.
5. **El ascensor no está modelado en Revit.** Hay que coordinar la ubicación del ducto y validar el foso con el proveedor: profundidad útil de 1.20 m y luz de 1.50 × 1.90 m.
6. **Dudas del EST-02 para el calculista:**
   - La sección de las columnas muestra 4 #4, pero el detalle indica 6 #4 L=1.80 + 6 #4 L=1.85. Se cuantificaron las 12 barras; si son solo 4 #4, el acero baja unos 60 kg.
   - El fleje aparece con L=1.50 y con L=1.90. Se usó 1.50; con 1.90 el acero sube unos 18 kg.
   - La lámina no indica f'c, fy ni si el perfil es HEA o HEB.
   - El rótulo dice "TANQUE ALMACENAMIENTO".
7. **Quedan pendientes** por falta de datos: enchapes, guardaescobas, mesones, instalaciones, cubierta, escaleras y P6–P7.
