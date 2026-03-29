# KPIs BANCARIOS CR
Pipeline de limpieza y visualización de KPIs del sistema bancario costarricense usando datos públicos de la SUGEF
Datos Proyecto KPIs Financieros/
Proyecto de portafolio 
---
Fuente de datos: Superintendencia General de Entidades Financieras (SUGEF)
---

## Problema de negocio

Los equipos de análisis en banca necesitan monitorear mensualmente los indicadores
de salud financiera del sistema para detectar tendencias de riesgo, comparar el
desempeño propio contra el mercado y preparar reportes ejecutivos.

El reporte oficial de la SUGEF se publica como un archivo Excel con formato visual
— títulos, celdas fusionadas, filas vacías — que no se puede conectar directamente
a Power BI. Este proyecto automatiza la transformación de ese archivo en datos
listos para análisis.

---

## Qué hace el script `sugef_to_powerbi.py`

El script recibe el reporte `.xls` descargado de la SUGEF y ejecuta tres etapas:

### 1. Lectura

Abre el archivo y navega su estructura no tabular. El reporte de SUGEF pone el
nombre del banco una sola vez y luego lista sus indicadores en filas consecutivas
— el script mantiene eso como estado mientras recorre las filas, descartando
automáticamente entidades que no son bancos comerciales (cooperativas, casas de
cambio, financieras).

De las más de 40 entidades del archivo, el script extrae solo los **13 bancos
comerciales principales**:

BAC San José, BCR, BCT, Banco General, Banco Nacional, Banco Popular,
Cathay, CMB, Davivienda, Improsa, Lafise, Promerica, Scotiabank.

### 2. Limpieza y transformación

- Convierte el período `"MM/YYYY"` a fecha real (`YYYY-MM-DD`) para que Power BI
  lo reconozca como eje de tiempo sin configuración adicional.
- Elimina registros nulos o duplicados.
- Ordena cronológicamente por banco e indicador.
- Transforma el formato largo (una fila por indicador) al formato ancho que
  espera Power BI (una fila por banco y período, una columna por KPI).

### 3. Exportación

Genera tres archivos en la carpeta `output_powerbi/`.

---

## Outputs

### `kpi_pivot.csv` — el principal, para conectar a Power BI

Una fila por banco y período. Una columna por KPI.

| Columna | Descripción | Unidad | Interpretación |
|---|---|---|---|
| `banco` | Nombre corto del banco | — | — |
| `periodo` | Período original (MM/YYYY) | — | Solo referencia |
| `fecha` | Fecha primer día del mes | YYYY-MM-DD | Usá esta para el eje de tiempo |
| `mora_90d` | Cartera con más de 90 días de atraso o en cobro judicial | % | Menor = mejor |
| `cartera_ab` | Cartera clasificada en categorías A y B (bajo riesgo SUGEF) | % | Mayor = mejor |
| `cobertura_provisiones` | Veces que las estimaciones cubren la cartera morosa >90d | veces | >1x = bien cubierto |
| `roe` | Rentabilidad nominal sobre patrimonio promedio | % | Mayor = más rentable |
| `eficiencia_op` | Veces que la utilidad operativa cubre los gastos administrativos | veces | Mayor = más eficiente |
| `activo_productivo_ratio` | % del activo total que genera ingresos financieros | % | Mayor = más eficiente |
| `captaciones_plazo` | % del fondeo proveniente de depósitos a plazo del público | % | Contextual |
| `spread_intermediacion` | Activo productivo de intermediación / pasivo con costo | veces | Mayor = mejor |

**Cómo conectarlo en Power BI:**
Inicio → Obtener datos → Texto/CSV → seleccioná `kpi_pivot.csv`.
La columna `fecha` ya viene lista — Power BI la detecta como fecha automáticamente.

---

### `kpi_largo.csv` — para análisis en Python o SQL

El mismo dato en formato long: una fila por combinación de banco + período +
indicador. Útil para escribir queries de análisis, calcular rankings o hacer
comparaciones puntuales fuera de Power BI.

```
banco, periodo, fecha, indicador, valor
BAC San José, 03/2025, 2025-03-01, mora_90d, 1.34
BAC San José, 03/2025, 2025-03-01, roe, 13.36
...
```

---

### `kpi_bancario.xlsx` — Excel formateado

Mismo contenido que `kpi_pivot.csv` pero en Excel con dos hojas:

- **KPIs** — datos formateados, filas de BAC San José resaltadas en verde,
  encabezados fijos al hacer scroll.
- **Diccionario** — descripción completa de cada indicador: qué mide, unidad
  y cómo interpretarlo.

Útil para compartir con personas que no usan Power BI, o para adjuntar en el
portafolio.



## Cobertura de datos

- **Entidades:** 13 bancos comerciales del sistema financiero costarricense
- **Indicadores:** 8 KPIs de calidad de cartera, rentabilidad y estructura financiera
- **Período disponible en SUGEF:** desde enero 2022 (el reporte actual llega a marzo 2025)

---

## Requisitos

```bash
pip install pandas openpyxl xlrd
```

Python 3.10 o superior.

---

## Estructura del proyecto

```
Datos Proyecto KPIs Financieros/
│
├── sugef_to_powerbi.py          ← script principal
├── reporte-20260328-121936.xls  ← archivo descargado de SUGEF
├── README.md                    ← este archivo
│
└── output_powerbi/
    ├── kpi_pivot.csv            ← conectar a Power BI
    ├── kpi_largo.csv            ← análisis Python/SQL
    └── kpi_bancario.xlsx        ← Excel formateado
```
