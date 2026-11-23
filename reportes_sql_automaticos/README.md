# 📊 Reportes SQL automáticos

Simulador de un sistema que **consulta una base de datos con SQL y genera un reporte de Excel completo** (cinco hojas con tablas, gráficas y formato) sin copiar ni pegar nada a mano.

**Stack:** Python 3 · SQLite · SQL · [openpyxl](https://openpyxl.readthedocs.io/)

> ⚠️ **Datos inventados.** La base de datos, las sucursales, los clientes y los productos son ficticios y sirven solo para demostrar el funcionamiento. No contienen información real de ninguna empresa.

---

## 📑 Contenido

1. [¿Qué hace?](#-qué-hace)
2. [Inicio rápido](#-inicio-rápido)
3. [El simulador](#️-el-simulador)
4. [El reporte de Excel](#-el-reporte-de-excel)
5. [Base de datos de ejemplo](#️-base-de-datos-de-ejemplo)
6. [Estructura del proyecto](#-estructura-del-proyecto)
7. [Adaptarlo a una base de datos real](#-adaptarlo-a-una-base-de-datos-real)
8. [Programar la ejecución automática](#-programar-la-ejecución-automática)
9. [Limitaciones conocidas](#️-limitaciones-conocidas)
10. [Autor y licencia](#-autor-y-licencia)

---

## ✨ ¿Qué hace?

En muchos equipos, cada semana o cada mes alguien abre la base de datos, copia información de varias tablas, la pega en Excel, la ordena, le da formato y la envía. Es lento y propenso a errores.

Este proyecto automatiza ese flujo:

1. Lee cuatro tablas: **sucursales, clientes, productos y ventas**.
2. Las cruza y las resume con **consultas SQL** (`JOIN`, `GROUP BY`, agregaciones).
3. Construye un archivo `.xlsx` con cinco hojas, tres de ellas con gráficas.
4. Permite generar el reporte completo o el de un mes específico, y registra las ejecuciones del simulador en una bitácora.

---

## 🚀 Inicio rápido

**Requisitos:** Python 3 (probado en 3.12) y la librería `openpyxl`. SQLite ya viene incluido con Python.

```bash
# 1. Clonar el repositorio
git clone https://github.com/Cristi-ronaldo/reportes-sql-automaticos.git
cd reportes-sql-automaticos

# 2. Instalar la dependencia
pip install -r requirements.txt

# 3. Crear la base de datos de ejemplo (una sola vez)
python datos/crear_base_demo.py

# 4a. Abrir el simulador con menú...
python simulador.py

# 4b. ...o generar el reporte directamente
python generar_reporte.py
```

El archivo de Excel queda en la carpeta `salidas/` con el nombre `reporte_ventas_AAAA-MM-DD_HHMM.xlsx`.

### Salida de ejemplo

Con los datos de ejemplo, la opción 3 del simulador muestra algo como esto:

```text
RESUMEN DE LA INFORMACIÓN CONSOLIDADA
--------------------------------------------------------------
  Ventas registradas : 420
  Unidades vendidas  : 1,870
  Ingresos totales   : $6,450,545.00
  Ticket promedio    : $15,358.44
--------------------------------------------------------------
  Top 3 sucursales por ingresos:
    - Sucursal Norte         $1,584,125.00
    - Sucursal Occidente     $1,475,625.00
    - Sucursal Centro        $1,374,890.00
  Top 3 productos por ingresos:
    - Laptop Pro 14          $2,497,500.00
    - Escritorio ajustable   $1,036,800.00
    - Monitor 27 pulgadas    $  663,600.00
```

---

## 🕹️ El simulador

`simulador.py` es un menú de consola para probar todo sin escribir comandos:

| Opción | Acción |
| --- | --- |
| 1 | Genera el reporte completo en Excel. |
| 2 | Genera el reporte de un mes específico (lista los meses disponibles). |
| 3 | Muestra el resumen en pantalla, sin crear archivo. |
| 4 | Simula el modo automático: genera 3 reportes seguidos, como una tarea programada. |
| 5 | Muestra la bitácora de ejecuciones (`salidas/bitacora_ejecuciones.txt`). |
| 0 | Salir. |

---

## 📗 El reporte de Excel

| Hoja | Contenido | Gráfica |
| --- | --- | --- |
| **Resumen** | Ventas registradas, unidades vendidas, ingresos totales y ticket promedio. | — |
| **Ventas por sucursal** | Comparativo de ventas, unidades e ingresos por tienda. | Barras de ingresos por sucursal |
| **Top productos** | Los 10 productos con más ingresos. | Barras horizontales |
| **Ventas por mes** | Evolución mensual de ventas e ingresos. | Barras de ingresos por mes |
| **Detalle de ventas** | Cada venta con cliente, producto y total. Con filtros y encabezado fijo. | — |

Todas las hojas llevan título, fecha de generación, encabezados con color y formato de moneda.

---

## 🗄️ Base de datos de ejemplo

`datos/crear_base_demo.py` crea `datos/empresa_demo.db` (SQLite) con datos inventados. Usa una semilla fija (`random.seed(7)`) para que los resultados sean reproducibles.

| Tabla | Contenido | Registros |
| --- | --- | --- |
| `sucursales` | Tiendas con ciudad y región | 5 |
| `clientes` | Empresas con su segmento (Consumidor, Empresa, Gobierno) | 15 |
| `productos` | Nombre, categoría y precio | 12 |
| `ventas` | Fecha, sucursal, cliente, producto, cantidad y precio unitario | 420 (últimos 180 días) |

Las consultas están en el diccionario `CONSULTAS` de `generar_reporte.py`: `resumen`, `por_sucursal`, `top_productos`, `por_mes` y `detalle`.

---

## 📁 Estructura del proyecto

```text
reportes-sql-automaticos/
├── generar_reporte.py        # Consultas SQL y construcción del Excel (núcleo del proyecto)
├── simulador.py              # Menú de consola para probar el sistema
├── requirements.txt          # Dependencias de Python
├── datos/
│   └── crear_base_demo.py    # Crea la base de datos de ejemplo (empresa_demo.db)
└── salidas/                  # Reportes .xlsx y bitácora (se crea al ejecutar)
```

---

## 🔌 Adaptarlo a una base de datos real

La lógica de reportes no cambia; solo se reemplaza la conexión. Hay que modificar `sqlite3.connect(...)` en `generar_reporte.py` (función `main`) y en `simulador.py` (función `consultar`).

| Base de datos | Librería de Python habitual |
| --- | --- |
| SQL Server | `pyodbc` |
| PostgreSQL | `psycopg2` |
| MySQL | `mysql-connector-python` |
| Oracle | `oracledb` |

Ten en cuenta las diferencias de dialecto SQL:

- `substr(v.fecha, 1, 7)` (agrupa por mes) es de SQLite; en otros motores se usa `FORMAT`, `TO_CHAR`, `DATE_FORMAT`, etc.
- `LIMIT 10` no existe en SQL Server (allí se usa `TOP 10`).
- Los parámetros usan `?` en `sqlite3`; `psycopg2` y `mysql-connector` usan `%s`.

---

## ⏰ Programar la ejecución automática

`generar_reporte.py` se puede ejecutar solo con el programador de tareas del sistema:

```bash
# Linux/macOS (cron): todos los lunes a las 7:00 a.m.
0 7 * * 1 cd /ruta/a/reportes-sql-automaticos && python3 generar_reporte.py
```

En Windows se puede usar el **Programador de tareas** con el mismo comando. Enviar el archivo por correo es una mejora posible, pero **no está implementada** en este proyecto.

---

## ⚠️ Limitaciones conocidas

- La **bitácora** solo se llena cuando se genera el reporte desde `simulador.py`, no al ejecutar `generar_reporte.py` directamente.
- El **filtro por mes** solo existe en el simulador (opción 2); `generar_reporte.py` siempre genera el reporte completo.
- El formato de moneda es fijo (`$` con dos decimales) y las ciudades de ejemplo son de México.
- Es un **proyecto de demostración**: no incluye pruebas automáticas ni manejo avanzado de errores.

---

## 👤 Autor y licencia

**Autor:** Cristian Lozano · GitHub: [@Cristi-ronaldo](https://github.com/Cristi-ronaldo)

**Licencia:** [MIT](LICENSE).
