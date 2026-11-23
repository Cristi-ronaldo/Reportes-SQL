"""
Genera un reporte en Excel a partir de la base de datos de ejemplo.

Este es el corazón del proyecto: en lugar de que una persona abra la base,
copie datos, los pegue en Excel y arme las tablas a mano, el script hace
todo el trabajo en segundos y deja el archivo listo para enviar.

El reporte trae varias hojas:
  1. Resumen general
  2. Ventas por sucursal
  3. Top productos
  4. Ventas por mes
  5. Detalle de ventas

Uso:
    python generar_reporte.py
"""

import os
import sqlite3
from datetime import datetime

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

CARPETA = os.path.dirname(os.path.abspath(__file__))
RUTA_BD = os.path.join(CARPETA, "datos", "empresa_demo.db")
CARPETA_SALIDAS = os.path.join(CARPETA, "salidas")

# Colores corporativos sencillos para que el reporte se vea presentable
AZUL = "1F4E79"
AZUL_CLARO = "DDEBF7"
GRIS = "F2F2F2"


# ---------------------------------------------------------------------------
# Consultas SQL: aquí se "consolida" la información de varias tablas
# ---------------------------------------------------------------------------
CONSULTAS = {
    "resumen": """
        SELECT
            COUNT(*)                          AS total_ventas,
            SUM(cantidad)                     AS unidades_vendidas,
            ROUND(SUM(cantidad * precio_unitario), 2) AS ingresos_totales,
            ROUND(AVG(cantidad * precio_unitario), 2) AS ticket_promedio
        FROM ventas
    """,
    "por_sucursal": """
        SELECT
            s.nombre                          AS sucursal,
            s.ciudad                          AS ciudad,
            s.region                          AS region,
            COUNT(v.id_venta)                 AS numero_ventas,
            SUM(v.cantidad)                   AS unidades,
            ROUND(SUM(v.cantidad * v.precio_unitario), 2) AS ingresos
        FROM ventas v
        JOIN sucursales s ON s.id_sucursal = v.id_sucursal
        GROUP BY s.id_sucursal
        ORDER BY ingresos DESC
    """,
    "top_productos": """
        SELECT
            p.nombre                          AS producto,
            p.categoria                       AS categoria,
            SUM(v.cantidad)                   AS unidades,
            ROUND(SUM(v.cantidad * v.precio_unitario), 2) AS ingresos
        FROM ventas v
        JOIN productos p ON p.id_producto = v.id_producto
        GROUP BY p.id_producto
        ORDER BY ingresos DESC
        LIMIT 10
    """,
    "por_mes": """
        SELECT
            substr(v.fecha, 1, 7)             AS mes,
            COUNT(*)                          AS numero_ventas,
            ROUND(SUM(v.cantidad * v.precio_unitario), 2) AS ingresos
        FROM ventas v
        GROUP BY mes
        ORDER BY mes
    """,
    "detalle": """
        SELECT
            v.id_venta                        AS folio,
            v.fecha                           AS fecha,
            s.nombre                          AS sucursal,
            c.nombre                          AS cliente,
            c.segmento                        AS segmento,
            p.nombre                          AS producto,
            v.cantidad                        AS cantidad,
            v.precio_unitario                 AS precio_unitario,
            ROUND(v.cantidad * v.precio_unitario, 2) AS total
        FROM ventas v
        JOIN sucursales s ON s.id_sucursal = v.id_sucursal
        JOIN clientes   c ON c.id_cliente  = v.id_cliente
        JOIN productos  p ON p.id_producto = v.id_producto
        ORDER BY v.fecha DESC, v.id_venta DESC
    """,
}


# ---------------------------------------------------------------------------
# Utilidades de formato
# ---------------------------------------------------------------------------
def estilo_encabezado(fila, color_fondo=AZUL):
    """Pinta una fila de encabezado con fondo, letra blanca y bordes."""
    borde = Border(
        left=Side(style="thin", color="BFBFBF"),
        right=Side(style="thin", color="BFBFBF"),
        top=Side(style="thin", color="BFBFBF"),
        bottom=Side(style="thin", color="BFBFBF"),
    )
    for celda in fila:
        celda.font = Font(bold=True, color="FFFFFF", size=11)
        celda.fill = PatternFill("solid", fgColor=color_fondo)
        celda.alignment = Alignment(horizontal="center", vertical="center")
        celda.border = borde


def ajustar_anchos(hoja, anchos):
    """Define el ancho de cada columna."""
    for i, ancho in enumerate(anchos, start=1):
        hoja.column_dimensions[get_column_letter(i)].width = ancho


def escribir_tabla(hoja, fila_inicio, encabezados, filas, formato_moneda=None):
    """Escribe encabezados + datos y devuelve la fila donde terminó."""
    for col, titulo in enumerate(encabezados, start=1):
        hoja.cell(row=fila_inicio, column=col, value=titulo)
    estilo_encabezado(hoja[fila_inicio])

    fila = fila_inicio + 1
    for datos in filas:
        for col, valor in enumerate(datos, start=1):
            celda = hoja.cell(row=fila, column=col, value=valor)
            if formato_moneda and col in formato_moneda:
                celda.number_format = '"$"#,##0.00'
            if fila % 2 == 0:
                celda.fill = PatternFill("solid", fgColor=GRIS)
        fila += 1
    return fila - 1


def titulo_hoja(hoja, texto, subtitulo):
    """Pone un título grande y la fecha de generación."""
    hoja["A1"] = texto
    hoja["A1"].font = Font(bold=True, size=16, color=AZUL)
    hoja["A2"] = subtitulo
    hoja["A2"].font = Font(size=10, italic=True, color="808080")
    hoja.merge_cells("A1:F1")
    hoja.merge_cells("A2:F2")


# ---------------------------------------------------------------------------
# Construcción del reporte
# ---------------------------------------------------------------------------
def construir_reporte(datos):
    libro = Workbook()
    fecha_texto = datetime.now().strftime("%d/%m/%Y %H:%M")

    # --- Hoja 1: Resumen -----------------------------------------------------
    hoja = libro.active
    hoja.title = "Resumen"
    titulo_hoja(hoja, "Reporte general de ventas", f"Generado automáticamente el {fecha_texto}")

    r = datos["resumen"][0]
    tarjetas = [
        ("Ventas registradas", r["total_ventas"], "#,##0"),
        ("Unidades vendidas", r["unidades_vendidas"], "#,##0"),
        ("Ingresos totales", r["ingresos_totales"], '"$"#,##0.00'),
        ("Ticket promedio", r["ticket_promedio"], '"$"#,##0.00'),
    ]
    fila = 4
    for etiqueta, valor, formato in tarjetas:
        hoja.cell(row=fila, column=1, value=etiqueta).font = Font(bold=True, size=11)
        celda_valor = hoja.cell(row=fila, column=2, value=valor)
        celda_valor.font = Font(bold=True, size=14, color=AZUL)
        celda_valor.number_format = formato
        celda_valor.fill = PatternFill("solid", fgColor=AZUL_CLARO)
        celda_valor.alignment = Alignment(horizontal="center")
        hoja.cell(row=fila, column=1).fill = PatternFill("solid", fgColor=GRIS)
        fila += 2

    hoja["A13"] = "Este reporte se genera solo. No requiere copiar ni pegar datos."
    hoja["A13"].font = Font(italic=True, size=10, color="808080")
    ajustar_anchos(hoja, [28, 22, 18, 18, 18, 18])

    # --- Hoja 2: Ventas por sucursal ----------------------------------------
    hoja = libro.create_sheet("Ventas por sucursal")
    titulo_hoja(hoja, "Ventas por sucursal", "Comparativo de ingresos y unidades por tienda")
    encabezados = ["Sucursal", "Ciudad", "Región", "N.º de ventas", "Unidades", "Ingresos"]
    filas = [
        (d["sucursal"], d["ciudad"], d["region"], d["numero_ventas"], d["unidades"], d["ingresos"])
        for d in datos["por_sucursal"]
    ]
    fin = escribir_tabla(hoja, 4, encabezados, filas, formato_moneda={6})
    ajustar_anchos(hoja, [24, 20, 14, 14, 12, 16])

    # Gráfica de barras de ingresos por sucursal
    grafica = BarChart()
    grafica.title = "Ingresos por sucursal"
    grafica.y_axis.title = "Ingresos"
    grafica.x_axis.title = "Sucursal"
    datos_grafica = Reference(hoja, min_col=6, min_row=4, max_row=fin)
    categorias = Reference(hoja, min_col=1, min_row=5, max_row=fin)
    grafica.add_data(datos_grafica, titles_from_data=True)
    grafica.set_categories(categorias)
    grafica.height = 8
    grafica.width = 16
    hoja.add_chart(grafica, "H4")

    # --- Hoja 3: Top productos ----------------------------------------------
    hoja = libro.create_sheet("Top productos")
    titulo_hoja(hoja, "Productos más vendidos", "Los 10 productos con mayores ingresos")
    encabezados = ["Producto", "Categoría", "Unidades", "Ingresos"]
    filas = [(d["producto"], d["categoria"], d["unidades"], d["ingresos"]) for d in datos["top_productos"]]
    fin = escribir_tabla(hoja, 4, encabezados, filas, formato_moneda={4})
    ajustar_anchos(hoja, [28, 18, 12, 16])

    grafica = BarChart()
    grafica.type = "bar"
    grafica.title = "Top 10 productos por ingresos"
    datos_grafica = Reference(hoja, min_col=4, min_row=4, max_row=fin)
    categorias = Reference(hoja, min_col=1, min_row=5, max_row=fin)
    grafica.add_data(datos_grafica, titles_from_data=True)
    grafica.set_categories(categorias)
    grafica.height = 10
    grafica.width = 18
    hoja.add_chart(grafica, "F4")

    # --- Hoja 4: Ventas por mes ---------------------------------------------
    hoja = libro.create_sheet("Ventas por mes")
    titulo_hoja(hoja, "Evolución mensual", "Ingresos y número de ventas mes con mes")
    encabezados = ["Mes", "N.º de ventas", "Ingresos"]
    filas = [(d["mes"], d["numero_ventas"], d["ingresos"]) for d in datos["por_mes"]]
    fin = escribir_tabla(hoja, 4, encabezados, filas, formato_moneda={3})
    ajustar_anchos(hoja, [14, 16, 18])

    grafica = BarChart()
    grafica.title = "Ingresos por mes"
    datos_grafica = Reference(hoja, min_col=3, min_row=4, max_row=fin)
    categorias = Reference(hoja, min_col=1, min_row=5, max_row=fin)
    grafica.add_data(datos_grafica, titles_from_data=True)
    grafica.set_categories(categorias)
    grafica.height = 8
    grafica.width = 16
    hoja.add_chart(grafica, "E4")

    # --- Hoja 5: Detalle -----------------------------------------------------
    hoja = libro.create_sheet("Detalle de ventas")
    titulo_hoja(hoja, "Detalle de ventas", "Cada movimiento con su cliente, producto y total")
    encabezados = ["Folio", "Fecha", "Sucursal", "Cliente", "Segmento", "Producto", "Cantidad", "Precio unitario", "Total"]
    filas = [
        (d["folio"], d["fecha"], d["sucursal"], d["cliente"], d["segmento"], d["producto"],
         d["cantidad"], d["precio_unitario"], d["total"])
        for d in datos["detalle"]
    ]
    escribir_tabla(hoja, 4, encabezados, filas, formato_moneda={8, 9})
    ajustar_anchos(hoja, [8, 12, 20, 28, 14, 24, 10, 16, 14])
    hoja.freeze_panes = "A5"
    hoja.auto_filter.ref = f"A4:I{4 + len(filas)}"

    return libro


def main():
    if not os.path.exists(RUTA_BD):
        print("No existe la base de datos de ejemplo.")
        print("Primero ejecuta: python datos/crear_base_demo.py")
        return

    conexion = sqlite3.connect(RUTA_BD)
    conexion.row_factory = sqlite3.Row
    try:
        datos = {}
        for nombre, consulta in CONSULTAS.items():
            datos[nombre] = [dict(fila) for fila in conexion.execute(consulta).fetchall()]
    finally:
        conexion.close()

    libro = construir_reporte(datos)

    os.makedirs(CARPETA_SALIDAS, exist_ok=True)
    nombre_archivo = f"reporte_ventas_{datetime.now().strftime('%Y-%m-%d_%H%M')}.xlsx"
    ruta_salida = os.path.join(CARPETA_SALIDAS, nombre_archivo)
    libro.save(ruta_salida)

    print("Reporte generado correctamente.")
    print(f"Archivo: {ruta_salida}")
    print(f"Ventas incluidas: {datos['resumen'][0]['total_ventas']}")
    print(f"Ingresos totales: ${datos['resumen'][0]['ingresos_totales']:,.2f}")


if __name__ == "__main__":
    main()
