"""
SIMULADOR del proyecto "Reportes SQL automáticos".

Es un menú sencillo para que cualquier persona (aunque no sepa programar)
pueda probar cómo funcionaría el sistema en una empresa.

Uso:
    python simulador.py
"""

import os
import sqlite3
import time
from datetime import datetime

import generar_reporte as gr

CARPETA = os.path.dirname(os.path.abspath(__file__))
BITACORA = os.path.join(CARPETA, "salidas", "bitacora_ejecuciones.txt")


def separador():
    print("-" * 62)


def registrar(mensaje):
    """Anota cada ejecución en una bitácora, como haría un sistema real."""
    os.makedirs(os.path.dirname(BITACORA), exist_ok=True)
    with open(BITACORA, "a", encoding="utf-8") as archivo:
        archivo.write(f"[{datetime.now():%Y-%m-%d %H:%M:%S}] {mensaje}\n")


def consultar(nombre_consulta, mes=None):
    """Ejecuta una consulta y devuelve los resultados como lista de diccionarios."""
    consulta = gr.CONSULTAS[nombre_consulta]
    parametros = ()

    # Si se pide un mes, cambiamos la tabla "ventas" por una versión filtrada.
    # Así funciona igual con las consultas que usan JOIN.
    if mes:
        filtro = "SELECT * FROM ventas WHERE substr(fecha,1,7) = ?"
        if "FROM ventas v" in consulta:
            consulta = consulta.replace("FROM ventas v", f"FROM ({filtro}) v")
        else:
            consulta = consulta.replace("FROM ventas", f"FROM ({filtro})")
        parametros = (mes,)

    conexion = sqlite3.connect(gr.RUTA_BD)
    conexion.row_factory = sqlite3.Row
    try:
        return [dict(fila) for fila in conexion.execute(consulta, parametros).fetchall()]
    finally:
        conexion.close()


def cargar_todo(mes=None):
    return {nombre: consultar(nombre, mes) for nombre in gr.CONSULTAS}


def generar(mes=None):
    """Genera el Excel y devuelve la ruta del archivo."""
    datos = cargar_todo(mes)
    libro = gr.construir_reporte(datos)

    os.makedirs(gr.CARPETA_SALIDAS, exist_ok=True)
    etiqueta = f"_{mes}" if mes else ""
    nombre = f"reporte_ventas{etiqueta}_{datetime.now():%Y-%m-%d_%H%M%S}.xlsx"
    ruta = os.path.join(gr.CARPETA_SALIDAS, nombre)
    libro.save(ruta)

    registrar(f"Reporte generado: {nombre} ({datos['resumen'][0]['total_ventas']} ventas)")
    return ruta, datos


def mostrar_resumen(datos):
    r = datos["resumen"][0]
    separador()
    print("RESUMEN DE LA INFORMACIÓN CONSOLIDADA")
    separador()
    print(f"  Ventas registradas : {r['total_ventas']:,}")
    print(f"  Unidades vendidas  : {r['unidades_vendidas']:,}")
    print(f"  Ingresos totales   : ${r['ingresos_totales']:,.2f}")
    print(f"  Ticket promedio    : ${r['ticket_promedio']:,.2f}")
    separador()
    print("  Top 3 sucursales por ingresos:")
    for s in datos["por_sucursal"][:3]:
        print(f"    - {s['sucursal']:<22} ${s['ingresos']:>12,.2f}")
    print("  Top 3 productos por ingresos:")
    for p in datos["top_productos"][:3]:
        print(f"    - {p['producto']:<22} ${p['ingresos']:>12,.2f}")
    separador()


def meses_disponibles():
    conexion = sqlite3.connect(gr.RUTA_BD)
    try:
        filas = conexion.execute(
            "SELECT DISTINCT substr(fecha,1,7) AS mes FROM ventas ORDER BY mes"
        ).fetchall()
        return [f[0] for f in filas]
    finally:
        conexion.close()


def opcion_reporte_completo():
    print("\nGenerando reporte completo, espera un momento...")
    ruta, datos = generar()
    mostrar_resumen(datos)
    print(f"Archivo Excel creado:\n  {ruta}")


def opcion_reporte_mes():
    meses = meses_disponibles()
    print("\nMeses con información disponible:")
    for i, mes in enumerate(meses, start=1):
        print(f"  {i}. {mes}")
    eleccion = input("Elige el número del mes: ").strip()
    try:
        mes = meses[int(eleccion) - 1]
    except (ValueError, IndexError):
        print("Opción no válida. Regresando al menú.")
        return

    print(f"\nGenerando reporte de {mes}...")
    ruta, datos = generar(mes)
    mostrar_resumen(datos)
    print(f"Archivo Excel creado:\n  {ruta}")


def opcion_modo_automatico():
    print("\nMODO AUTOMÁTICO (simulación de tarea programada)")
    print("Se generarán 3 reportes seguidos, como si el sistema corriera solo.")
    print("Cada reporte queda guardado y registrado en la bitácora.\n")
    for numero in range(1, 4):
        print(f"Ejecución {numero} de 3...")
        ruta, datos = generar()
        print(f"  OK -> {os.path.basename(ruta)}")
        if numero < 3:
            time.sleep(2)
    print(f"\nBitácora de ejecuciones:\n  {BITACORA}")


def opcion_ver_bitacora():
    if not os.path.exists(BITACORA):
        print("\nTodavía no hay ejecuciones registradas.")
        return
    print("\nBITÁCORA DE EJECUCIONES")
    separador()
    with open(BITACORA, encoding="utf-8") as archivo:
        print(archivo.read().strip())
    separador()


def main():
    if not os.path.exists(gr.RUTA_BD):
        print("Primero hay que crear la base de datos de ejemplo.")
        print("Ejecuta: python datos/crear_base_demo.py")
        return

    while True:
        print()
        print("=" * 62)
        print("   SIMULADOR: REPORTES SQL AUTOMÁTICOS")
        print("=" * 62)
        print("  1. Generar el reporte completo en Excel")
        print("  2. Generar el reporte de un mes específico")
        print("  3. Ver el resumen en pantalla")
        print("  4. Simular el modo automático (3 ejecuciones)")
        print("  5. Ver la bitácora de ejecuciones")
        print("  0. Salir")
        print("=" * 62)
        opcion = input("Elige una opción: ").strip()

        if opcion == "1":
            opcion_reporte_completo()
        elif opcion == "2":
            opcion_reporte_mes()
        elif opcion == "3":
            mostrar_resumen(cargar_todo())
        elif opcion == "4":
            opcion_modo_automatico()
        elif opcion == "5":
            opcion_ver_bitacora()
        elif opcion == "0":
            print("\nHasta luego. El reporte queda listo para enviarse.")
            break
        else:
            print("\nOpción no válida, intenta de nuevo.")


if __name__ == "__main__":
    main()
