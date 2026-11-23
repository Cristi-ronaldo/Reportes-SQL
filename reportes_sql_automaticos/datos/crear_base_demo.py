"""
Crea una base de datos de ejemplo (simulada) para probar el reporte automático.

La idea es imitar lo que en una empresa real serían varias tablas:
  - sucursales  -> las tiendas u oficinas
  - clientes    -> quién compra
  - productos   -> qué se vende
  - ventas      -> el detalle de cada venta (la tabla "movimiento")

En un caso real, este script se reemplaza por la conexión a la base de datos
verdadera (SQL Server, Oracle, MySQL, PostgreSQL, etc.). Aquí usamos SQLite
porque es un archivo local y no requiere instalar nada.
"""

import os
import random
import sqlite3
from datetime import date, timedelta

# Ruta de la base de datos (queda dentro de la carpeta "datos")
CARPETA = os.path.dirname(os.path.abspath(__file__))
RUTA_BD = os.path.join(CARPETA, "empresa_demo.db")

# Fijamos la semilla para que los datos de ejemplo siempre sean los mismos
random.seed(7)


def crear_tablas(conexion):
    """Borra y vuelve a crear las tablas del ejemplo."""
    cursor = conexion.cursor()

    cursor.executescript(
        """
        DROP TABLE IF EXISTS ventas;
        DROP TABLE IF EXISTS productos;
        DROP TABLE IF EXISTS clientes;
        DROP TABLE IF EXISTS sucursales;

        CREATE TABLE sucursales (
            id_sucursal   INTEGER PRIMARY KEY,
            nombre        TEXT NOT NULL,
            ciudad        TEXT NOT NULL,
            region        TEXT NOT NULL
        );

        CREATE TABLE clientes (
            id_cliente    INTEGER PRIMARY KEY,
            nombre        TEXT NOT NULL,
            segmento      TEXT NOT NULL
        );

        CREATE TABLE productos (
            id_producto   INTEGER PRIMARY KEY,
            nombre        TEXT NOT NULL,
            categoria     TEXT NOT NULL,
            precio        REAL NOT NULL
        );

        CREATE TABLE ventas (
            id_venta      INTEGER PRIMARY KEY AUTOINCREMENT,
            fecha         TEXT NOT NULL,
            id_sucursal   INTEGER NOT NULL,
            id_cliente    INTEGER NOT NULL,
            id_producto   INTEGER NOT NULL,
            cantidad      INTEGER NOT NULL,
            precio_unitario REAL NOT NULL
        );
        """
    )
    conexion.commit()


def cargar_datos(conexion):
    """Inserta datos de ejemplo realistas pero inventados."""
    cursor = conexion.cursor()

    sucursales = [
        (1, "Sucursal Centro", "Ciudad de México", "Centro"),
        (2, "Sucursal Norte", "Monterrey", "Norte"),
        (3, "Sucursal Occidente", "Guadalajara", "Occidente"),
        (4, "Sucursal Sur", "Puebla", "Sur"),
        (5, "Sucursal Costa", "Veracruz", "Golfo"),
    ]
    cursor.executemany("INSERT INTO sucursales VALUES (?, ?, ?, ?)", sucursales)

    segmentos = ["Consumidor", "Empresa", "Gobierno"]
    nombres_clientes = [
        "Comercializadora del Valle", "Grupo Horizonte", "Distribuidora Azteca",
        "Servicios Integrales MX", "Tiendas La Estrella", "Constructora del Norte",
        "Papelería Central", "Tecnología Aplicada", "Alimentos del Campo",
        "Transportes Rápidos", "Ferretería El Tornillo", "Textiles Modernos",
        "Farmacias Vida", "Cafetería Origen", "Muebles y Diseño",
    ]
    clientes = [
        (i + 1, nombre, random.choice(segmentos))
        for i, nombre in enumerate(nombres_clientes)
    ]
    cursor.executemany("INSERT INTO clientes VALUES (?, ?, ?)", clientes)

    productos = [
        (1, "Laptop Pro 14", "Equipos", 18500.00),
        (2, "Monitor 27 pulgadas", "Equipos", 4200.00),
        (3, "Teclado mecánico", "Accesorios", 850.00),
        (4, "Mouse inalámbrico", "Accesorios", 420.00),
        (5, "Impresora láser", "Equipos", 3600.00),
        (6, "Tóner negro", "Consumibles", 980.00),
        (7, "Resma de papel", "Consumibles", 95.00),
        (8, "Silla ergonómica", "Mobiliario", 2900.00),
        (9, "Escritorio ajustable", "Mobiliario", 5400.00),
        (10, "Disco externo 2TB", "Accesorios", 1450.00),
        (11, "Licencia de software", "Software", 2200.00),
        (12, "Servicio de instalación", "Servicios", 1500.00),
    ]
    cursor.executemany("INSERT INTO productos VALUES (?, ?, ?, ?)", productos)

    # Generamos ventas de los últimos 6 meses
    hoy = date.today()
    inicio = hoy - timedelta(days=180)
    ventas = []

    for _ in range(420):
        dias_aleatorios = random.randint(0, 180)
        fecha = inicio + timedelta(days=dias_aleatorios)
        id_sucursal = random.randint(1, 5)
        id_cliente = random.randint(1, len(clientes))
        id_producto = random.randint(1, len(productos))
        precio = next(p[3] for p in productos if p[0] == id_producto)
        cantidad = random.randint(1, 8)
        ventas.append(
            (fecha.isoformat(), id_sucursal, id_cliente, id_producto, cantidad, precio)
        )

    cursor.executemany(
        """
        INSERT INTO ventas (fecha, id_sucursal, id_cliente, id_producto, cantidad, precio_unitario)
        VALUES (?, ?, ?, ?, ?, ?)
        """,
        ventas,
    )
    conexion.commit()
    return len(ventas)


def main():
    print("Creando base de datos de ejemplo...")
    conexion = sqlite3.connect(RUTA_BD)
    try:
        crear_tablas(conexion)
        total = cargar_datos(conexion)
    finally:
        conexion.close()

    print(f"Listo. Base creada en: {RUTA_BD}")
    print(f"Se cargaron {total} ventas de ejemplo.")


if __name__ == "__main__":
    main()
