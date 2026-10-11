"""Genera un archivo de inventario SIMULADO con errores típicos.
Todos los datos son inventados (no son datos reales de ninguna empresa)."""
import random
from datetime import date, timedelta
import pandas as pd

random.seed(42)

productos = [
    ("Rodamiento 6205", "Rodamientos", 1850), ("Rodamiento 6308", "Rodamientos", 3200),
    ("Correa en V A-42", "Correas", 2400), ("Correa dentada 8M", "Correas", 5100),
    ("Filtro de aire industrial", "Filtros", 7800), ("Filtro de aceite hidráulico", "Filtros", 6400),
    ("Aceite hidráulico 20L", "Lubricantes", 38500), ("Grasa multipropósito 1kg", "Lubricantes", 4200),
    ("Contactor 25A", "Electricos", 9600), ("Relé térmico 18A", "Electricos", 8300),
    ("Cable 2.5mm (rollo)", "Electricos", 21000), ("Sensor inductivo M18", "Electricos", 12500),
    ("Tornillo M10 (caja)", "Ferreteria", 2800), ("Tuerca M10 (caja)", "Ferreteria", 1900),
    ("Guantes de seguridad", "EPP", 1500), ("Casco industrial", "EPP", 6200),
    ("Protector auditivo", "EPP", 2300), ("Lentes de seguridad", "EPP", 1100),
    ("Husillo para hilado", "Repuestos textiles", 15400), ("Cinta transportadora (m)", "Repuestos textiles", 9900),
]
proveedores = ["Industrial Norte", "Repuestos del NEA", "Tecno Parts", "Suministros Chaco"]
ubicaciones = ["A1", "A2", "B1", "B2", "C1"]

filas = []
for i in range(1, 121):
    nombre, cat, costo = random.choice(productos)
    ingreso = date(2025, 1, 1) + timedelta(days=random.randint(0, 600))
    minimo = random.choice([5, 10, 15, 20])
    stock = random.randint(0, 60)
    filas.append({
        "Codigo": f"INV-{i:04d}", "Producto": nombre, "Categoria": cat,
        "Proveedor": random.choice(proveedores), "Fecha_Ingreso": ingreso,
        "Stock_Actual": stock, "Stock_Minimo": minimo,
        "Costo_Unitario": costo, "Ubicacion": random.choice(ubicaciones),
    })

df = pd.DataFrame(filas).astype(object)

# --- Errores típicos de planillas reales ---
for idx in random.sample(range(len(df)), 12):          # fechas en texto, formatos mezclados
    f = df.at[idx, "Fecha_Ingreso"]
    df.at[idx, "Fecha_Ingreso"] = random.choice([f.strftime("%d/%m/%Y"), f.strftime("%d-%m-%y"), f.strftime("%Y.%m.%d")])
for idx in random.sample(range(len(df)), 15):          # costos como texto con $ y comas
    df.at[idx, "Costo_Unitario"] = f"${df.at[idx, 'Costo_Unitario']:,}"
for idx in random.sample(range(len(df)), 10):          # espacios y mayúsculas inconsistentes
    df.at[idx, "Categoria"] = random.choice([" " + df.at[idx, "Categoria"].upper(), df.at[idx, "Categoria"].lower() + " "])
for idx in random.sample(range(len(df)), 8):           # proveedor vacío
    df.at[idx, "Proveedor"] = random.choice([None, "", "N/A"])
for idx in random.sample(range(len(df)), 6):           # stock vacío o texto
    df.at[idx, "Stock_Actual"] = random.choice([None, "s/d", "12 unid."])
for idx in random.sample(range(len(df)), 4):           # stock negativo (error de carga)
    df.at[idx, "Stock_Actual"] = -random.randint(1, 10)

dup = df.sample(7, random_state=1)                      # filas duplicadas
df = pd.concat([df, dup]).sample(frac=1, random_state=3).reset_index(drop=True)

df.to_excel("inventario_sucio.xlsx", index=False)
print(f"Archivo creado con {len(df)} filas")
