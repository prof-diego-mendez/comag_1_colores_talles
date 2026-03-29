import pandas as pd
import os

# 1. Solicitar nombres sin extensión
base_lista = input("Nombre del archivo de lista de precios (sin extensión): ").strip()
base_pedido = input("Nombre del archivo del pedido (sin extensión): ").strip()

# 2. Agregar extensión automáticamente
lista_file = f"{base_lista}.xlsx"
pedido_file = f"{base_pedido}.xlsx"

# Validar que existan
if not os.path.isfile(lista_file):
    print(f"❌ No se encontró el archivo '{lista_file}'")
    exit()
if not os.path.isfile(pedido_file):
    print(f"❌ No se encontró el archivo '{pedido_file}'")
    exit()

# 3. Construir nombre de salida fusionado
salida = f"{base_lista}_{base_pedido}_enriquecido.xlsx"

# 4. Cargar archivos
precios_df = pd.read_excel(lista_file)
pedido_df = pd.read_excel(pedido_file)

# 5. Limpiar y normalizar columnas
def normalizar(x):
    return str(x).strip().split('.')[0]

precios_df["asignado"] = precios_df["asignado"].apply(normalizar)
pedido_df["asignado"] = pedido_df["asignado"].apply(normalizar)

# 6. Tabla de descuentos
MARCAS = {
    "BIFERDIL": 5,
    "OSLO": 15,
    "BIOLOOK": 20,
    "DEPILISSIMA": 30,
    "NEWCOLOR": 15,
    "NAIL PROTECT": 15,
    "CAPRI": 20,#,
    "IYOSEI": 10,
    "DODDY": 15,
    "TAN NATURAL": 20
}

# 7. Detectar columnas opcionales
tiene_pvp = "pvp" in precios_df.columns
desc_col = None
for col in precios_df.columns:
    if precios_df[col].dtype == object:
        if precios_df[col].str.contains("|".join(MARCAS), case=False, na=False).any():
            desc_col = col
            break

# 8. Función de búsqueda
def encontrar_filas(valor_pedido):
    match = precios_df[precios_df["asignado"] == valor_pedido]
    if match.empty:
        match = precios_df[precios_df["asignado"].str.contains(valor_pedido, na=False)]
    if not match.empty:
        row = match.iloc[0]
        precio = row["precio_sin_iva"]
        pvp = row["pvp"] if tiene_pvp else None
        desc = None
        if desc_col:
            texto = str(row[desc_col]).upper()
            for marca, dto in MARCAS.items():
                if marca in texto:
                    desc = dto
                    break
        return row["asignado"], precio, pvp, desc
    return None, None, None, None

# 9. Aplicar y crear columnas
pedido_df[["asignado_lista", "precio", "pvp", "descuento"]] = pedido_df["asignado"].apply(
    lambda x: pd.Series(encontrar_filas(x))
)

# 10. Si no hay descuentos, eliminar columna
if pedido_df["descuento"].isna().all():
    pedido_df.drop(columns=["descuento"], inplace=True)

# 11. Guardar con nombre fusionado
pedido_df.to_excel(salida, index=False)
print(f"✅ Archivo guardado como '{salida}' con precio, pvp y (si aplica) descuento.")
