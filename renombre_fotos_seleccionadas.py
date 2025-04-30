import os
import pandas as pd
import sys
import re
import string

def limpiar_nombre(nombre):
    """Eliminar caracteres no válidos para nombres de archivo."""
    return "".join(c for c in nombre if c in string.printable and c not in r'<>:"/\\|?*')

def obtener_columnas_nombres_apellidos(df):
    """Obtener las columnas que contienen 'NOMBRE' o 'APELLIDO' y 'ORDEN DE FOTO'."""
    columnas_nombres = [col for col in df.columns if re.match(r'NOMBRE\d*', col, re.IGNORECASE)]
    columnas_apellidos = [col for col in df.columns if re.match(r'APELLIDO\d*', col, re.IGNORECASE)]
    
    if "ORDEN DE FOTO" not in df.columns:
        raise ValueError("❌ El archivo Excel debe contener la columna 'ORDEN DE FOTO'.")
    
    return columnas_nombres, columnas_apellidos

try:
    if len(sys.argv) < 2:
        print("❌ No se proporcionó la ruta del archivo Excel.")
        input("\nPresiona Enter para salir...")
        sys.exit()

    ruta_excel = sys.argv[1]
    carpeta_actual = os.path.dirname(os.path.abspath(ruta_excel))

    print(f"📂 Ruta Excel recibida: {ruta_excel}")
    print(f"📂 Carpeta contenedora: {carpeta_actual}")

    df = pd.read_excel(ruta_excel, engine="openpyxl")

    # Obtener columnas necesarias
    columnas_nombres, columnas_apellidos = obtener_columnas_nombres_apellidos(df)

    df["IMAGEN"] = ""  # Crear columna de imágenes vacía

    # Combinar las columnas de nombres y apellidos
    df["NOMBRE"] = df[columnas_nombres].apply(lambda x: ' '.join(x.dropna().astype(str)), axis=1)
    df["APELLIDO"] = df[columnas_apellidos].apply(lambda x: ' '.join(x.dropna().astype(str)), axis=1)

    # Filtrar filas válidas
    df_validos = df[pd.to_numeric(df["ORDEN DE FOTO"], errors="coerce").notna()].copy()
    df_validos["ORDEN DE FOTO"] = df_validos["ORDEN DE FOTO"].astype(int)
    df_validos = df_validos.sort_values(by="ORDEN DE FOTO").reset_index(drop=True)

    # Crear mapa de nombres
    mapa_nombres = {
        int(row["ORDEN DE FOTO"]): f"{row['ORDEN DE FOTO']} {row['NOMBRE']} {row['APELLIDO']}"
        for _, row in df_validos.iterrows()
    }

    print("\n📋 Datos con orden de foto extraídos del Excel:")
    for orden, nombre in mapa_nombres.items():
        print(f"  {orden} → {nombre}")

    extensiones = (".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tiff")
    archivos_en_carpeta = [f for f in os.listdir(carpeta_actual) if f.lower().endswith(extensiones) and re.search(r"\d+", f)]

    if not archivos_en_carpeta:
        print("⚠️ No se encontraron imágenes con números en la carpeta.")
        input("\nPresiona Enter para salir...")
        sys.exit()

    # Asociar imágenes
    archivos_mapeados = []
    for archivo in archivos_en_carpeta:
        match = re.search(r"(\d+)", archivo)
        if match:
            numero = int(match.group(1))
            extension = os.path.splitext(archivo)[1].lower()
            archivos_mapeados.append((numero, archivo, extension))

    archivos_mapeados.sort()

    print("\n📸 Lista de archivos ordenados por número detectado:")
    for num, archivo, ext in archivos_mapeados:
        print(f"  {archivo} → {num}")

    print("\n🔄 Iniciando proceso de renombrado...")
    archivos_renombrados = []

    for i, (num, archivo, ext) in enumerate(archivos_mapeados):
        if i < len(mapa_nombres):
            orden_foto = list(mapa_nombres.keys())[i]
            nombre_limpio = limpiar_nombre(mapa_nombres[orden_foto])
            nuevo_nombre = f"{nombre_limpio}{ext}"
            ruta_original = os.path.join(carpeta_actual, archivo)
            ruta_nueva = os.path.join(carpeta_actual, nuevo_nombre)

            contador = 1
            while os.path.exists(ruta_nueva):
                nuevo_nombre = f"{nombre_limpio}_{contador}{ext}"
                ruta_nueva = os.path.join(carpeta_actual, nuevo_nombre)
                contador += 1

            try:
                os.rename(ruta_original, ruta_nueva)
                print(f"✅ {archivo} → {nuevo_nombre}")
                df.loc[df["ORDEN DE FOTO"] == orden_foto, "IMAGEN"] = ruta_nueva  # ← RUTA COMPLETA
                archivos_renombrados.append((archivo, nuevo_nombre))
            except Exception as e:
                print(f"❌ Error al renombrar {archivo}: {e}")

    # Guardar en el MISMO archivo Excel original
    df.to_excel(ruta_excel, index=False, engine="openpyxl")

    print("\n🚀 Renombrado completado con éxito.")
    print(f"📂 Archivo Excel original actualizado: {ruta_excel}")

    if archivos_renombrados:
        print("\n📄 Archivos renombrados exitosamente:")
        for original, nuevo in archivos_renombrados:
            print(f"  {original} → {nuevo}")
    else:
        print("\n⚠️ No se renombró ningún archivo.")

except Exception as e:
    print(f"\n❌ Se ha producido un error: {e}")
