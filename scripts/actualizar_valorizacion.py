# ==============================================================================
# Script de Automatización DataOps - Valorización Diaria de Bonos Soberanos
# Autor: Yorgo Rodrigo Gonzales Coricasa
# Proyecto: Valorización de Bonos Soberanos (SBS & Portafolio MTM)
# ==============================================================================

import os
import pandas as pd
import openpyxl

def actualizar_valorizacion():
    dir_data = '../data'
    if not os.path.exists(dir_data):
        dir_data = 'data'
    
    archivo_portafolio = os.path.join(dir_data, 'Portafolio de Bonos Soberanos.xlsx')
    if not os.path.exists(archivo_portafolio):
        archivo_portafolio = 'Portafolio de Bonos Soberanos.xlsx'

    # Buscar explícitamente el archivo del 02.11.26
    archivos_sbs = [f for f in os.listdir(dir_data) if f.startswith('Renta Fija 02.11.26') and f.endswith('.xls')]
    if not archivos_sbs:
        archivos_sbs = [f for f in os.listdir('.') if f.startswith('Renta Fija 02.11.26') and f.endswith('.xls')]
    
    if not archivos_sbs:
        # Si no encuentra el del 02 específicamente, toma el más reciente disponible
        archivos_sbs = [f for f in os.listdir(dir_data) if f.startswith('Renta Fija') and f.endswith('.xls')]
        if not archivos_sbs:
            archivos_sbs = [f for f in os.listdir('.') if f.startswith('Renta Fija') and f.endswith('.xls')]

    if not archivos_sbs:
        raise FileNotFoundError("No se encontró ningún archivo de vector SBS (Renta Fija *.xls).")
    
    # Seleccionar el archivo de forma prioritaria para el 02.11.26
    archivo_sbs_elegido = 'Renta Fija 02.11.26.xls' if 'Renta Fija 02.11.26.xls' in archivos_sbs else archivos_sbs[-1]
    
    archivo_sbs_nuevo = os.path.join(dir_data, archivo_sbs_elegido) if os.path.exists(os.path.join(dir_data, archivo_sbs_elegido)) else archivo_sbs_elegido
    archivo_salida = os.path.join(dir_data, 'Portafolio de Bonos Soberanos_Actualizado.xlsx')

    print(f"[*] Leyendo portafolio desde: {archivo_portafolio}")
    print(f"[*] Leyendo vector SBS oficial: {archivo_sbs_nuevo}")

    # 2. Leer el vector de precios de la SBS (tabla HTML embebida en .xls)
    df_sbs = pd.read_html(archivo_sbs_nuevo)[0]
    df_sbs['Nemónico'] = df_sbs['Nemónico'].astype(str).str.strip()

    # 3. Leer la hoja de valorización del portafolio (header=1 para capturar columnas correctas)
    df_val_original = pd.read_excel(archivo_portafolio, sheet_name='Valorización', header=1)
    df_val_original = df_val_original.dropna(subset=['Nemonico']).copy()

    # 4. Mapear el nuevo Precio Limpio SBS
    precios_sbs_dict = dict(zip(df_sbs['Nemónico'], df_sbs['P. Limpio (%)']))
    df_val_original['Valor SBS'] = df_val_original['Nemonico'].map(precios_sbs_dict)

    # 5. Asegurar la escala correcta del Monto de Compra
    if df_val_original['Monto de Compra'].iloc[0] > 1e8:
        df_val_original['Monto de Compra'] = df_val_original['Monto de Compra'] / 1000.0

    # 6. Cálculo exacto y proporcional del Monto SBS y el Resultado Mark-to-Market (MTM)
    df_val_original['Monto SBS'] = (df_val_original['Valor SBS'] / df_val_original['Valor Compra']) * df_val_original['Monto de Compra']
    df_val_original['Resultado'] = df_val_original['Monto SBS'] - df_val_original['Monto de Compra']

    print("\n--- VALORIZACIÓN AUTOMATIZADA EXITOSA (02.11.2026) ---")
    print(df_val_original[['Nemonico', 'Valor Compra', 'Valor SBS', 'Monto de Compra', 'Monto SBS', 'Resultado']])

    # 7. Generar el nuevo archivo Excel consolidado manteniendo el estándar del portafolio
    with pd.ExcelWriter(archivo_salida, engine='openpyxl') as writer:
        df_port = pd.read_excel(archivo_portafolio, sheet_name='Portafolio')
        df_port.to_excel(writer, sheet_name='Portafolio', index=False)
        
        df_val_original.to_excel(writer, sheet_name='Valorización', index=False)
        
        try:
            df_rf = pd.read_excel(archivo_portafolio, sheet_name='Renta Fija 30.10.26')
            df_rf.to_excel(writer, sheet_name='Renta Fija 30.10.26', index=False)
        except:
            pass
            
        df_sbs.to_excel(writer, sheet_name='Renta Fija Actualizada', index=False)

    print(f"\n[¡ÉXITO!] Archivo de valorización generado en: {archivo_salida}")

if __name__ == '__main__':
    actualizar_valorizacion()
