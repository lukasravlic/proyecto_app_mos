# %%
def main():
    import pandas as pd
    import os
    from pathlib import Path
    import datetime
    import getpass
    import tkinter as tk
    from tkinter import simpledialog, messagebox

    usuario = getpass.getuser()

    hoy = datetime.datetime.today()

    # Function to create CSV files


    ruta_repo = Path(f'C:/Users/{usuario}/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras Maestros/Bases Indicadores en CSV {hoy.year}-{hoy.month:02d}')
    ruta_repo.mkdir(parents=True, exist_ok=True)

    ruta_maestro = Path(f"C:/Users/{usuario}/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras Maestros/{hoy.year}/{hoy.year}-{hoy.month:02d}")
    rep_maestros = ruta_maestro.glob('*')  # List all files in the directory


    # 1. Inicializamos todas las banderas
    encontro_cod = False
    encontro_cod_prem = False
    encontro_mara = False
    encontro_obso = False
    encontro_suggested = False
    encontro_sp = False

    # --- SECCIÓN 1: Maestros ---
    for file_path in rep_maestros:
        # Código Actual
        if 'COD_ACTUAL_R3' in file_path.name:
            df_cod_actual = pd.read_excel(file_path, sheet_name='Sheet1')
            if df_cod_actual is not None:
                print(f'Archivo de Codigo Actual usado: {file_path.name}')
                df_cod_actual.to_csv(f"C:/Users/{usuario}/Inchcape/Planificación y Compras Chile - Documentos/Bases Indicadores en CSV {hoy.year}-{hoy.month:02d}/COD_ACTUAL.csv")
                encontro_cod = True

        #Codigo Actual Premium
        if 'COD_ACTUAL_PREM' in file_path.name:
            df_cod_actual_prem = pd.read_excel(file_path, sheet_name='Código Actual')
            if df_cod_actual_prem is not None:
                print(f'Archivo de Codigo Actual Premium usado: {file_path.name}')
                df_cod_actual_prem.to_csv(f"C:/Users/{usuario}/Inchcape/Planificación y Compras Chile - Documentos/Bases Indicadores en CSV {hoy.year}-{hoy.month:02d}/COD_ACTUAL_PREM.csv", index=False)
                encontro_cod_prem = True

        # MARA
        if 'MARA_R3' in file_path.name:
            df_mara = pd.read_excel(file_path, sheet_name='Sheet1')
            if df_mara is not None:
                print(f'Archivo de MARA usado: {file_path.name}')
                df_mara.drop(columns={'Marca'}).to_csv(f"C:/Users/{usuario}/Inchcape/Planificación y Compras Chile - Documentos/Bases Indicadores en CSV {hoy.year}-{hoy.month:02d}/MARA_R3.csv", index=False)
                encontro_mara = True

        # OBSOLESCENCIA
        if 'new_obso' in file_path.name:
            df_obs = pd.read_excel(file_path, sheet_name='Sheet1')
            if df_obs is not None:
                print(f'Archivo de OBSOLESCENCIA usado: {file_path.name}')
                df_obs.to_csv(f"C:/Users/{usuario}/Inchcape/Planificación y Compras Chile - Documentos/Bases Indicadores en CSV {hoy.year}-{hoy.month:02d}/OBSOLECENCIA.csv", index=False)
                encontro_obso = True

    # --- SECCIÓN 2: Plan de Compras (PDC) ---
    ruta_pdc = f"C:/Users/{usuario}/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras OEM/Planificación/Plan de Compras/{hoy.year}"

    # Verificamos si la ruta de la carpeta del año existe antes de entrar
    if os.path.exists(ruta_pdc):
        for i in os.listdir(ruta_pdc):
            # Filtramos por el mes actual (ej: "03")
            if str(hoy.month).zfill(2) in i:
                print(f"Buscando en carpeta de plan de compras del mes: {i}")
                ruta_mes = os.path.join(ruta_pdc, i)
                
                for j in os.listdir(ruta_mes):
                    # Suggested Purchase v2
                    if 'Suggested Purchase' in j and 'v2' in j:
                        print(f'Archivo encontrado: {j}')
                        df_ddp = pd.read_excel(os.path.join(ruta_mes, j), sheet_name='Base', header=1)
                        df_ddp.to_csv(f"C:/Users/{usuario}/Inchcape/Planificación y Compras Chile - Documentos/Bases Indicadores en CSV {hoy.year}-{hoy.month:02d}/Suggested_Purchase.csv", index=False)
                        encontro_suggested = True
                    
                    # SP v2
                    if 'SP' in j and 'v2' in j:
                        print(f'Archivo encontrado: {j}')
                        df_sp = pd.read_excel(os.path.join(ruta_mes, j), sheet_name='Base',header=1)
                        df_sp.to_csv(f"C:/Users/{usuario}/Inchcape/Planificación y Compras Chile - Documentos/Bases Indicadores en CSV {hoy.year}-{hoy.month:02d}/SP.csv", index=False)
                        encontro_sp = True
    else:
        print(f"⚠️ La ruta del año {hoy.year} no existe o no es accesible.")

    # --- SECCIÓN 3: Reporte final de faltantes ---
    print("-" * 50)
    print("RESUMEN DE CARGA:")

    if not encontro_cod:       print('❌ No se encontró: CODIGO ACTUAL')
    if not encontro_cod_prem:  print('❌ No se encontró: CODIGO ACTUAL PREMIUM')
    if not encontro_mara:      print('❌ No se encontró: MARA')
    if not encontro_obso:      print('❌ No se encontró: OBSOLESCENCIA')
    if not encontro_suggested: print('❌ No se encontró: Suggested Purchase v2')
    if not encontro_sp:        print('❌ No se encontró: SP v2')

    if all([encontro_cod, encontro_mara, encontro_obso, encontro_suggested, encontro_sp]):
        print("✅ ¡Todos los archivos se cargaron correctamente!")
    print("-" * 50)


if __name__ == '__main__':
    main()