    # %%
def main():
        # Librerías para manipulación y análisis de datos
    # %%
    # Librerías para manipulación y análisis de datos

    import pandas as pd
    import datetime
    from datetime import datetime
    import os
    from dateutil.relativedelta import relativedelta    
    from datetime import timedelta


    import glob
    import getpass
    # Obtiene el nombre de usuario del sistema operativo actual
    nombre_usuario = getpass.getuser()


    # Ruta a la carpeta donde están los archivos Parquet
    carpeta = f'C:/Users/{nombre_usuario}/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras KPI-Reportes/Forecast Error (Python)/OEM Mainstream/Base Forecast Error/Consolidado'  # Reemplaza con tu ruta real

    # Lista para guardar los DataFrames
    dataframes = []

    # Recorremos los archivos en la carpeta
    for archivo in os.listdir(carpeta):
        if archivo.endswith('.parquet') and 'Consolidado' not in archivo:
            print(archivo)
            ruta_archivo = os.path.join(carpeta, archivo)
            try:
                df = pd.read_parquet(ruta_archivo)
                df['archivo_origen'] = archivo  # Agregamos una columna para rastrear el origen
                dataframes.append(df)
            except Exception as e:
                print(f"No se pudo leer {archivo}: {e}")



    # Concatenamos todos los DataFrames
    if dataframes:
        consolidado = pd.concat(dataframes, ignore_index=True)
        ruta_salida = os.path.join(carpeta, 'Consolidado.parquet')
        consolidado['Mes'] = pd.to_datetime(consolidado['Mes'])
        consolidado.to_parquet(ruta_salida, index=False)
        print(f"Consolidado guardado en: {ruta_salida}")
    else:
        print("No se encontraron archivos válidos para consolidar.")

    print(consolidado.head())


if __name__ == '__main__':
    main()  


