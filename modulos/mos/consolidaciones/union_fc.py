def main():    # %%
    print('⌚Proceso de consolidación iniciado, los archivos a utilizar son: ')
    import pandas as pd

    # %%
    import datetime
    import getpass
    usuario = getpass.getuser()

    # %%
    hoy = datetime.datetime.today()

    # %%
    str(hoy.month).zfill(2)

    # %%
    fc_folder = f"C:/Users/{usuario}/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras KPI-Reportes/Gerenciamiento MOS/Panel PBI/bases mensuales/forecast/{str(hoy.year)}-{str(hoy.month).zfill(2)}"

    # %%
    import pandas as pd
    import os

    # Ruta de la carpeta donde están los archivos
    carpeta = fc_folder

    # Lista para guardar los DataFrames
    dataframes = []

    # Iterar sobre todos los archivos .csv
    for archivo in os.listdir(carpeta):
        if archivo.endswith('.csv'):
            print(f'\t{archivo}')
            df = pd.read_csv(os.path.join(carpeta, archivo))
            dataframes.append(df)

    # Concatenar todos los DataFrames
    df_consolidado = pd.concat(dataframes, ignore_index=True)

    # Guardar en un nuevo archivo
    # df_consolidado.to_csv('archivo_consolidado.csv', index=False)


    # %%
    meses = {
        '01': 'ene',
        '02': 'feb',
        '03': 'mar',
        '04': 'abr',
        '05': 'may',
        '06': 'jun',
        '07': 'jul',
        '08': 'ago',
        '09': 'sep',
        '10': 'oct',
        '11': 'nov',
        '12': 'dic'
    }

    # %%
    df_consolidado.drop(columns={'Unnamed: 0'}, inplace=True)

    # %%
    # df_consolidado = pd.concat([df_consolidado,df_consolidado_30],ignore_index=True)

    # %%
    df_consolidado['fc_mean'] = df_consolidado['fc_mean'].astype(str).str.replace(',', '.').astype(float)
    df_consolidado.to_csv(f"C:/Users/{usuario}/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras KPI-Reportes/Gerenciamiento MOS/Panel PBI/automatizacion/consolidado_fc_{meses.get(str(hoy.month).zfill(2))}.csv")
    print(f"\n🎊Proceso finalizado de manera correcta! Archivo guardado en: \nC:/Users/{usuario}/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras KPI-Reportes/Gerenciamiento MOS/Panel PBI/automatizacion/consolidado_fc_{meses.get(str(hoy.month).zfill(2))}.csv")
    # %%

if __name__ == '__main__':
    main()