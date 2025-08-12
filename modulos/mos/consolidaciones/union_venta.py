# %%


def main():
    import pandas as pd
    import pandas as pd
    import os
    import datetime
    import getpass
    usuario = getpass.getuser()
    print('⌚Proceso de consolidación iniciado, los archivos a utilizar son: ')
    # %%
    hoy = datetime.datetime.today
    año = str(hoy().year)
    mes = str(hoy().month).zfill(2)

    # %%
    dict_mes = {
        '01':'ene',
        '02':'feb',
        '03':'mar',
        '04':'abr',
        '05':'may',
        '06':'jun',
        '07':'jul',
        '08':'ago',
        '09':'sep',
        '10':'oct',
        '11':'nov',
        '12':'dic',
    }

    mes_1 = dict_mes.get(str(hoy().month).zfill(2))
    mes_2 = dict_mes.get(str((hoy() + datetime.timedelta(days=30)).month).zfill(2))
    mes_3 = dict_mes.get(str((hoy() + datetime.timedelta(days=60)).month).zfill(2))

    # %%
    mes

    # %%
    venta = f"C:/Users/{usuario}/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras KPI-Reportes/Gerenciamiento MOS/Panel PBI/bases mensuales/venta/{año}-{mes}"

    # %%
    import pandas as pd
    import os

    # Ruta de la carpeta donde están los archivos
    carpeta = venta

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



    # %%
    df_consolidado.drop(columns={'Unnamed: 0'}, inplace=True)

    # %%



    # %%
    df_consolidado['venta_mean'] = df_consolidado['venta_mean'].astype(str).str.replace(',', '.').astype(float)
    df_consolidado.to_csv(f"C:/Users/{usuario}/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras KPI-Reportes/Gerenciamiento MOS/Panel PBI/automatizacion/consolidado_venta_{mes_1}.csv")
    print(f"\n🎊Proceso finalizado de manera correcta! Archivo guardado en: \nC:/Users/{usuario}/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras KPI-Reportes/Gerenciamiento MOS/Panel PBI/automatizacion/consolidado_fc_{mes_1}.csv")

if __name__ == '__main__':
    main()
