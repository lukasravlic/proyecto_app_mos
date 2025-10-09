def main():
    import pandas as pd
    import os
    import datetime
    import getpass
    usuario = getpass.getuser()

    # %%
    hoy = datetime.datetime.today


    # %%
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
    fecha_1 = mes_1 + '-' + str(hoy().year)[2:]
    fecha_2 = mes_2 + '-' + str(hoy().year)[2:]
    fecha_3 = mes_3 + '-' + str(hoy().year)[2:]

    # %%
    carpeta_venta_premium = f"C:/Users/{usuario}/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras OEM/Demanda y New Model Parts/Demanda/Demanda Premium/S&OP/{año}-{mes}/OEM Premium/"
    ruta_premium = os.listdir(carpeta_venta_premium)
    print(ruta_premium)

    # %%
    ruta_premium

    # %%
    for i in ruta_premium:
        if 'Consolidado - Facturación' in i:
            print(i)
            archivo = carpeta_venta_premium + '/' + i
            venta_premium = pd.read_excel(archivo, sheet_name='MOS Venta Data', header=1)

    # %%
    venta_premium.drop(columns='Unnamed: 0', inplace=True)

    # %%
    venta_premium_cols = [col for col in venta_premium.columns]

    # %%
    venta_premium_cols

    # %%
    from datetime import datetime, timedelta

    hoy = datetime.today()  # o datetime.now()
    mes_n1 = hoy - timedelta(days=30)

    mes_n1 = str(mes_n1.month).zfill(2)

    # %%
    mes_n1

    # %%
    premium = f"C:/Users/{usuario}/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras OEM/Demanda y New Model Parts/Demanda/Demanda Premium/Forecast Colaborado/{año}-{mes_n1}"
    ruta_premium = os.listdir(premium)

    for i in ruta_premium:
        #if 'Premium' in i:
            print(i)
            archivo_premium = premium + '/' +i 
            # print(pd.ExcelFile(archivo_mainstream).sheet_names)
            fc_premium = pd.read_excel(archivo_premium , sheet_name = 'MOS Forecast Data' , header = 3)
    fc_premium.rename(columns={'Último Eslabón':'Último Eslabón y Material SAP'}, inplace=True)
    fc_premium = fc_premium.loc[:, ~fc_premium.columns.duplicated()]
    df_brecha = fc_premium[fc_premium['Familia']=='BRECHA COSTOS'][['Último Eslabón y Material SAP','Marca']]
    df_brecha = df_brecha.assign(
        **{' mar-25':0,
    ' abr-25':0,
    ' may-25':0}
    )

    #df_brecha = df_brecha[['sept-24', 'oct-24', 'nov-24']]


    df_brecha.rename(columns={'Último Eslabón y Material SAP':'Último Eslabón'},inplace=True)

    # %%

    df_prom_venta_premium = venta_premium[venta_premium_cols]


    # %% [markdown]
    # AQUI
    # 

    # %%

    df_prom_venta_premium['venta_mean'] =  df_prom_venta_premium.select_dtypes(include='number').mean(axis=1)


    # %%
    df_prom_venta_premium['Marca'].value_counts()

    # %%
    df_brecha['venta_mean'] =  df_brecha.select_dtypes(include='number').mean(axis=1)

    # %%
    df_brecha = df_brecha[['Último Eslabón','Marca','venta_mean']]

    # %%
    df_prom_venta_premium.rename(columns={'Marca':'Nombre Sector MU'}, inplace=True)
    df_brecha.rename(columns={'Marca':'Nombre Sector MU'}, inplace=True)

    # %%


    # %%

    df_prom_venta_premium = df_prom_venta_premium[['Último Eslabón','Nombre Sector MU','venta_mean']]

    df_prom_venta_consolidado = pd.concat([df_brecha,df_prom_venta_premium])

    # %%
    import tkinter as tk
    from tkinter import ttk
    from tkcalendar import Calendar
    from datetime import datetime, timedelta
    import pandas as pd

    # DataFrame base (reemplaza 'df_fc_prom' con tu DataFrame completo)

    dfs = []
    df_consolidado = pd.DataFrame()  # Definir df_consolidado globalmente
    dfs_venta = []
    def seleccionar_fecha():
        def obtener_fecha():
            global df_consolidado , df_consolidado_venta # Declarar que usamos la variable global

            # Convertimos la fecha seleccionada a un objeto datetime
            fecha_seleccionada = datetime.strptime(cal.get_date(), '%m/%d/%y')
            ventana_cal.destroy()  # Cierra la ventana del calendario

            # Obtenemos el año y mes de la fecha seleccionada
            mes_inicial = fecha_seleccionada.month

            # Iniciamos el loop con la fecha seleccionada y continuamos hasta que cambie de mes
            fecha_actual = fecha_seleccionada
            while fecha_actual.month == mes_inicial:
                # Crear una copia completa de df_fc_prom y agregar la fecha
                df2 = df_prom_venta_consolidado.copy()
                #df = df_fc_prom_consolidado.copy()
                #df['Fecha'] = fecha_actual  # Asigna la misma fecha a todas las filas de este DataFrame
                df2['Fecha'] = fecha_actual
                # Agregar el DataFrame al listado de dfs
                #dfs.append(df)
                dfs_venta.append(df2)

                # Incrementar la fecha en 7 días
                fecha_actual += timedelta(days=7)

            # Concatenar todos los DataFrames de dfs en un único DataFrame
            #df_consolidado = pd.concat(dfs, ignore_index=True)
            df_consolidado_venta = pd.concat(dfs_venta, ignore_index=True)


        # Crear ventana de selección de fecha
        ventana_cal = tk.Tk()
        ventana_cal.title("Seleccionar Fecha de Inicio")

        cal = Calendar(ventana_cal, selectmode='day')
        cal.pack(pady=20)

        # Botón para confirmar selección
        ttk.Button(ventana_cal, text="Seleccionar", command=obtener_fecha).pack(pady=10)

        ventana_cal.mainloop()

    # Llamamos a la función para abrir el selector de fecha
    seleccionar_fecha()







    # %%
    # Sample DataFrame
    import numpy as np

    # Assign values to a new column based on multiple conditions
    df_consolidado_venta['Tipo'] = np.select(
        [
            df_consolidado_venta['Nombre Sector MU'].isin(['Nacional WBM','Harley Davidson','Nacional Ditec', 'Mini', 'BMW Motorrad', 'Jaguar', 'Land Rover', 'BMW', 'Porsche', 'Volvo']),
            df_consolidado_venta['Nombre Sector MU'].isin(['Subaru', 'Geely', 'DFSK'])
        ],
        [
            'OEM Premium',  # Value if condition for OEM Premium is True
            'OEM Inchcape'  # Value if condition for OEM Inchcape is True
        ],
        default='OEM Derco'  # Value if none of the conditions are True
    )

    print(df_consolidado_venta)



    # %%
    df_consolidado_venta['Último Eslabón'] = df_consolidado_venta['Último Eslabón'].astype('str')

    # %%
    def append_material_venta(row):
        if row['Tipo'] == 'OEM Inchcape':
            return row['Último Eslabón'] + 'INP300'
        elif row['Tipo'] == 'OEM Derco':
            return row['Último Eslabón'] + 'R3'
        else:
            return row['Último Eslabón']

    # Apply the function to modify the column
    df_consolidado_venta['Último Eslabón'] = df_consolidado_venta.apply(append_material_venta, axis=1)

    # %%
    mes_1


    # %%
    
    

    # %%
    folder_path = f"C:/Users/{usuario}/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras KPI-Reportes/Gerenciamiento MOS/Panel PBI/bases mensuales/venta/{año}-{mes}"

    # Check if the folder exists, and create it if it doesn't
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)
        print(f"Carpeta creada: {folder_path}")
    else:
        print(f"La carpeta ya existe: {folder_path}")
    df_consolidado_venta.to_csv(f'{folder_path}/consolidado_venta_premium_{mes_1}.csv')


if __name__ == '__main__':
    main()