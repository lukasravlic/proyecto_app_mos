def main():    # %%
    import pandas as pd
    import os
    import datetime
    import getpass
    usuario = getpass.getuser()
    from   datetime import timedelta
    


    # %%
    hoy = datetime.datetime.today()


    # %%
    año = str(hoy.year)
    mes = str(hoy.month).zfill(2)

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


    # # %%
    # str((hoy()).month).zfill(2)

    # %%


    # # %%
    # mes_1 = dict_mes.get(str(hoy().month).zfill(2))
    # mes_2 = dict_mes.get(str((hoy() + datetime.timedelta(days=30)).month).zfill(2))
    # mes_3 = dict_mes.get(str((hoy() + datetime.timedelta(days=60)).month).zfill(2))

    # # %%


    # # %%
    # fecha_1 = mes_1 + '-' + str(hoy().year)[2:]
    # fecha_2 = mes_2 + '-' + str(hoy().year)[2:]
    # fecha_3 = mes_3 + '-' + str(hoy().year)[2:]

    # %%
    carpeta_ventas = f"C:/Users/{usuario}/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras OEM/Demanda y New Model Parts/Demanda/Demanda Mainstream/S&OP/{str((hoy-timedelta(days=30)).year).zfill(2)}/{str((hoy-timedelta(days=30)).year).zfill(2)}-{str((hoy-timedelta(days=30)).month).zfill(2)}/OEM Mainstream"
    ruta_ventas = os.listdir(carpeta_ventas)
    for i in ruta_ventas:

        if "PBI" in i:
            print(i)
            venta= pd.read_excel(carpeta_ventas + '/' + i, sheet_name="Mos Venta Data", header=2)
            

    # %%
    ventas_cols = [col for col in venta.columns]  # Remove one set of brackets
    df_prom_venta = venta[ventas_cols]


    # %%
    df_prom_venta

    # %% [markdown]
    # AQUI
    # 

    # %%
    df_prom_venta['venta_mean'] =  df_prom_venta.select_dtypes(include='number').mean(axis=1)


    # %%
    df_prom_venta = df_prom_venta[['Último Eslabón','Nombre Sector MU','venta_mean']]

    df_prom_venta_consolidado = df_prom_venta

    # %% [markdown]
    # venta promedio lista '\n'
    # fc promedio listo
    # determinar "apostrofe" y poner apostrofe en codigo de material (condicional)
    # consolidar mara

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
        import tkinter as tk
        from tkinter import ttk
        from tkcalendar import Calendar
        from datetime import datetime, timedelta
        import pandas as pd
        dfs = []
        df_consolidado_local = pd.DataFrame()

        def obtener_fechas():
            nonlocal dfs, df_consolidado_local

            fecha_inicio = datetime.strptime(cal_inicio.get_date(), '%m/%d/%y')
            fecha_fin = datetime.strptime(cal_fin.get_date(), '%m/%d/%y')
            ventana_cal.destroy()  # Cierra la ventana del calendario

            fecha_actual = fecha_inicio
            while fecha_actual <= fecha_fin:
                df = df_prom_venta.copy()  # Asegúrate de que df_fc_prom esté definido antes
                df['Fecha'] = fecha_actual
                dfs.append(df)
                fecha_actual += timedelta(days=7)

            if dfs:
                df_consolidado_local = pd.concat(dfs, ignore_index=True)
            else:
                df_consolidado_local = pd.DataFrame()

        # Crear ventana principal
        ventana_cal = tk.Tk()
        ventana_cal.title("Seleccionar Rango de Fechas")

        # Fecha de inicio
        tk.Label(ventana_cal, text="Fecha de Inicio").pack()
        cal_inicio = Calendar(ventana_cal, selectmode='day')
        cal_inicio.pack(pady=10)

        # Fecha de fin
        tk.Label(ventana_cal, text="Fecha de Fin").pack()
        cal_fin = Calendar(ventana_cal, selectmode='day')
        cal_fin.pack(pady=10)

        # Botón para confirmar selección
        ttk.Button(ventana_cal, text="Seleccionar", command=obtener_fechas).pack(pady=20)

        ventana_cal.mainloop()
        return df_consolidado_local
    # Llamamos a la función para abrir el selector de fecha
    df_consolidado_venta = seleccionar_fecha()







    # %%
    # Sample DataFrame
    import numpy as np

    # Assign values to a new column based on multiple conditions
    df_consolidado_venta['Tipo'] = np.select(
        [
            df_consolidado_venta['Nombre Sector MU'].isin(['Nacional WBM', 'Mini', 'BMW Motorrad', 'Jaguar', 'Land Rover', 'BMW', 'Porsche', 'Volvo']),
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


    # %%
    folder_path = f"C:/Users/{usuario}/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras KPI-Reportes/Gerenciamiento MOS/Panel PBI/bases mensuales/venta/{año}-{mes}"

    # Check if the folder exists, and create it if it doesn't
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)
        print(f"Carpeta creada: {folder_path}")
    else:
        print(f"La carpeta ya existe: {folder_path}")
    df_consolidado_venta.to_csv(f'{folder_path}/consolidado_venta_{dict_mes.get(str((hoy-timedelta(days=30)).month).zfill(2))}.csv')

    # %%

if __name__ == '__main__':
    main()

