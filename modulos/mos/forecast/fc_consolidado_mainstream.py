## %%
def main():
    import pandas as pd
    import os
    import datetime
    import getpass
    usuario = getpass.getuser()

    # %%
    hoy = datetime.datetime.today()


    # %%
    año = str(hoy.year)
    mes = str(hoy.month).zfill(2)
    mes_n1 = str(hoy.month-1).zfill(2)

    # %%
    mes_n1

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
    dict_mes_archivos = {
        '01': 'Enero',
        '02': 'Febrero',
        '03': 'Marzo',
        '04': 'Abril',
        '05': 'Mayo',
        '06': 'Junio',
        '07': 'Julio',
        '08': 'Agosto',
        '09': 'Septiembre',
        '10': 'Octubre',
        '11': 'Noviembre',
        '12': 'Diciembre',
    }

    mes_n1_nombre = dict_mes_archivos.get(mes_n1)


    # %%


    # %%
    mes_1 = dict_mes.get(str(hoy.month).zfill(2))
    mes_2 = dict_mes.get(str((hoy + datetime.timedelta(days=30)).month).zfill(2))
    mes_3 = dict_mes.get(str((hoy + datetime.timedelta(days=60)).month).zfill(2))

    # %%
    fecha_1 = mes_1 + '-' + str(hoy.year)[2:]
    fecha_2 = mes_2 + '-' + str(hoy.year)[2:]
    fecha_3 = mes_3 + '-' + str(hoy.year)[2:]

    # %%
    ruta_bases = f"C:/Users/{usuario}/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras OEM/Demanda y New Model Parts/Demanda"
    bases = os.listdir(ruta_bases)

    # %% [markdown]
    # MAINSTREAM

    # %%
    mainstream =f"C:/Users/{usuario}/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras OEM/Demanda y New Model Parts/Demanda/Demanda Mainstream/Forecast Colaborado/{año}/{año}-{mes_n1} {mes_n1_nombre}"

    ruta_mainstream = os.listdir(mainstream)


    # %%
    for i in ruta_mainstream:
        if 'OEM' in i:

            archivo_mainstream = mainstream + '/' +i 
            print(f'\n📄Archivo usado: {archivo_mainstream}')

            fc = pd.read_excel(archivo_mainstream , sheet_name = 'MOS Forecast Data' , header = 3)

    # %%
    fc.rename(columns={'Último Eslabón':'Último Eslabón y Material SAP'}, inplace=True)


  
    # %%


    # %%
    fc_cols = ['Último Eslabón y Material SAP','Marca'] + [col for col in fc.columns if 'FC' in col and 'Prom' not in col]


    # %%
    fc_cols

    # %%
    fc_cols_prom = fc_cols[:8]# %%

    df_fc_prom= fc[fc_cols_prom].copy()
    df_fc_prom['fc_mean'] = df_fc_prom.select_dtypes(include='number').mean(axis=1)
    df_fc_prom= df_fc_prom[['Último Eslabón y Material SAP','Marca', 'fc_mean']]# %%

    



    df_consolidado = pd.DataFrame()  # Definir df_consolidado globalmente

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
                df = df_fc_prom.copy()  # Asegúrate de que df_fc_prom esté definido antes
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

    # Llamamos a la función para abrir el selector de fecha y obtener el DataFrame
    df_consolidado = seleccionar_fecha()





    


    # %%
    # Sample DataFrame
    import numpy as np
   
    # Assign values to a new column based on multiple conditions
    df_consolidado['Tipo'] = np.select(
        [
            df_consolidado['Marca'].isin(['Nacional WBM', 'Mini', 'BMW Motorrad', 'Jaguar', 'Land Rover', 'BMW', 'Porsche', 'Volvo']),
            df_consolidado['Marca'].isin(['Subaru', 'Geely', 'DFSK'])
        ],
        [
            'OEM Premium',  # Value if condition for OEM Premium is True
            'OEM Inchcape'  # Value if condition for OEM Inchcape is True
        ],
        default='OEM Derco'  # Value if none of the conditions are True
    )

    df_consolidado.head(2)
    # %%
    df_consolidado['Último Eslabón y Material SAP'] = df_consolidado['Último Eslabón y Material SAP'].astype('str')

    cad_remp = pd.read_csv(f"C:/Users/{usuario}/Inchcape/Planificación y Compras Chile - Documentos/Bases Indicadores en CSV {año}-{mes}/COD_ACTUAL.csv", usecols=['Nro_pieza_fabricante_1','Cod_Actual_1'])
    df_consolidado_ue = df_consolidado.merge(cad_remp, left_on='Último Eslabón y Material SAP', right_on='Nro_pieza_fabricante_1', how='left')

    df_consolidado_ue['Cod_Actual_1'].fillna(df_consolidado_ue['Último Eslabón y Material SAP'], inplace=True)
    df_consolidado_ue.drop(columns='Último Eslabón y Material SAP', inplace=True)
    df_consolidado_ue.drop(columns='Nro_pieza_fabricante_1', inplace=True)
    df_consolidado_ue.rename(columns={'Cod_Actual_1':'Último Eslabón y Material SAP'}, inplace=True)
    df_consolidado_ue = df_consolidado_ue[['Último Eslabón y Material SAP'] + [col for col in df_consolidado_ue.columns if col != 'Último Eslabón y Material SAP']]
    # %%

    # %%
    def append_material(row):
        if row['Tipo'] == 'OEM Inchcape':
            return row['Último Eslabón y Material SAP'] + 'INP300'
        elif row['Tipo'] == 'OEM Derco':
            return row['Último Eslabón y Material SAP'] + 'R3'
        else:
            return row['Último Eslabón y Material SAP']

    # Apply the function to modify the column
    df_consolidado_ue['Último Eslabón y Material SAP'] = df_consolidado_ue.apply(append_material, axis=1)
    # %%
    # Define the path for the folder
    folder_path = f"C:/Users/{usuario}/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras KPI-Reportes/Gerenciamiento MOS/Panel PBI/bases mensuales/forecast/{año}-{mes}"

    # Check if the folder exists, and create it if it doesn't
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)
        print(f"\n📂Carpeta creada, archivo se guardará en: {folder_path}")
    else:
        print(f"\n📂La carpeta ya existe, el archivo será guardado en: {folder_path}")

    df_consolidado_ue.to_csv(f'{folder_path}/consolidado_fc_{mes_n1_nombre}.csv')
    print("\n🎊Proceso finalizado de manera correcta!")


if __name__ == '__main__':
    main()