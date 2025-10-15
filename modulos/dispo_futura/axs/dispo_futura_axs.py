def main(): 
    # %%
    #IMPORTACION DE LIBRERIAS
    import pandas as pd
    import datetime
    import os
    import numpy as np
    import getpass

    #hoy = datetime.datetime.today()
    #hoy = datetime.date(2024, 10,14)
    #LECTURA DE DFS
    from pathlib import Path
    usuario = getpass.getuser()

    # %%


    # %%
    import tkinter as tk
    from tkinter import ttk
    from tkcalendar import DateEntry
    import datetime

    # Variable global para almacenar la fecha seleccionada
    

    # Función que captura la fecha seleccionada y cierra la ventana
    def seleccionar_y_continuar():
        global fecha_seleccionada
        
        # Obtener la fecha seleccionada como un objeto datetime.date
        fecha_input = calendario.get_date()

        # Convertir a datetime.date
        fecha_seleccionada = fecha_input
        
        # Cerrar la ventana
        ventana.destroy()

    # Crear la ventana principal
    ventana = tk.Tk()
    ventana.title("Selección de Fecha")
    ventana.geometry("300x250")

    # Etiqueta de instrucción
    label_instruccion = tk.Label(ventana, text="Selecciona una fecha:")
    label_instruccion.pack(pady=10)

    # Calendario de selección de fecha
    calendario = DateEntry(ventana, date_pattern='dd.mm.yyyy', background='darkblue', foreground='white', borderwidth=2)
    calendario.pack(pady=10)

    # Botón para capturar la fecha y continuar
    boton_ok = ttk.Button(ventana, text="OK", command=seleccionar_y_continuar)
    boton_ok.pack(pady=10)

    # Iniciar la aplicación
    ventana.mainloop()

    # Una vez que la ventana se cierra, la fecha ya está disponible como un objeto datetime.date
    print(f"Fecha seleccionada: {fecha_seleccionada}")

    # Aquí puedes continuar con el resto del código
    # Ejemplo: 
    # print(f"Usando la fecha seleccionada: {fecha_seleccionada}")



    # %%
    hoy = fecha_seleccionada


    # %%

    ruta = f'C:/Users/{usuario}/Inchcape/Planificación y Compras Chile - Documentos/Bases Indicadores en CSV {hoy.year}-{hoy.month:02d}'
    ruta_repo = Path(ruta)

    # %%
    import pandas as pd

    def excel_to_dataframe(xl_name: str, sh_name: str) -> pd.DataFrame:
        """
        Convert an Excel sheet to a pandas DataFrame.

        :param xl_name: The path to the Excel file.
        :param sh_name: The name of the sheet to be read.
        :return: A pandas DataFrame containing the data from the specified Excel sheet.
        """
        # Load the Excel file
        xls = pd.ExcelFile(xl_name)
        
        # Parse the specified sheet into a DataFrame
        df = xls.parse(sh_name)
        
        return df

    # Example usage:



    # %%
    dtypes = {'Material actual':'str'}

    # %%
    ruta_sp = ruta_repo.joinpath('SP.csv')

    df_sp = pd.read_csv(ruta_sp , dtype = dtypes, low_memory=False)

    # %%
    columnas= ['Nro_pieza_fabricante_1',	'Cod_Actual_1']
    ruta_cod = ruta_repo.joinpath('COD_ACTUAL.csv')

    # Leer el archivo CSV en un DataFrame
    cadena_de_remplazo = pd.read_csv(ruta_cod)
    cadena_de_remplazo = cadena_de_remplazo[columnas]


    # %%
    #MARA
    columnas_mara = ['Material_R3','Part_number','Material_dsc','Modelo','Familia', 'Subfamilia', 'Categoría', 'Subcatgería','Sector_dsc']
    ruta_mara = ruta_repo.joinpath('MARA_R3.csv')

    # Leer el archivo CSV en un DataFrame
    df_mara = pd.read_csv(ruta_mara, dtype={'Part_number':'str'})

    print('Ruta Mara: ' + '\n' + str(ruta_mara))

    # %%

    #OBSOLECENCIA
    #OBSOLECENCIA
    columnas = ['ZFI_INNV1_T','ZFI_INNV2_T','ZFI_INNV3_T','ZFI_INNV4_T','ZFI_INNV5_T','ZFI_INNV6_T','ZFI_INNV7_T','sociedad_orig','Último Eslabón','Centro','obso_inchcape']

    # for i in lista_maestro:
    #     if 'new_obso' in i:
    #         carpeta_obso = ruta_maestro + '/' + i
    # df_obs_1 = excel_to_dataframe(carpeta_obso,'Base Obs Cierre Abr-24')

    ruta_obs = ruta_repo.joinpath('OBSOLECENCIA.csv')

    # Leer el archivo CSV en un DataFrame
    df_obs_1 = pd.read_csv(ruta_obs)
    df_obs_1 = df_obs_1[columnas]

    # %%
    ruta_lt = f"C:/Users/{usuario}/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras KPI-Reportes/Disponibilidad Futura/2024/AXS/bases_python/LT Actuales Mar-24.xlsx"
    df_lt = pd.read_excel(ruta_lt, header=1)


    import tkinter as tk
    from tkinter import filedialog
    import pandas as pd
    import os

    # Crear la ventana principal oculta (necesaria para abrir el explorador de archivos)
    root = tk.Tk()
    root.withdraw()  # Ocultar la ventana principal de tkinter

    # Abrir un cuadro de diálogo para seleccionar el archivo de stock
    archivo_tubo = filedialog.askopenfilename(
        title="Selecciona el archivo de Stock",
        filetypes=(("Archivos de Excel", "*.xlsx"), ("Todos los archivos", "*.*"))
    )

    # Verificar si se seleccionó algún archivo
    if archivo_tubo:
        print(f"Archivo de Stock seleccionado: {archivo_tubo}")
        dtypes = {'Almacén': 'str', 'Centro': 'str'}
        
        # Leer el archivo seleccionado
        df_stock = pd.read_excel(archivo_tubo, dtype=dtypes, sheet_name='Sheet1')
        print("Archivo de Stock cargado correctamente.")
    else:
        print("No se seleccionó ningún archivo de Stock.")

    # Abrir un cuadro de diálogo para seleccionar el archivo de TR (Transito)
    archivo_tr = filedialog.askopenfilename(
        title="Selecciona el archivo de TR FINAL R3 Consolidado",
        filetypes=(("Archivos de Excel", "*.xlsx"), ("Todos los archivos", "*.*"))
    )

    # Verificar si se seleccionó algún archivo
    if archivo_tr:
        print(f"Archivo de TR seleccionado: {archivo_tr}")
        
        # Leer el archivo seleccionado
        df_tr = pd.read_excel(archivo_tr, sheet_name='Sheet1')
        print("Archivo de TR cargado correctamente.")
    else:
        print("No se seleccionó ningún archivo de TR.")



    # #STOCK
    # ruta_tubo = f"C:/Users/{usuario}/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras KPI-Reportes/Tubo Semanal"
    # carpetas_tubo = os.listdir(ruta_tubo)
    # tubo = carpetas_tubo[-2]
    # ubi_tubo = ruta_tubo + '/' + tubo + '/' + tubo
    # archivo_tr = ruta_tubo + '/' + tubo + '/' + tubo + ' - TR FINAL R3 Consolidado.xlsx'
    # archivo_tubo = ruta_tubo + '/' + tubo + '/' + tubo + ' - Stock R3.xlsx'
    # print(tubo)
    # dtypes = {'Almacén':'str', 'Centro':'str'}
    # # dtypes = {'Almacén':'str'}
    # df_stock = pd.read_excel(archivo_tubo, dtype = dtypes,sheet_name = 'Sheet1')


    # # #TRANSITO
    # df_tr = pd.read_excel(archivo_tr,sheet_name = 'Sheet1')

    # %%
    #Reseteo de OBS
    df_sp_1 = df_sp
    df_obs = df_obs_1

    # %%
    df_obs_1 = df_obs_1.rename(columns={'Último Eslabón': 'Ultimo Eslabon'}, inplace = True)

    # %%
    df_sp_1 = df_sp_1.merge(cadena_de_remplazo, left_on='Material actual', right_on='Nro_pieza_fabricante_1', how='left')
    df_sp_1['Cod_Actual_1'] = df_sp_1['Cod_Actual_1'].fillna(df_sp_1['Material actual'])
    df_sp_1 = df_sp_1.drop('Nro_pieza_fabricante_1', axis=1)

    # %%
    df_codigo = df_sp_1[df_sp_1['Vigencia Derco']==1]

    # %%
    df_codigo.drop_duplicates(subset='Material actual', inplace=True)

    # %%
    df_mara.drop_duplicates(subset='Material_R3', inplace=True)

    # %%


    # %%
    df_codigo = df_codigo[['Material actual','Cod_Actual_1','Descripción','Cod. Proveedor','Proveedor','Plan Híbrido 2 UN v2','Origen','CES','PDI','Nuevo','Lead Time']]

    # %%
    df_base = pd.merge(df_codigo, df_mara, left_on = 'Cod_Actual_1', right_on='Material_R3', how='left')
    df_base['Part_number'] = df_base['Part_number'].str.replace(r'\[\#\]', '', regex=True)




    anio = hoy.year if hoy.month > 1 else hoy.year - 1
    mes = hoy.month - 1 if hoy.month > 1 else 12

    # Formatear mes como dos dígitos
    mes_formateado = str(mes).zfill(2)

    # Ruta base
    # Cambia esto al nombre del usuario correspondiente
    ruta_fc = f"C:/Users/{usuario}/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras OEM/Demanda y New Model Parts/Demanda/Demanda Mainstream/Forecast Colaborado/{anio}"

    # Listar carpetas en la ruta base
    try:
        lista_fc = os.listdir(ruta_fc)
        for i in lista_fc:
            if str(anio) in i and mes_formateado in i:
                archivos_fc = os.listdir(f"{ruta_fc}/{i}")
                for j in archivos_fc:
                    print(j)
                    if 'AXS' in j:
                        archivo = f"{ruta_fc}/{i}/{j}"
                        df_fc = pd.read_excel(archivo, sheet_name='MOS Forecast Data', header=3)
    except FileNotFoundError:
        print(f"La ruta {ruta_fc} no existe. Verifica el nombre del usuario o la e")
    df_fc_final = ['Último Eslabón']
    for item in df_fc.columns.to_list():
        if "FC" in item:
            df_fc_final.append(item)
    df_fc_final = df_fc_final[:13]
    df_fc_prom = df_fc[df_fc_final]

    df_fc_prom = df_fc_prom.merge(cadena_de_remplazo, left_on='Último Eslabón', right_on='Nro_pieza_fabricante_1', how ='left')
    df_fc_prom['Cod_Actual_1'] = df_fc_prom['Cod_Actual_1'].fillna(df_fc_prom['Último Eslabón'])

    df_fc_prom.drop(columns=['Último Eslabón','Nro_pieza_fabricante_1'], inplace=True)
    columnas_prom = [col for col in df_fc_prom.columns if 'Suma' in col]
    df_fc_prom
    df_fc_prom.columns = [col.replace('Suma de', 'FC') for col in df_fc_prom.columns]
    df_fc_prom.columns = [col.replace('sept', 'sep')  for col in df_fc_prom.columns]
    # columnas_seleccionadas = ['Cod_Actual_1'] + [col for col in df_fc_prom.columns if 'FC' in col and 'Prom' not in col][:10]

    # nuevo_df_fc_prom = df_fc_prom[columnas_seleccionadas].copy()
    df_fc_prom = df_fc_prom.groupby('Cod_Actual_1').sum()
    df_fc_prom = df_fc_prom.reset_index()

    df_fc_prom.columns = [col[:-1] if col.startswith('FC') else col for col in df_fc_prom.columns]


    df_fc_prom.columns

    # %%
    # Filter columns that start with 'FC'
    columns_to_process = [col for col in df_fc_prom.columns if col.startswith('FC')]
    df_fc_prom
    columns_to_process
    df_fc_prom[columns_to_process[0]]
    df_fc_prom[columns_to_process[1:7]]

    df_fc_prom['Promedio FC'] = df_fc_prom[columns_to_process[1:7]].mean(axis=1)
    df_fc_prom['Promedio FC Piso'] = df_fc_prom[columns_to_process[0:3]].mean(axis=1)

    #Multiply the first column by 0.33 and add the result to each of the next three columns
    multiplied_column = df_fc_prom[columns_to_process[0]] * 0.333
    for col in columns_to_process[1:4]:
        df_fc_prom[col] = df_fc_prom[col] + multiplied_column
        #print(col)



    for col in columns_to_process:
        df_fc_prom[col] = df_fc_prom[col]/4.33






    #Drop the first column
    df_fc_prom['Stock de piso'] = df_fc_prom[columns_to_process[0]] 
    df_fc_prom.drop(columns_to_process[0], axis=1, inplace=True)

    # %%


    # %%
    df_fc_final_venta = ['Último Eslabón']
    for item in [col for col in df_fc.columns if 'Vta R' in col and not 'Vigencia' in col][:6]:
        print(item)
        df_fc_final_venta.append(item)


    # %%
    df_fc_final_venta

    # %%
    df_fc_prom_venta = df_fc[df_fc_final_venta]

    # %%
    df_fc_prom_venta = df_fc_prom_venta.merge(cadena_de_remplazo, left_on='Último Eslabón', right_on='Nro_pieza_fabricante_1', how ='left')
    df_fc_prom_venta['Cod_Actual_1'] = df_fc_prom_venta['Cod_Actual_1'].fillna(df_fc_prom_venta['Último Eslabón'])
    df_fc_prom_venta.drop(columns=['Último Eslabón','Nro_pieza_fabricante_1'], inplace=True)
    columnas_prom_venta = [col for col in df_fc_prom_venta.columns if 'Vta R' in col][:6]
    df_fc_prom_venta['Promedio Venta'] = df_fc_prom_venta[columnas_prom_venta].mean(axis=1)
    # columnas_seleccionadas = ['Cod_Actual_1'] + [col for col in df_fc_prom.columns if 'FC' in col and 'Prom' not in col][:10]
    # nuevo_df_fc_prom = df_fc_prom[columnas_seleccionadas].copy()


    df_fc_prom_venta = df_fc_prom_venta[['Cod_Actual_1', 'Promedio Venta']]
    df_fc_prom_venta = df_fc_prom_venta.groupby(['Cod_Actual_1'])['Promedio Venta'].sum().reset_index()

    # %%
    df_stock['Total'] = df_stock['Libre utilización'] + df_stock['Trans./Trasl.'] + df_stock['En control calidad']

    # Eliminar las columnas no necesarias
    columns_to_drop = ['Libre utilización', 'Trans./Trasl.', 'En control calidad']
    df_stock = df_stock.drop(columns=columns_to_drop)

    # Filtrar las filas que cumplen con las condiciones especificadas
    condicion = (
        ((df_stock['Centro'] == '201') & (df_stock['Almacén'] == '1100'))
    )

    condicion_2 = (
        ((df_stock['Almacén'] == '1500') | (df_stock['Almacén'] == '1505')))


    df_stock_cd = df_stock[condicion]
    df_stock_pañol = df_stock[condicion_2]

    # Agrupar por 'Ult. Eslabon' y sumar la columna 'Total'
    df_stock_cd = df_stock_cd.groupby(['Ult. Eslabon']).agg({'Total': 'sum'}).reset_index()
    df_stock_pañol = df_stock_pañol.groupby(['Ult. Eslabon']).agg({'Total': 'sum'}).reset_index()

    # %%
    df_base = df_base.merge(df_stock_cd, left_on='Cod_Actual_1', right_on='Ult. Eslabon', how='left')
    df_base = df_base.merge(df_stock_pañol, left_on='Cod_Actual_1', right_on='Ult. Eslabon', how='left')


    # %%
    columnas = {'Total_x':'Stock CD','Total_y':'Stock Pañol'}
    columnas_drop = ['Ult. Eslabon_x','Ult. Eslabon_y']
    df_base.rename(columns=columnas, inplace=True)
    df_base.drop(columns = columnas_drop, inplace=True)


    # %%
    df_fc_prom

    # %%
    df_base = df_base.merge(df_fc_prom, left_on='Cod_Actual_1', right_on = 'Cod_Actual_1', how='left')
    df_base = df_base.merge(df_fc_prom_venta, left_on='Cod_Actual_1', right_on = 'Cod_Actual_1', how='left')

    # %%
    def custom_formula(row):
        AH = row['Stock CD']
        T = row['Promedio FC']
        if AH == 0:
            return 0 if T == 0 else (AH / T)
        else:
            return AH / T if T != 0 else 12

    # Apply the custom function to the DataFrame
    df_base['Cobertura Stock'] = df_base.apply(custom_formula, axis=1)

    # %%
    df_base['Cobertura Stock'].sum()

    # %%
    df_sp_costo_1 = df_sp_1[['Material actual','Costo Un']]
    df_sp_costo_1.drop_duplicates(subset=['Material actual'], inplace=True)
    df_base = df_base.merge(df_sp_costo_1[['Material actual','Costo Un']], left_on='Material actual', right_on='Material actual', how= 'left')
    df_sp_costo = df_sp_1[['Cod_Actual_1','Costo Un']].sort_values(by='Costo Un', ascending=False)
    df_sp_costo.drop_duplicates(subset=['Cod_Actual_1'], keep='first', inplace=True)
    df_base = df_base.merge(df_sp_costo, left_on='Cod_Actual_1', right_on='Cod_Actual_1', how='left')
    df_base['Costo Un_x'] = df_base['Costo Un_x'].fillna(df_base['Costo Un_y'])
    df_base.drop(['Costo Un_y'], inplace = True, axis=1)
    df_base = df_base.rename(columns = {'Costo Un_x':'Costo CPP'})


    # %%
    df_sp_moq_1 = df_sp_1[['Material actual','Cod_Actual_1','MOQ']]
    df_sp_moq_1.drop_duplicates(subset=['Material actual'],keep='first', inplace=True)
    df_base = df_base.merge(df_sp_moq_1[['Material actual','MOQ']], left_on='Material actual', right_on='Material actual', how= 'left')
    df_sp_moq = df_sp_1[['Cod_Actual_1','MOQ']].sort_values(by='MOQ',ascending=False)
    df_sp_moq.drop_duplicates(subset=['Cod_Actual_1'], keep='first', inplace=True)
    df_base = df_base.merge(df_sp_moq, left_on='Cod_Actual_1', right_on='Cod_Actual_1', how='left')
    df_base['MOQ_x'] = df_base['MOQ_x'].fillna(df_base['MOQ_y'])

    df_base = df_base.rename(columns = {'MOQ_x':'MOQ'})
    df_base.drop(columns='MOQ_y', inplace=True)




    # %%
    df_sp_precio_1 = df_sp_1[['Material actual','Cod_Actual_1','Precio Neto', 'Moneda']]
    df_sp_precio_1.drop_duplicates(subset=['Material actual'],keep='first', inplace=True)
    df_base = df_base.merge(df_sp_precio_1[['Material actual','Precio Neto','Moneda']], left_on='Material actual', right_on='Material actual', how= 'left')
    df_sp_precio = df_sp_1[['Cod_Actual_1','Precio Neto', 'Moneda']].sort_values(by='Precio Neto',ascending=False)
    df_sp_precio.drop_duplicates(subset=['Cod_Actual_1'], keep='first', inplace=True)
    df_base = df_base.merge(df_sp_precio, left_on='Cod_Actual_1', right_on='Cod_Actual_1', how='left')
    df_base['Precio Neto_x'] = df_base['Precio Neto_x'].fillna(df_base['Precio Neto_y'])
    df_base['Moneda_x'] = df_base['Moneda_x'].fillna(df_base['Moneda_y'])

    df_base = df_base.rename(columns = {'Precio Neto_x':'Precio Neto', 'Moneda_x':'Moneda'})
    df_base.drop(columns=['Precio Neto_y','Moneda_y'], inplace=True)

    # %%
    df_base['Moneda'].value_counts()

    # %%
    df_base['Vigencia Canal'] = df_base.apply(lambda row: 'PDI' if row['PDI'] == 1 else ('CES' if row['CES'] == 1 else 'NUEVO'), axis=1)


    # %%
    df_base.fillna(0, inplace=True)

    # %%
    df_base['Menor a 1'] = df_base.apply(lambda row: 'Menor que 1' if row['Promedio Venta'] < 1 else 'Mayor o igual a 1', axis=1)


    # %%
    # hoy = datetime.date(2024, 6, 5)
    #hoy = pd.to_datetime(hoy)

    # Add the 'Lead Time' to today's date to create the 'Semana LT' column
    df_base['Semana LT'] = (pd.Timestamp(hoy) + pd.to_timedelta(df_base['Lead Time'], unit='D')).apply(lambda x: x.isocalendar().week)

    df_base['Mes LT'] = (pd.Timestamp(hoy) + pd.to_timedelta(df_base['Lead Time'], unit='D')).dt.month



    # %%
    df_base.groupby(['Mes LT'])['Cod_Actual_1'].count().reset_index().to_clipboard()

    # %%
    din_obs = df_obs[(df_obs['ZFI_INNV1_T'] == 'CHILE') & 
                    (df_obs['ZFI_INNV2_T'] == 'BACK OFFICE PAISES') & 
                    (df_obs['ZFI_INNV3_T'] == 'SOPORTE PAIS') & 
                    (df_obs['ZFI_INNV4_T'] == 'SOPORTE PAIS') & 
                    (df_obs['ZFI_INNV5_T'] == 'SOPORTE PAIS') & 
                    (df_obs['ZFI_INNV6_T'] == 'OPERACIONES Y LOGIST') & 
                    (df_obs['ZFI_INNV7_T'] == 'PLANIFICACIN Y ABAST') & 
                    (df_obs['sociedad_orig'] == 'CL02') &
                    (df_obs['Centro'] == '0201')]


    # %%
    din_obs_final = din_obs.groupby('Ultimo Eslabon').sum(['obso_inchcape']).reset_index()

    # %%
    df_base.columns

    # %% [markdown]
    # NO SE ENCUENTRAN LOS ARCHIVOS NI DE LT NI OBSOLECENCIA

    # %%
    df_base = df_base.merge(din_obs_final,left_on='Cod_Actual_1', right_on='Ultimo Eslabon', how='left')

    # %%
    df_base.fillna(0, inplace=True)

    # %%
    df_base['Obsolescencia'] = np.where(df_base['obso_inchcape'].notna() & (df_base['obso_inchcape'] > 0), 1, 0)

    # %%
    df_base['obso_inchcape'].sum()

    # %% [markdown]
    # OBTENCION TRANSITO Y TUBO

    # %%
    cl_doc = ['ZIPL','ZSTO','ZSPT']
    # Assuming your DataFrame is named df_tr
    # Assuming 'año' and 'semanas' are already present in the DataFrame

    # Apply filters to the DataFrame if needed


    # Create a pivot table with 'year' and 'week' as index columns



    filtered_df = df_tr[df_tr['Cl.documento compras'].isin(cl_doc)]
    filtered_df = filtered_df[['Material','Cantidad','Fecha']]
    filtered_df.reset_index(drop=True)


    # %%
    df_base_2 =df_base

    # %%
    from datetime import timedelta
    from datetime import timedelta, date
    import pandas as pd
    #current_date = datetime.date(2024,8,28)
    #current_date = datetime.datetime.today()
    current_date = hoy
    print(current_date.isocalendar())
    # Crear las columnas en base a las próximas 39 semanas en la base de datos 'b'
    # Define the custom ISO week function again
    def get_iso_week(date_obj):
        # Ensure date_obj is a datetime.date object
        # (though in this specific loop, it should already be date objects)
        if isinstance(date_obj, pd.Timestamp): # Keep this for robustness if used elsewhere
            date_obj = date_obj.date()

        iso_year, iso_week, _ = date_obj.isocalendar()

        # Define the specific date range for week 1 of 2026
        start_date_range = date(2025, 12, 29)
        end_date_range = date(2026, 1, 4)

        # Check if the date falls within the special week 1, 2026 range
        if start_date_range <= date_obj <= end_date_range:
            return 2026, "01" # Return year as int, week as string
        else:
            return iso_year, f"{iso_week:02d}"

    # Create a placeholder DataFrame for demonstration

    nombre_meses = {
        1: 'jan', 2: 'feb', 3: 'mar', 4: 'apr', 5: 'may', 6: 'jun',
        7: 'jul', 8: 'aug', 9: 'sep', 10: 'oct', 11: 'nov', 12: 'dec'
    }

    def nombrar_mes(mes_num):
        return nombre_meses.get(mes_num)

    # Crear las columnas en base a las próximas 39 semanas en la base de datos 'df_base'
    for i in range(39):
        week_start_date = hoy + timedelta(weeks=i)

        # Use the custom get_iso_week function to get the year and week number
        year, week_number_str = get_iso_week(week_start_date)

        # Determine the month name based on the original date's month,
        # or adjust if the ISO week shifted the year
        # For simplicity and to match previous logic, we'll use the month of the week_start_date
        month_name = nombrar_mes(week_start_date.month)

        column_name = f"{year}-{month_name}-{week_number_str}"
        print(column_name)

        df_base[column_name] = 0 # Inicializar todas las columnas con 0


    # %%
    df_base.rename(columns={'2026-dec-01':'2026-jan-01'},inplace=True)

    # %%
    # filtered_df['Año'] = filtered_df['Fecha'].dt.year
    # filtered_df['Month'] = filtered_df['Fecha'].dt.strftime('%B').str.lower().str[:3]
    # filtered_df['Semana'] = filtered_df['Fecha'].dt.isocalendar().week



    # %%
    filtered_df = filtered_df[filtered_df['Cantidad']>0]


    # %%
    import datetime

    def get_iso_week(date_obj):
        # Ensure date_obj is a datetime.date object
        if isinstance(date_obj, pd.Timestamp):
            date_obj = date_obj.date()

        # Get the ISO week for the given date
        iso_year, iso_week, _ = date_obj.isocalendar()

        # Define the specific date range for week 1 of 2026
        start_date_range = date(2025, 12, 29)
        end_date_range = date(2026, 1, 4)

        # Check if the date falls within the special week 1, 2026 range
        if start_date_range <= date_obj <= end_date_range:
            return 2026, "01"
        else:
            # For dates outside the special range, use the standard ISO week
            return iso_year, f"{iso_week:02d}" # Asegura que la semana tenga dos dígitos

    # Actualizar la columna de semana y año
    filtered_df['Año'], filtered_df['Semana'] = zip(*filtered_df['Fecha'].apply(lambda x: get_iso_week(x)))

    # Función para obtener el mes
    def get_month(year, week):
        return datetime.datetime.strptime(f'{year}-W{int(week)}-1', "%Y-W%W-%w").strftime('%B').lower()[:3]

    # Aplicar la función de mes
    filtered_df['Month'] = filtered_df.apply(lambda row: get_month(row['Año'], row['Semana']), axis=1)

    # Ahora df_base debe tener las ventas cruzadas en las columnas correspondientes
    # Primero, agrupamos las ventas por material, año, mes y semana
    grouped_sales = filtered_df.groupby(['Material', 'Año', 'Month', 'Semana'])['Cantidad'].sum().reset_index()
    grouped_sales['Año'] = grouped_sales['Año'].astype('str')
    grouped_sales['Semana'] = grouped_sales['Semana'].astype('int')  # Asegurarse de que Semana sea entero

    # Luego, cruzamos los datos de ventas en df_base
    for index, row in grouped_sales.iterrows():

        product_code = row['Material']
        week_number = int(row['Semana'])  # Asegurar que sea un entero
        year = row['Año']
        column_name_pattern = f"{year}-{week_number:02d}"
        
        # Encuentra la columna en df_base que coincida exactamente con el patrón
        matching_columns = [col for col in df_base.columns if f'{year}-' in col and f'-{week_number:02d}' in col]
        
        # Verificar si hay exactamente una coincidencia
        if len(matching_columns) == 1:
            matching_column = matching_columns[0]
            df_base.loc[df_base['Cod_Actual_1'] == product_code, matching_column] = row['Cantidad']
        elif len(matching_columns) > 1:
            # Si hay más de una coincidencia, mostrar un mensaje de advertencia
            print(f"Advertencia: Múltiples coincidencias para el patrón '{column_name_pattern}' en las columnas: {matching_columns}")
        else:
            # Si no se encuentra ninguna coincidencia
            print(f"No se encontró ninguna columna que coincida con el patrón '{column_name_pattern}'")


    # %%


    df_base['Faltante AP'] = 0

    # %%
    df_base = df_base.fillna(0)

    # %%
    meses_ingles_español = {
        "jan": "ene",
        "feb": "feb",
        "mar": "mar",
        "apr": "abr",
        "may": "may",
        "jun": "jun",
        "jul": "jul",
        "aug": "ago",
        "sep": "sep",
        "oct": "oct",
        "nov": "nov",
        "dec": "dic"
    }
    def obtener_mes_español(mes):
        mes_español = meses_ingles_español.get(mes)
        if mes_español:
            return mes_español.lower()
        else:
            return None

    # %%
    year_columns = [col for col in df_base.columns if col.split('-')[0].isdigit() and 'POS-STOCK' not in col]

    nueva_columna = f'POS-STOCK-{year_columns[0]}'
    df_base[nueva_columna] = df_base.apply(lambda row:  row['Stock CD'] + row['Stock Pañol'], axis=1)


    # %%
    for col in year_columns[1:]:
        column_name = f'POS-STOCK-{col}'
        
        last_column_name = df_base.columns[-1]
        year_month = last_column_name[-11:]
        
        mes = col[5:8]
        año = col[2:4]

    

        mes_español = obtener_mes_español(mes)
        if mes_español is None:
            print(f"Could not find Spanish equivalent for month: {mes}")
            continue

        columna_fc = f'FC {mes_español}-{año}'
        columna_tr = year_month

        



        calculo_columna = np.where((df_base[last_column_name] + df_base[columna_tr] - df_base[columna_fc]) < 0, 0, df_base[last_column_name] + df_base[columna_tr] - df_base[columna_fc])
        
        df_base[column_name] = calculo_columna
        print(column_name)



    # %%
    df_base_aux = df_base



    pos_columns = [col for col in df_base_aux.columns if 'POS-STOCK' in col]

    pos_columns[0][15:18]
    pos_columns[0][12:14]
    mes = pos_columns[0][15:18]
    año = pos_columns[0][12:14]

    mes_español = obtener_mes_español(mes)
    if mes_español is None:
        print(f"Could not find Spanish equivalent for month: {mes}")
        



    columna_fc = f'FC {mes_español}-{año}'
    df_base_aux[f'COBERTURA-{pos_columns[0][10:]}']= (df_base_aux[f'POS-STOCK-{pos_columns[0][10:]}']/((df_base_aux[columna_fc]/2)))
    df_base_aux[f'COBERTURA-{pos_columns[0][10:]}'].replace([np.inf, -np.inf, np.nan], '-', inplace=True)
    mes = pos_columns[1][15:18]
    año = pos_columns[1][12:14]

    mes_español = obtener_mes_español(mes)
    if mes_español is None:
        print(f"Could not find Spanish equivalent for month: {mes}")
        



    columna_fc = f'FC {mes_español}-{año}'
    df_base_aux[f'COBERTURA-{pos_columns[1][10:]}']= (df_base_aux[f'POS-STOCK-{pos_columns[1][10:]}']/df_base_aux[columna_fc])
    df_base_aux[f'COBERTURA-{pos_columns[1][10:]}'].replace([np.inf, -np.inf, np.nan], '-', inplace=True)

    for col in pos_columns[2:]:
        column_name = f'COBERTURA-{col[10:]}'
        
    
        mes = col[15:18]
        año = col[12:14]

        mes_español = obtener_mes_español(mes)
        if mes_español is None:
            print(f"Could not find Spanish equivalent for month: {mes}")
            



        columna_fc = f'FC {mes_español}-{año}'






        df_base_aux[column_name]= (df_base_aux[f'POS-STOCK-{column_name[10:]}']/df_base_aux[columna_fc])

        df_base_aux[column_name].replace([np.inf, -np.inf, np.nan], '-', inplace=True)
    # Set display options to show all columns and rows without truncation

    # Display the DataFrame without column truncation



    # %%
    df_base_aux['transito'] = df_base[year_columns].sum(axis=1)
        

    df_base_aux['pos_stock'] = df_base_aux['Stock CD'] + df_base_aux['Stock Pañol'] + df_base_aux['transito']

    # %%
    cob_columns = [col for col in df_base_aux.columns if 'COBERTURA' in col]
    # for c in cob_columns:
    #     print(c[10:])

    for col in cob_columns:
        nombre_columna = f'CUMPLIMIENTO-{col[10:]}'

        def calculate_value(row):
        
            cobertura = row[col]
            pos_stock_semanal = row[f'POS-STOCK-{col[10:]}']
            
            if cobertura == '-':
                return 1
            elif cobertura > 1:
                return 1
            elif cobertura < 0:
                return 0
            else:       
                return cobertura
            

            #cobertura es '-', y pos stock de esa semana es > 0 , 1
            #cobertura es '-', y pos stock de esa semana es 0 , 0

        # Apply the function row-wise using apply() and axis=1
        df_base_aux[nombre_columna] = df_base_aux.apply(calculate_value, axis=1)


    # %%
    cump_cols = [col for col in df_base_aux.columns if 'CUMPLIMIENTO' in col]

    # %%


    # %%
    for col in cump_cols:
        nombre_columna = f'NNSS_P - {col[13:]}'
        mes = col[18:21]
        año = col[15:17]

        mes_español = obtener_mes_español(mes)
        if mes_español is None:
            print(f"Could not find Spanish equivalent for month: {mes}")
            



        columna_fc = f'FC {mes_español}-{año}'

        df_base_aux[nombre_columna] = df_base[col] * df_base_aux[columna_fc]

        


    # %%
    ns_cols = [col for col in df_base_aux.columns if 'NNSS_P' in col]

    # %%
    for col in ns_cols:

        mes = col[14:17]
        año = col[11:13]
        #print(nombre_columna)

        

        mes_español = obtener_mes_español(mes)
        if mes_español is None:
            print(f"Could not find Spanish equivalent for month: {mes}")


        nombre_columna = f'forecast - {col[9:]}'
        
        
        columna_fc = f'FC {mes_español}-{año}'

        

        df_base_aux[nombre_columna] = df_base_aux[columna_fc]


            



        # columna_fc = f'FC {mes_español}-{año}'

        # df_base_aux[nombre_columna] = df_base[col] * df_base_aux[columna_fc]

        


    # %%
    cump_cols = [col for col in df_base_aux.columns if 'CUMPLIMIENTO' in col]

    df_base_aux['NNSS_Promedio'] = df_base_aux[cump_cols[:20]].mean(axis=1)
    df_base_aux['NNSS_Promedio_Aereo'] = df_base_aux[cump_cols[:12]].mean(axis=1)


    # %%
    df_base_aux.drop_duplicates(subset='Cod_Actual_1',inplace=True)
    df_base_aux.to_excel(f'C:/Users/{usuario}/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras KPI-Reportes/Disponibilidad Futura/2024/AXS/bases_python/Base_Final.xlsx')

    # %%
    sub_df = df_base_aux.filter(regex='^Cod_Actual_1$|^NNSS_P - ')
    sub_df_2 = df_base_aux.filter(regex = '^Cod_Actual_1$|^forecast - ')

    # %%


    # %%
    sub_df.columns

    # %%


    #declarar id
    id_vars = ['Cod_Actual_1']



    # Luego, usamos melt para transformar el DataFrame
    df_transformado = pd.melt(sub_df, id_vars=id_vars, var_name='NNSS - AÑO-MES-SEM', value_name='Cumplimiento')

    df_transformado_2 = pd.melt(sub_df_2, id_vars=id_vars, var_name='FC SEM', value_name='Forecast')


    # Puedes resetear los índices si lo deseas
    df_transformado.reset_index(drop=True, inplace=True)
    #f_transformado_2.reset_index(drop=True, inplace=True)





    # Ahora df_transformado contiene el DataFrame transformado como lo necesitas


    # %%


    # %%
    df_transformado_2['FC SEM'] = df_transformado_2['FC SEM'].str[11:]

    # %%
    df_transformado_2

    # %%
    df_transformado_2['ID'] = df_transformado_2['Cod_Actual_1'] + df_transformado_2['FC SEM']

    # %%
    df_transformado_2

    # %%
    df_transformado

    # %%
    df_transformado['ID_AUX'] = df_transformado['NNSS - AÑO-MES-SEM'].str[9:]


    # %%
    df_transformado

    # %%
    df_transformado['ID'] = df_transformado['Cod_Actual_1'] + df_transformado['ID_AUX']


    # %%
    df_transformado.nunique()

    # %%
    df_transformado = df_transformado.merge(df_transformado_2, left_on='ID',right_on='ID', how='left')

    # %%
    rename_cols = {'Cod_Actual_1_x':'Cod_Actual_1'}
    df_transformado.drop('Cod_Actual_1_y', inplace = True, axis=1)
    df_transformado.rename(columns=rename_cols, inplace = True)

    # %%
    reducir_cols = ['Cod_Actual_1','NNSS - AÑO-MES-SEM','Cumplimiento','Forecast']

    # %%
    df_transformado = df_transformado[reducir_cols]

    # %%
    df_transformado.to_csv(f'C:/Users/{usuario}/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras KPI-Reportes/Disponibilidad Futura/2024/AXS/bases_python/base_pbi.csv')
    print(f"🎊Proceso completo, archivo guardado en --> Planificación y Compras Chile - Documentos/Planificación y Compras KPI-Reportes/Disponibilidad Futura/2024/AXS")


if __name__ == '__main__':
    main()



