def main():    # %%
    #IMPORTACION DE LIBRERIAS
    import pandas as pd
    import datetime
    import os
    import numpy as np
    import getpass

    import warnings
    warnings.simplefilter(action="ignore", category=pd.errors.SettingWithCopyWarning)
    warnings.simplefilter(action="ignore", category=pd.errors.DtypeWarning)
    # hoy = datetime.datetime.today() dejar esta linea cuadno se haga el calculo real
    #hoy = datetime.datetime.today()
    #hoy = datetime.date(2024,10,17)
    #LECTURA DE DFS
    from pathlib import Path
    usuario = getpass.getuser()

    # %%
    import tkinter as tk
    from tkinter import ttk
    from tkcalendar import DateEntry
    import datetime

    # Variable global para almacenar la fecha seleccionada
    # fecha_seleccionada = None

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




    print(f'📂Ubicación base maestros: {ruta_repo}')
    #ruta = os.path.join(ruta_repo + '/DDP.csv')
    ruta_sugg = ruta_repo.joinpath('Suggested_Purchase.csv')

    # Leer el archivo CSV en un DataFrame
    df_ddp_1 = pd.read_csv(ruta_sugg)


    columnas= ['Nro_pieza_fabricante_1',	'Cod_Actual_1']
    ruta_cod = ruta_repo.joinpath('COD_ACTUAL.csv')

    # Leer el archivo CSV en un DataFrame
    cadena_de_remplazo = pd.read_csv(ruta_cod)
    cadena_de_remplazo = cadena_de_remplazo[columnas]


    # %%
    #MARA
    #columnas_mara = ['Material_R3','Part_number','Material_dsc','Modelo','Familia', 'Subfamilia', 'Categoría', 'Subcatgería','Sector_dsc', 'Material']
    ruta_maestro = f"C:/Users/{usuario}/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras Maestros/{hoy.year}/{hoy.year}-{str(hoy.month).zfill(2)}"


    # lista_maestro= os.listdir(ruta_maestro)
    # for i in lista_maestro:
    #     if 'MARA' in i and 'R3' in i:
    #         carpeta_mara = ruta_maestro + '/' + i
    # df_mara = excel_to_dataframe(carpeta_mara,'Sheet1')

    ruta_mara = ruta_repo.joinpath('MARA_R3.csv')

    # Leer el archivo CSV en un DataFrame
    df_mara = pd.read_csv(ruta_mara)
    #df_mara = df_mara[columnas_mara]


    # %%

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

    # ruta_obs = "C:/Users/lravlic/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras Maestros/2024/2024-05/new_obso_repuestos_cl_inchcape_202404.xlsx"
    # df_obs_1 = pd.read_excel(ruta_obs, sheet_name='Base Obs Cierre Abr-24', usecols=columnas)

    # %%
    df_obs_1

    # %%
    df_obs_1.columns.to_list()

    # %%
    #FC
    ruta_fc = f"C:/Users/{usuario}/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras OEM/Demanda y New Model Parts/Demanda/Demanda Mainstream/Forecast Colaborado/{(hoy).year}"
    lista_fc = os.listdir(ruta_fc)
    for i in lista_fc:

        if i[0:4] == str(hoy.year) and i[5:7] == str((hoy).month-1).zfill(2) :


            archivos_fc = os.listdir(ruta_fc + '/' + i )

            for j in archivos_fc:

                if 'OEM' in j:
                    archivo = ruta_fc + '/' + i + '/' + j
                    print(f'📂Achivo de FC usado:\n{j}')
                    df_fc = pd.read_excel(archivo,  sheet_name='MOS Forecast Data', header=3)
    #df_fc = pd.read_excel("C:/Users/lravlic/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras OEM/Demanda/Forecast Inbound/2024/2024-04 Abril/04.2024 S&OP Demanda Sin Restricciones OEM_Inbound.xlsx", sheet_name='Inbound', header=4)



    # #LT
    ruta_lt = f"C:/Users/{usuario}/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras KPI-Reportes/Disponibilidad Futura/2024/OEM/Bases Python/LT Actuales Mar-24.xlsx"
    df_lt = pd.read_excel(ruta_lt, header=1)
    # #STOCK

    # %%
    df_fc

    # %%
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


    # %%
    df_fc.rename(columns= {'FC sept-24': 'FC sep-24', 'FC sept-253':'FC sep-253'}, inplace=True)

    # %%
    df_fc.columns.to_list()

    # %%
    #Respaldo DDP para no cargar de nuevo el df
    df_ddp = df_ddp_1
    df_obs = df_obs_1

    # %%
    df_obs_1

    # %%
    df_obs_1 = df_obs_1.rename(columns={'Último Eslabón': 'Ultimo Eslabon'}, inplace = True)

    # %% [markdown]
    # LECTURA CAD REMPLAZO

    # %%
    df_ddp

    # %%
    df_ddp = df_ddp.merge(cadena_de_remplazo, left_on='Material', right_on='Nro_pieza_fabricante_1', how='left')
    df_ddp['Cod_Actual_1'] = df_ddp['Cod_Actual_1'].fillna(df_ddp['Material'])
    df_ddp = df_ddp.drop('Nro_pieza_fabricante_1', axis=1)

    # %%
    df_ddp

    # %%
    df_ddp.rename(columns={'Precio ':'Precio'}, inplace= True)

    # %%
    df_ddp['Precio'] = ""
    df_ddp['Moneda'] = ""


    # %%
    #traer el valor desde el material r3 y los casos que no crucen hacer lo mismo con cod_actual
    ddp_precio_moneda = df_ddp[['Material','Precio','Moneda']]

    # %%
    #aplicar lo mismo
    #para precio, moneda, origen, proveedor regular, costo, leadtime
    #ddp_origen = df_ddp.groupby(['Cod_Actual_1'])['Origen'].first()
    ddp_origen = df_ddp[['Material','Origen']]

    # %%
    ddp_filtro_origen = df_ddp.groupby('Cod_Actual_1').agg({'Marca':'first', 'Origen':'first'})

    # %%
    df_ddp.rename(columns={'Total Segmentation':'Segmentacion'}, inplace=True)
    df_ddp.rename(columns={'Apertura de parque':'Apertura Parque'}, inplace=True)

    # %%
    df_ddp

    # %%
    segmentacion = ['AA','AB','AC','BA','BB','BC','CA','CB','CC']
    ddp_segmentacion = df_ddp[df_ddp['Segmentacion'].isin(segmentacion)][['Cod_Actual_1','Segmentacion']].reset_index()

    #campo parque puede sustituir el campo apertura parque en el "o"
    ddp_estrategico = df_ddp[~df_ddp['Segmentacion'].isin(segmentacion) & ((df_ddp['Estratégico'] == 1) & ((df_ddp['Apertura Parque'] == 'Vigente') | (df_ddp['Apertura Parque'] == 'Nuevo')))][['Cod_Actual_1','Segmentacion']].reset_index()
    #aplicar logica anterior

    df_codigo = pd.concat([ddp_estrategico,ddp_segmentacion],axis=0).reset_index(drop=True)
    df_codigo = df_codigo.drop('index', axis=1).reset_index(drop=True)
    df_codigo = df_codigo.reset_index(drop=True)
    df_codigo.drop_duplicates(inplace = True)

    # %%
    df_codigo.sort_values(by='Segmentacion', inplace=True)

    # %%
    df_codigo

    # %%
    df_codigo.drop_duplicates(subset='Cod_Actual_1', inplace=True)

    # %%
    df_mara.drop_duplicates(subset='Material_R3', inplace=True)

    # %%
    df_base = pd.merge(df_codigo, df_mara, left_on = 'Cod_Actual_1', right_on='Material_R3', how='left')
    df_base['Part_number'] = df_base['Part_number'].str.replace(r'\[\#\]', '', regex=True)

    # %%
    #hacerlo a traves de la logica anterior
    #df_ddp_marca_origen = df_ddp[df_ddp['En dispo']==1].groupby('Cod_Actual_1').agg({'Marca': 'first', 'Origen': 'first'}).reset_index()
    df_ddp_marca_origen = df_ddp[['Material','Cod_Actual_1','Marca','Origen']]
    df_ddp_marca_origen.drop_duplicates(subset=['Material'],keep='first', inplace=True)
    df_base = df_base.merge(df_ddp_marca_origen[['Material','Marca','Origen']], left_on='Material_R3', right_on='Material', how= 'left')
    df_ddp_marca_origen.drop_duplicates(subset=['Cod_Actual_1'], keep='first', inplace=True)
    df_base = df_base.merge(df_ddp_marca_origen[['Cod_Actual_1','Marca','Origen']], left_on='Cod_Actual_1', right_on='Cod_Actual_1', how='left')

    df_base['Marca_x'] = df_base['Marca_x'].fillna(df_base['Marca_y'])
    df_base['Origen_x'] = df_base['Origen_x'].fillna(df_base['Origen_y'])
    df_base.drop(['Marca_y','Origen_y'], inplace = True, axis=1)
    df_base = df_base.rename(columns = {'Marca_x':'Marca','Origen_x':'Origen'})

    # %%
    df_base['Origen'][df_base['Marca'].isin(['Jac', 'Great Wall', 'Changan'])].value_counts()

    # %%
    df_fc = df_fc.merge(cadena_de_remplazo, left_on='Último Eslabón', right_on='Nro_pieza_fabricante_1', how ='left')


    # %%
    df_fc['Cod_Actual_1'] = df_fc['Cod_Actual_1'].fillna(df_fc['Último Eslabón'])

    # %%
    df_fc_prom = df_fc

    # %%
    df_fc_prom

    # %%
    #df_base['Faltante AP'] = 0

    # %%
    columnas_prom = [col for col in df_fc_prom.columns if 'FC' in col and 'Prom' not in col][:10]
    df_fc_prom['Promedio FC'] = df_fc_prom[columnas_prom].mean(axis=1)

    # %%
    df_fc_prom.Marca.value_counts()

    # %%
    columnas_seleccionadas = ['Cod_Actual_1'] + [col for col in df_fc_prom.columns if 'FC' in col and 'Prom' not in col][:10]

    nuevo_df_fc_prom = df_fc_prom[columnas_seleccionadas].copy()

    # %%
    nuevo_df_fc_prom

    # %%
    nuevo_df_fc_prom = nuevo_df_fc_prom.groupby('Cod_Actual_1').sum()/4.33

    # %%
    nuevo_df_fc_prom = nuevo_df_fc_prom.reset_index()

    # %%
    # Itera sobre las columnas del DataFrame
    nuevo_df_fc_prom.columns = [col[:-1] if col != "Cod_Actual_1" else col for col in nuevo_df_fc_prom.columns]


    # %%
    nuevo_df_fc_prom

    # %%
    df_fc_venta = df_fc
    columnas_venta = [col for col in df_fc_venta.columns if 'Vta R' in col][-3:]
    df_fc_venta['Promedio Venta'] = df_fc_venta[columnas_venta].mean(axis=1)

    # %%
    df_fc_venta = df_fc_venta.groupby(['Cod_Actual_1'])['Promedio Venta'].sum().reset_index()

    # %%
    df_fc = df_fc[['Cod_Actual_1', 'Segmentación Inchcape']].sort_values(by='Segmentación Inchcape')
    df_fc = df_fc.groupby('Cod_Actual_1').first().reset_index()

    # %%
    df_fc

    # %%
    df_ddp.rename(columns={'Plan Mantención':'Plan mantención'}, inplace=True)

    # %%
    df_plan_mantencion = df_ddp[['Cod_Actual_1', 'Plan mantención']].sort_values(by=['Cod_Actual_1','Plan mantención'])

    # %%
    df_plan_mantencion = df_plan_mantencion.groupby('Cod_Actual_1').max('Plan mantención').reset_index()

    # %%
    df_estrategicos  = df_ddp[['Cod_Actual_1', 'Estratégico']].sort_values(by=['Cod_Actual_1','Estratégico'])

    # %%
    df_estrategicos = df_estrategicos.groupby('Cod_Actual_1').max('Estratégico').reset_index()

    # %%
    #df_base = df_base.drop('Material_R3', axis=1)

    # %%
    #df_ddp.drop_duplicates(subset='Cod_Actual_1', inplace=True)

    # %%
    #hacerlo a traves de la logica anterior
    #df_ddp_marca_origen = df_ddp[df_ddp['En dispo']==1].groupby('Cod_Actual_1').agg({'Marca': 'first', 'Origen': 'first'}).reset_index()
    df_ddp_moq_1 = df_ddp[['Material','Cod_Actual_1','MOQ']]
    df_ddp_moq_1.drop_duplicates(subset=['Material'],keep='first', inplace=True)
    df_base = df_base.merge(df_ddp_moq_1[['Material','MOQ']], left_on='Material_R3', right_on='Material', how= 'left')
    df_ddp_moq = df_ddp[['Cod_Actual_1','MOQ']].sort_values(by='MOQ',ascending=False)
    df_ddp_moq['Cod_Actual_1'].value_counts()
    df_ddp_moq.drop_duplicates(subset=['Cod_Actual_1'], keep='first', inplace=True)
    df_base = df_base.merge(df_ddp_moq[['Cod_Actual_1','MOQ']], left_on='Cod_Actual_1', right_on='Cod_Actual_1', how='left')
    df_base['MOQ_x'] = df_base['MOQ_x'].fillna(df_base['MOQ_y'])
    df_base.drop(['MOQ_y','Material_x','Material_y'], inplace = True, axis=1)
    df_base = df_base.rename(columns = {'MOQ_x':'MOQ'})

    #df_base = df_base.merge(df_ddp[['Cod_Actual_1','Material Proveedor','Input DDA','MOQ']])

    # %%
    df_base = df_base.merge(df_fc, left_on='Cod_Actual_1', right_on = 'Cod_Actual_1', how='left')

    # %%
    df_base['Segmentación Inchcape'] = df_base['Segmentación Inchcape'].fillna('OO')

    # %%
    df_base['Segm. Planf']  = df_base['Segmentación Inchcape'].apply(lambda x: 1 if x in ['AA', 'AB', 'AC','BA','BB','BC','CA','CB','CC'] else 0)

    # %%
    df_base = df_base.merge(df_fc_venta, left_on='Cod_Actual_1', right_on='Cod_Actual_1', how='left')

    # %%
    df_base = df_base.merge(df_plan_mantencion, left_on='Cod_Actual_1', right_on='Cod_Actual_1', how='left')

    # %%
    df_input = df_ddp[['Cod_Actual_1', 'Input DDA']].sort_values(by=['Cod_Actual_1','Input DDA'])

    # %%
    df_input = df_input.groupby('Cod_Actual_1').max('Input DDA').reset_index()

    # %%
    df_base = df_base.merge(df_input, left_on='Cod_Actual_1', right_on='Cod_Actual_1', how='left')

    # %%
    df_base = df_base.merge(df_estrategicos, left_on='Cod_Actual_1', right_on='Cod_Actual_1', how='left')

    # %%
    df_base = df_base.merge(nuevo_df_fc_prom, left_on='Cod_Actual_1',right_on='Cod_Actual_1', how='left')

    # %%
    columnas_fc = [col for col in df_base.columns if 'FC' in col][:3]

    # Crear la nueva columna 'fc promedio' que contiene el promedio de las primeras tres columnas
    df_base['fc promedio'] = df_base[columnas_fc].mean(axis=1)*4.33

    # %%
    #df_base['Cobertura de stock'] = df.apply(lambda row: 0 if row['AI5'] == 0 else row['AI5']/row['R5'] if pd.notnull(row['AI5']) and pd.notnull(row['R5']) else 12, axis=1)

    # %%
    #material y luego codigo actual
    #df_ddp_costo = df_ddp[['Cod_Actual_1','Costo UN CLP']].groupby('Cod_Actual_1').max('Costo UN CLP').reset_index()

    # %%
    #material y luego codigo actual

    #df_ddp_descont = df_ddp[['Cod_Actual_1','Mateiales descontinuados']].groupby('Cod_Actual_1').max('Mateiales descontinuados').reset_index()

    # %%
    #df_ddp_costo['Cod_Actual_1'].nunique() == df_ddp_costo['Cod_Actual_1'].count()

    # %%
    #df_base = df_base.merge(df_ddp_costo, left_on='Cod_Actual_1',right_on='Cod_Actual_1', how='left')

    # %%
    #df_base = df_base.merge(df_ddp_descont, left_on='Cod_Actual_1',right_on='Cod_Actual_1', how='left')

    # %%
    df_base['Marca/Origen'] = df_base['Marca'] + df_base['Origen']

    # %%
    df_ddp['Parque'].fillna(0, inplace=True)
    idx_max_parque = df_ddp.groupby('Cod_Actual_1')['Parque'].idxmax()

    # %%


    #Vigente : 1
    #No Vigente: 0
    #Nuevo:2



    # Seleccionar las filas correspondientes a los índices encontrados
    df_ddp_parque = df_ddp.loc[idx_max_parque]

    # Restablecer los índices si es necesario
    df_ddp_parque.reset_index(drop=True, inplace=True)

    df_ddp_parque = df_ddp_parque[['Cod_Actual_1','Parque','Apertura Parque']]

    # %%
    df_base = df_base.merge(df_ddp_parque, left_on='Cod_Actual_1',right_on='Cod_Actual_1', how='left')

    # %%
    df_ddp

    # %%
    df_ddp.rename(columns={'Costo CPP':'Costo UN CLP'}, inplace=True)

    # %%
    df_ddp_costo_1 = df_ddp[['Material','Costo UN CLP']]
    df_ddp_costo_1.drop_duplicates(subset=['Material'], inplace=True)
    df_base = df_base.merge(df_ddp_costo_1[['Material','Costo UN CLP']], left_on='Material_R3', right_on='Material', how= 'left')
    df_ddp_costo = df_ddp[['Cod_Actual_1','Costo UN CLP']].sort_values(by='Costo UN CLP', ascending=False)
    df_ddp_costo.drop_duplicates(subset=['Cod_Actual_1'], keep='first', inplace=True)
    df_base = df_base.merge(df_ddp_costo, left_on='Cod_Actual_1', right_on='Cod_Actual_1', how='left')
    df_base['Costo UN CLP_x'] = df_base['Costo UN CLP_x'].fillna(df_base['Costo UN CLP_y'])
    df_base.drop(['Costo UN CLP_y'], inplace = True, axis=1)
    df_base = df_base.rename(columns = {'Costo UN CLP_x':'Costo CPP'})
    df_base.drop(['Material'], inplace = True, axis=1)

    # %%
    df_ddp

    # %%
    df_ddp_desc_1 = df_ddp[['Material','Materiales Descontinuados']]
    df_ddp_desc_1.drop_duplicates(subset=['Material'], inplace=True)
    df_base = df_base.merge(df_ddp_desc_1[['Material','Materiales Descontinuados']], left_on='Material_R3', right_on='Material', how= 'left')
    # df_ddp_desc = df_ddp[['Cod_Actual_1','Materiales Descontinuados']].sort_values(by='Materiales Descontinuados', ascending=False)
    # df_ddp_desc.drop_duplicates(subset=['Cod_Actual_1'], keep='first', inplace=True)
    # df_base = df_base.merge(df_ddp_desc, left_on='Cod_Actual_1', right_on='Cod_Actual_1', how='left')
    df_base['Materiales Descontinuados'] = df_base['Materiales Descontinuados'].fillna(0)
    df_base = df_base.rename(columns = {'Materiales Descontinuados':'Materiales Descontinuados'})
    df_base.drop(['Material'], inplace = True, axis=1)





    # %%
    columnas_fc = df_base.filter(like='FC')

    # Sumar las columnas
    suma_fc = columnas_fc.sum()

    # Mostrar el resultado


    # %%
    df_lt = df_lt[['Marca.1', 'Origen.1','Unnamed: 6',
        'Marca&Origen.1', 'Proveedor', 'LT', 'Sem LT']]
    columnas = {'Marca.1':'Marca', 'Origen.1':'Origen',
        'Marca&Origen.1':'Marca&Origen','Unnamed: 6':'Cod Proveedor'}

    df_lt.rename(columns=columnas, inplace=True)
    df_lt['Marca&Origen'].nunique() == df_lt['Marca&Origen'].count()

    # %%
    df_obs

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

    # %% [markdown]
    # NO SE ENCUENTRAN LOS ARCHIVOS NI DE LT NI OBSOLECENCIA

    # %%
    df_base = df_base.merge(din_obs_final,left_on='Cod_Actual_1', right_on='Ultimo Eslabon', how='left')

    # %%
    df_base.fillna(0, inplace=True)

    # %%
    df_base['Obsolescencia'] = np.where(df_base['obso_inchcape'].notna() & (df_base['obso_inchcape'] > 0), 1, 0)

    # %% [markdown]
    # diferencias con obs

    # %%
    df_base = df_base.merge(df_lt[['Marca&Origen', 'LT','Cod Proveedor']], left_on='Marca/Origen', right_on='Marca&Origen', how='left')

    # %%
    #hoy_datetime = datetime.datetime.combine(hoy, datetime.datetime.min.time())

    #hoy_datetime = datetime.date(2024,8,28)

    hoy_datetime = hoy

    # Adding the 'LT' values to hoy
    hoy_datetime = pd.to_datetime(hoy_datetime)

    # %%
    df_base['LT Semana'] = (hoy_datetime + pd.to_timedelta(df_base['LT'], unit='D')).dt.isocalendar().week

    # %%
    df_base['Mes'] = (hoy_datetime + pd.to_timedelta(df_base['LT'], unit='D')).dt.month

    # %% [markdown]
    # OBTENCION TRANSITO Y TUBO

    # %%
    df_stock['Centro'] = df_stock['Centro'].astype('str')

    # %%
    df_stock.Centro.value_counts()

    # %%
    # Convertir columnas 'Centro' y 'Almacén' a tipo string
    df_stock['Centro'] = df_stock['Centro'].astype(str)
    df_stock['Almacén'] = df_stock['Almacén'].astype(str)

    # Crear la columna 'Total' sumando las columnas especificadas
    df_stock['Total'] = df_stock['Libre utilización'] + df_stock['Trans./Trasl.'] + df_stock['En control calidad']

    # Eliminar las columnas no necesarias
    columns_to_drop = ['Libre utilización', 'Trans./Trasl.', 'En control calidad']
    df_stock = df_stock.drop(columns=columns_to_drop)

    # Filtrar las filas que cumplen con las condiciones especificadas
    condicion = (
        ((df_stock['Centro'] == '201') & (df_stock['Almacén'] == '1100')) |
        ((df_stock['Centro'] == '501') & (df_stock['Almacén'].isin(['1500', '1505'])))
    )
    df_stock_cd = df_stock[condicion]

    # Agrupar por 'Ult. Eslabon' y sumar la columna 'Total'
    df_stock_cd = df_stock_cd.groupby(['Ult. Eslabon']).agg({'Total': 'sum'}).reset_index()


    # %%
    df_base = df_base.merge(df_stock_cd, left_on='Cod_Actual_1', right_on='Ult. Eslabon', how='left')
    #df_base = df_base.merge(df_stock_entrante, left_on='Cod_Actual_1', right_on='Ult. Eslabon', how='left')


    # %%
    df_base['Stock_711'] = 0

    # %%
    df_base = df_base.fillna(0)

    # %% [markdown]
    #
    #

    # %%
    df_base['Cobertura Stock'] = np.where((df_base['fc promedio'] == 0),
                                            "FC 0",
                                        df_base['Total'] / df_base['fc promedio'])

    # Reemplazar inf con un valor específico (por ejemplo, 9999)
    df_base.replace([np.inf, -np.inf], 9999, inplace=True)

    # %%
    df_base.shape

    # %%
    df_base['Cobertura Stock'].value_counts()

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
    filtered_df[filtered_df['Material']=='PC2010041701']

    # %%
    df_base_2 =df_base

    # %%
    from datetime import timedelta

    # %%
    from datetime import timedelta, date
    import pandas as pd
    # Let's assume 'hoy' is a datetime.date object. For demonstration, I'll set it.
    # In your actual code, 'hoy' would be defined elsewhere.
    hoy = date.today() # This will be May 29, 2025



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


        df_base[column_name] = 0

    df_base.rename(columns={'2026-dec-01':'2026-jan-01'},inplace=True)
    # %%
    filtered_df = filtered_df[filtered_df['Cantidad']>0]

    # %%
    filtered_df['Fecha'] = pd.to_datetime(filtered_df['Fecha'])

    # %% [markdown]
    # ASIGNACION DE FECHAS V02 (05-09)
    from datetime import date, timedelta
    import pandas as pd # Assuming you are using pandas

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
            return iso_year, f"{iso_week:02d}"

    # Example DataFrame (replace with your actual filtered_df)
    # Let's create a sample DataFrame that simulates the issue
    # data = {'Fecha': pd.to_datetime(['2025-12-29', '2026-01-01', '2026-01-05', '2025-12-28'])}
    # filtered_df = pd.DataFrame(data)

    # Apply the function
    filtered_df['Año'], filtered_df['Semana'] = zip(*filtered_df['Fecha'].apply(lambda x: get_iso_week(x)))


    # Función para obtener el mes
    def get_month(year, week):
        return datetime.datetime.strptime(f'{year}-W{int(week)}-1', "%Y-W%W-%w").strftime('%B').lower()[:3]

    # Aplicar la función de mes
    filtered_df['Month'] = filtered_df.apply(lambda row: get_month(row['Año'], row['Semana']), axis=1)

    # %%
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


    # %%
    # # Supongamos que df_base es tu DataFrame base
    # # y filtered_df es el DataFrame con las ventas filtradas

    # # Primero, agrupamos las ventas por material, año, mes y semana
    # # grouped_sales = filtered_df.groupby(['Material', 'Año', 'Month', 'Semana'])['Cantidad'].sum().reset_index()
    # # grouped_sales['Año'] = grouped_sales['Año'].astype('str')
    # # grouped_sales['Semana'] = grouped_sales['Semana'].astype('str')
    # # # Luego, cruzamos los datos de ventas en df_base
    # # for index, row in grouped_sales.iterrows():
    # #     product_code = row['Material']
    # #     month = row['Month']
    # #     week_number = row['Semana']
    # #     year = row['Año']
    # #     column_name = f"{year}-{month}-{week_number}"
    # #     if column_name in df_base.columns:
    # #         df_base.loc[df_base['Cod_Actual_1'] == product_code, column_name] = row['Cantidad']

    # # Supongamos que df_base es tu DataFrame base
    # # y filtered_df es el DataFrame con las ventas filtradas

    # # Primero, agrupamos las ventas por material, año, mes y semana
    # grouped_sales = filtered_df.groupby(['Material', 'Año', 'Month', 'Semana'])['Cantidad'].sum().reset_index()
    # grouped_sales['Año'] = grouped_sales['Año'].astype('str')
    # grouped_sales['Semana'] = grouped_sales['Semana'].astype('int')  # Asegurarse de que Semana sea entero

    # # Luego, cruzamos los datos de ventas en df_base
    # for index, row in grouped_sales.iterrows():

    #     product_code = row['Material']
    #     week_number = row['Semana']
    #     year = row['Año']
    #     column_name_pattern = f"{year}-{week_number:02d}"

    #     # # Encuentra la columna en df_base que contenga el patrón
    #     matching_columns = [col for col in df_base.columns if f'{year}-' in col and f'-{str(week_number)}' in col]

    #     if matching_columns:
    #         print(matching_columns)
    #         matching_column = matching_columns[0]  # Asumimos que solo hay una coincidencia por patrón
    #         df_base.loc[df_base['Cod_Actual_1'] == product_code, matching_column] = row['Cantidad']







    # # Ahora df_base debe tener las ventas cruzadas en las columnas correspondientes


    # %%
    columnas = ['Ult. Eslabon','Ultimo Eslabon']
    df_base = df_base.drop(columns=columnas)
    df_base = df_base.rename({'Total_x':'Stock CD', 'Total_y':'Stock Entrante'})


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

    df_base['Qty Filial'] = 0

    nueva_columna = f'POS-STOCK-{year_columns[0]}'
    df_base[nueva_columna] = df_base.apply(lambda row: 0 if row['Total'] - row['Faltante AP'] - row['Qty Filial']<= 0 else row['Total'] - row['Faltante AP'] - row['Qty Filial'], axis=1)


    # %%
    df_base.columns.to_list()

    # %%
    df_base.dtypes

    # %%
    nueva_columna_2 = f'POS-STOCK-{year_columns[1]}'
    first_fc_column = df_base.filter(like='FC').columns[0]

    mes = year_columns[1][5:8]
    año = year_columns[1][2:4]

    mes_español = obtener_mes_español(mes)
    if mes_español is None:
        print(f"Could not find Spanish equivalent for month: {mes}")


    columna_fc = f'FC {mes_español}-{año}'


    df_base[nueva_columna_2] = np.where((df_base[nueva_columna] + df_base[year_columns[0]] - df_base[columna_fc]) < 0, 0, df_base[nueva_columna] + df_base[year_columns[0]] - df_base[columna_fc])

    # %%
    nueva_columna_3 = f'POS-STOCK-{year_columns[2]}'

    mes = year_columns[2][5:8]
    año = year_columns[2][2:4]

    mes_español = obtener_mes_español(mes)
    if mes_español is None:
        print(f"Could not find Spanish equivalent for month: {mes}")




    columna_fc = f'FC {mes_español}-{año}'


    df_base[nueva_columna_3] = np.where((df_base[nueva_columna_2] + df_base[year_columns[1]] + df_base['Stock_711'] - df_base[columna_fc]) < 0, 0, df_base[nueva_columna_2] + df_base[year_columns[1]] + df_base['Stock_711'] - df_base[columna_fc])


    # %%
    for col in year_columns[3:]:
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




    # %%
    #cobertura
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
    df_base_aux['transito'] = df_base[year_columns].sum(axis=1)


    df_base_aux['pos_stock'] = df_base_aux['Total'] + df_base_aux['Stock_711'] + df_base_aux['transito']
    cob_columns = [col for col in df_base_aux.columns if 'COBERTURA' in col]
    # for c in cob_columns:
    #     print(c[10:])

    import warnings
    from pandas.errors import PerformanceWarning

    warnings.simplefilter(action="ignore", category=PerformanceWarning)


    for col in cob_columns:
        nombre_columna = f'CUMPLIMIENTO-{col[10:]}'

        def calculate_value(row):
            segmentacion = row['Segmentacion']
            vta_prom = row['Promedio Venta']
            pos_stock = row['pos_stock']
            cobertura = row[col]
            pos_stock_semanal = row[f'POS-STOCK-{col[10:]}']


            if cobertura == '-':
                return 1
            elif segmentacion == 'AA':
                if cobertura > 1:
                    return 1
                elif cobertura < 0:
                    return 0
                else:
                    return cobertura

            elif vta_prom < 1 and pos_stock > 0:
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

    import pandas as pd

    # Set display options to show all columns and rows without truncation




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
    df_base_aux.head()

    # %%
    #df_base_aux.to_excel('C:/Users/lravlic/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras KPI-Reportes/Disponibilidad Futura/2024/OEM/Bases Python/Base Analisis.xlsx')

    df_base_aux = df_base_aux.rename(columns={'Total': 'Stock CD'})



    # %%
    columnas = ["Nro_material", "Fecha_creacion", "Tipo_material", "Grupo_articulo", "Grupo_art_desc", "Grupo_art_externo", "Sector", "Sector_dsc", "Jerarquia_producto", "Jquia_desc", "Material_R3", "Tamaño", "Nodo de Jerarquía", "Segmentación Inchcape", "Marca&Origen"]

    # %%
    df_base_aux.drop(columns=columnas, inplace= True)

    # %%
    df_base_aux

    # %%
    #df_base_aux.to_excel(f"C:/Users/{usuario}/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras KPI-Reportes/Disponibilidad Futura/2024/OEM/Bases Python/base_final.xlsx", index=False)

    # %%
    df_base_aux

    # %%
    sub_df = df_base_aux.filter(regex='^Cod_Actual_1$|^NNSS_P - ')
    sub_df_2 = df_base_aux.filter(regex = '^Cod_Actual_1$|^forecast - ')

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
    df_transformado_2['FC SEM'] = df_transformado_2['FC SEM'].str[11:]

    # %%
    df_transformado_2['ID'] = df_transformado_2['Cod_Actual_1'].astype('str') + df_transformado_2['FC SEM'].astype('str')

    # %%
    df_transformado_2

    # %%
    df_transformado

    # %%
    df_transformado['ID_AUX'] = df_transformado['NNSS - AÑO-MES-SEM'].str[9:]


    # %%
    df_transformado['ID'] = df_transformado['Cod_Actual_1'].astype('str') + df_transformado['ID_AUX'].astype(str)

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
    #df_transformado.to_csv(f'C:/Users/lravlic/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras KPI-Reportes/Disponibilidad Futura/2024/OEM/Bases Python/Base_PBI.csv')
    #df_transformado.to_csv(f'C:/Users/{usuario}/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras KPI-Reportes/Disponibilidad Futura/2024/OEM/Bases Python/base_pbi.csv')

    # %%
    df_mara.dropna(subset=['Material_R3'], inplace=True)

    # Assuming 'df_mara' is your DataFrame


    # %%
    # Eliminar duplicados basados en la columna 'Material_R3'
    df_mara.drop_duplicates(subset=['Material_R3'], inplace=True)

    # %%
    df_mara.dtypes

    # %%
    #df_mara.to_csv(f'C:/Users/{usuario}/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras KPI-Reportes/Disponibilidad Futura/2024/OEM/Bases Python/mara_tratada.csv')

    # %%


    # %%


    # %%


    # %%


    # %%


    # %%


    # %%


    # %%


    # %%


    # %%


    # %%


    # %%


    # %%


    # %%


    # %%


    # %%


    # %%



if __name__ == '__main__':
    main()
