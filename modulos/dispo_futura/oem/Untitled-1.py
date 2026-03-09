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


# %%
# Obtiene el nombre de usuario del sistema operativo actual
nombre_usuario = getpass.getuser()

# %%
# Obtiene la fecha actual
hoy = datetime.today()  # Ajuste de zona horaria para Chile (UTC-3)

# Obtiene el año actual
año = hoy.year

# Obtiene el año actual con dos dígitos
año_2 = año % 100

# Obtiene el número de mes actual
numero_mes = str(hoy.month).zfill(2)

# Obtiene la fecha actual en formato YYYY-MM
fecha_n = hoy.strftime('%Y-%m')

# Obtiene la fecha actual en formato MM.YYYY
numero_mes_punto = hoy.strftime('%m.%Y')

# Obtiene la fecha de hace un mes
n_date = hoy - relativedelta(months=1)

# Obtiene la fecha de hace dos meses
n_1_date = hoy - relativedelta(months=2)

# Obtiene la fecha de hace cinco meses
n_4_date = hoy - relativedelta(months=5)

# Obtiene la fecha de hace dos meses en formato YYYY-MM
n_1 = n_1_date.strftime("%Y-%m")

# Obtiene la fecha de hace cinco meses en formato YYYY-MM
n_4 = n_4_date.strftime("%Y-%m")

# Obtiene la fecha de hace dos meses en formato MM.YYYY
n1_punto = n_1_date.strftime("%m.%Y")

# Obtiene la fecha de hace cinco meses en formato MM.YYYY
n4_punto = n_4_date.strftime("%m.%Y")

# Imprime las variables
print("n_1:", n_1)
print("n_4:", n_4)
print(n1_punto)
print(n4_punto)



# %%
# Crea una cadena con la fecha del primer día del mes anterior
n_b_fc = n_date.strftime( '%Y-%m'+'-01')

# %%
n_b_fc

# %%
#Funcion que entrega el nombre correspondiente al numero del mes
#input: numero de mes
#output: nombre de dicho mes
def obtener_nombre_mes(numero_mes):
    meses = {
        "01": "Enero",
        "02": "Febrero",
        "03": "Marzo",
        "04": "Abril",
        "05": "Mayo",
        "06": "Junio",
        "07": "Julio",
        "08": "Agosto",
        "09": "Septiembre",
        "10": "Octubre",
        "11": "Noviembre",
        "12": "Diciembre"
    }
    return meses.get(numero_mes, "Mes no válido")

# se crean variables de meses relativos a los peridos de medicion n-1 y n-4
mes_n1 = (hoy - relativedelta(months=2)).month
mes_n4 = (hoy - relativedelta(months=5)).month
ciclo_n1 = (hoy - relativedelta(months=1)).month
ciclo_n4 = (hoy - relativedelta(months=4)).month


#se obtienen nombres de dichos meses
nombre_mes_n1 = obtener_nombre_mes(mes_n1)
nombre_mes_n4 = obtener_nombre_mes(mes_n4)
nombre_ciclo_n1 = obtener_nombre_mes(ciclo_n1)
nombre_ciclo_n4 = obtener_nombre_mes(ciclo_n4)
nombre_numero_mes = obtener_nombre_mes(numero_mes)


# Imprimir el nombre del mes anterior
print(nombre_mes_n1)
print(nombre_mes_n4)


# %%
#Se especifican las rutas de los archivos de forecast para los ciclos n-1 y n-4
ruta_n1 = f"C:/Users/{nombre_usuario}/Inchcape/Planificación y Compras Chile - Documentos/Bases Indicadores en CSV {str(hoy.year)}-{numero_mes}/FORECAST N-1 PREMIUM.csv"
ruta_n4 = f"C:/Users/{nombre_usuario}/Inchcape/Planificación y Compras Chile - Documentos/Bases Indicadores en CSV {str(hoy.year)}-{numero_mes}/FORECAST N-4 PREMIUM.csv"


# %%
# arch = pd.ExcelFile(ruta_n1)
# arch_2 = pd.ExcelFile(ruta_n4)

# %%
numero_mes

# %%

df_n1 = pd.read_csv(ruta_n1, encoding='latin')
df_n4 = pd.read_csv(ruta_n4, encoding='latin')

# %%
maestro = f"C:/Users/{nombre_usuario}/Inchcape/Planificación y Compras Chile - Documentos/Bases Indicadores en CSV {str(hoy.year)}-{numero_mes}"
dir_maestro = os.listdir(maestro)

for a in dir_maestro:

#     if str(hoy.year) in c_año:
#         c_carpeta = os.path.join(maestro, c_año)
#         c_mes = os.listdir(c_carpeta)
#         c_arch = os.path.join(c_carpeta, c_mes[-1])
#         archivos = os.listdir(c_arch)
#         print(archivos)

    if 'COD_ACTUAL_PREMIUM' in a:
        ruta_cad = os.path.join(maestro, a)
        cadena_de_remplazo = pd.read_csv(ruta_cad, usecols= ['Nro_pieza_fabricante_1',	'Cod_Actual_1', 'Legacy'])

cadena_de_remplazo.drop_duplicates(subset='Nro_pieza_fabricante_1', inplace=True)
cadena_de_remplazo = cadena_de_remplazo[cadena_de_remplazo['Legacy']=='Legacy BMW']
# %%
ventas = f"C:/Users/{nombre_usuario}/Inchcape/Planificación y Compras Chile - Documentos/Bases Indicadores en CSV {str(hoy.year)}-{numero_mes}"
dir_maestro = os.listdir(maestro)

for a in dir_maestro:
    print(a)
    

#     if str(hoy.year) in c_año:
#         c_carpeta = os.path.join(maestro, c_año)
#         c_mes = os.listdir(c_carpeta)
#         c_arch = os.path.join(c_carpeta, c_mes[-1])
#         archivos = os.listdir(c_arch)
#         print(archivos)

    if 'Sell In Premium' in a:
    #and str(numero_mes).zfill(2) in a and str(año) in a:

        ruta_ventas = os.path.join(ventas, a)
df_ventas= pd.read_excel(ruta_ventas, sheet_name='BV Sell In Premium', dtype={'Último Eslabón':'str','Material':'str'},header=2)        
#df_ventas_p= pd.read_excel(ruta_ventas, sheet_name='TD Venta Sell In', header=2, dtype={'Último Eslabón':'str'})
df_ventas_p = df_ventas[['Último Eslabón', 'Promedio de Venta 12M']]


# %%


# %%

# %%
df_n1.rename(columns={'Último Eslabón':'SKU ERP'}, inplace=True)
df_n4.rename(columns={'Último Eslabón':'SKU ERP'}, inplace=True)

# %%
df_n1['SKU ERP'] = df_n1['SKU ERP'].astype('str')
df_n1 = pd.merge(df_n1,cadena_de_remplazo, left_on='SKU ERP', right_on='Nro_pieza_fabricante_1', how='left')

# # #df_n1 = df_n1.drop(['Cod_Actual_1','Nro_pieza_fabricante_1','Cod_Actual_1_UE2','Nro_pieza_fabricante_1_UE2','Cod_Actual_1_Cod_Actual_1', 'Nro_pieza_fabricante_1_Cod_Actual_1'], axis=1)
# # #df_n4['Ult. Eslabón'] = df_n4['Ult. Eslabón'].astype('str')
#df_n4 = pd.merge(df_n4,cadena_de_remplazo, left_on='SKU ERP', right_on='Nro_pieza_fabricante_1', how='left')



# %%
df_n4['SKU ERP'] = df_n4['SKU ERP'].astype('str')
df_n4 = pd.merge(df_n4,cadena_de_remplazo, left_on='SKU ERP', right_on='Nro_pieza_fabricante_1', how='left')


# %%
df_n1['Cod_Actual_1'] = df_n1['Cod_Actual_1'].fillna(df_n1['SKU ERP'])

df_n4['Cod_Actual_1'] = df_n4['Cod_Actual_1'].fillna(df_n4['SKU ERP'])


# %%
df_n1 = df_n1.drop('Nro_pieza_fabricante_1', axis=1)

df_n4 = df_n4.drop('Nro_pieza_fabricante_1', axis=1)


# %%
df_n1 = df_n1.rename(columns = {'SKU ERP': 'UE_2'})
df_n4 = df_n4.rename(columns = {'SKU ERP': 'UE_2'})


# %%
df_n1['fc'] = df_n1['n1_colab']
df_n4['fc'] = df_n4['n4_colab']


# %%


df_n1['fc_anas'] = df_n1['n1_base']
df_n4['fc_anas'] = df_n4['n4_base']

df_n1['fc_std'] = df_n1['n1_std']
df_n4['fc_std'] = df_n4['n4_std']

# %%
df_n1 = df_n1.iloc[2:]
df_n4 = df_n4.iloc[2:]

        # %%
    # Corrige todas las columnas de una sola vez
# ==========================================
# PROCESAMIENTO PARA DATAFRAME: df_n1
# ==========================================

# 1. Columna: fc_anas
# Convertimos a string, quitamos espacios, cambiamos coma por punto
df_n1['fc_anas'] = df_n1['fc_anas'].astype(str).str.strip().str.replace(',', '.', regex=False)
# Convertimos a número (si falla algo, pondrá NaN en lugar de romper el código)
df_n1['fc_anas'] = pd.to_numeric(df_n1['fc_anas'], errors='coerce')

# 2. Columna: fc
df_n1['fc'] = df_n1['fc'].astype(str).str.strip().str.replace(',', '.', regex=False)
df_n1['fc'] = pd.to_numeric(df_n1['fc'], errors='coerce')

# 3. Columna: fc_std
df_n1['fc_std'] = df_n1['fc_std'].astype(str).str.strip().str.replace(',', '.', regex=False)
df_n1['fc_std'] = pd.to_numeric(df_n1['fc_std'], errors='coerce')


# ==========================================
# PROCESAMIENTO PARA DATAFRAME: df_n4
# ==========================================

# 1. Columna: fc_anas
df_n4['fc_anas'] = df_n4['fc_anas'].astype(str).str.strip().str.replace(',', '.', regex=False)
df_n4['fc_anas'] = pd.to_numeric(df_n4['fc_anas'], errors='coerce')

# 2. Columna: fc
df_n4['fc'] = df_n4['fc'].astype(str).str.strip().str.replace(',', '.', regex=False)
df_n4['fc'] = pd.to_numeric(df_n4['fc'], errors='coerce')

# 3. Columna: fc_std
df_n4['fc_std'] = df_n4['fc_std'].astype(str).str.strip().str.replace(',', '.', regex=False)
df_n4['fc_std'] = pd.to_numeric(df_n4['fc_std'], errors='coerce')

# %%
df_n1['ID2'] = df_n1['Cod_Actual_1'] 
df_n4['ID2'] = df_n4['Cod_Actual_1'] 

# %%
df_n1['Periodo'] = 'N-1'
df_n4['Periodo'] = 'N-4'

# %%
df_consolidado = pd.concat([df_n1,df_n4], ignore_index=True)



# %%





# %%
df_n1

# %%
cols = [ 'UE_2', 'ID2','fc_anas','fc','fc_std', 'Marca', 'Segmentación Inchcape','Familia','Descripción Material','Periodo','Costo Promedio Ponderado', 'Input','Legacy_x']


# %%
df_consolidado_red = df_consolidado[cols]


# %%
consolidado_seg = df_consolidado_red[['ID2', 'Segmentación Inchcape']]
consolidado_seg.sort_values(by='Segmentación Inchcape', inplace=True, ascending=False)

consolidado_input = df_consolidado[['ID2', 'Input']]
consolidado_seg.sort_values(by='Segmentación Inchcape', inplace=True, ascending=False)



# %%
df_consolidado_red = df_consolidado_red.groupby(['ID2', 'Periodo']).agg({'UE_2':'first', 'fc': 'sum', 'fc_anas':'sum', 'fc_std': 'sum', 'Costo Promedio Ponderado': 'max',  'Marca':'first','Familia':'first','Legacy_x':'first'}).reset_index()

# %%
df_consolidado_red

# %%
consolidado_seg.drop_duplicates(subset='ID2', inplace=True, keep='first')
consolidado_input.drop_duplicates(subset='ID2', inplace=True, keep='first')

# %%
df_consolidado_red.shape

# %%
df_consolidado_red = df_consolidado_red.merge(consolidado_seg, on='ID2', how='left')
df_consolidado_red = df_consolidado_red.merge(consolidado_input, on='ID2', how='left')



# %%
indice_deseado = 27 # El índice 57 corresponde a la columna número 58


# Accede a la columna por su índice
col_deseada = df_ventas.iloc[:, indice_deseado]

# %%
df_ventas['Venta'] = col_deseada

# %%
df_ventas['Venta'].sum()

# %%


# %%


# %%
red_columnas = ['Último Eslabón', 'Venta']
df_ventas = df_ventas[red_columnas]

# %%
df_ventas.shape

# %%
cadena_de_remplazo
print(df_ventas.columns)
# %%
# df_ventas = df_ventas.merge(cadena_de_remplazo, left_on='Material', right_on = 'Nro_pieza_fabricante_1' , how='left')


# # %%
# df_ventas['Cod_Actual_1'].fillna(df_ventas['Material'], inplace=True)
# df_ventas = df_ventas[['Cod_Actual_1','Venta']]
# df_ventas.rename(columns = {'Material':'Último Eslabón'},inplace=True)
df_ventas = df_ventas.groupby(['Último Eslabón']).agg({'Venta': sum}).reset_index()

# %%


# # %%
# df_ventas_p = df_ventas_p.merge(cadena_de_remplazo, left_on='Material', right_on = 'Nro_pieza_fabricante_1' , how='left')
# df_ventas_p['Cod_Actual_1'].fillna(df_ventas_p['Material'], inplace=True)
# df_ventas_p = df_ventas_p[['Cod_Actual_1','Promedio de Venta 12M']]
#df_ventas_p.rename(columns = {'Material':'Último Eslabón'},inplace=True)
df_ventas_p = df_ventas_p.groupby(['Último Eslabón']).agg({'Promedio de Venta 12M': sum}).reset_index()

# %%
consolidado_cruce_ventas = pd.merge(df_consolidado_red, df_ventas, left_on='ID2', right_on='Último Eslabón', how='left')
consolidado_cruce_ventas = consolidado_cruce_ventas.merge(df_ventas_p, left_on='ID2', right_on='Último Eslabón', how='left')

# %%
consolidado_cruce_ventas = consolidado_cruce_ventas.drop(['Último Eslabón_x','Último Eslabón_y'], axis=1)


# %%
consolidado_cruce_ventas

# %%
consolidado_cruce_ventas['Legacy_x'].value_counts()

# %%


# %%
# Actualiza 'Marca' para los registros Legacy BMW según el primer caracter de 'UE_2'
mask = consolidado_cruce_ventas['Legacy_x'] == 'Legacy BMW'
first_char = consolidado_cruce_ventas.loc[mask, 'UE_2'].astype(str).str.strip().str[0].str.upper()

mapping = {
    'B': 'BMW', 'R': 'BMW', 'I': 'BMW',
    'C': 'BMW Motorrad',
    'M': 'Mini',
    'Z': 'Nacional WBM', 'O': 'Nacional WBM'
}

consolidado_cruce_ventas.loc[mask, 'Marca'] = first_char.map(mapping).fillna(consolidado_cruce_ventas.loc[mask, 'Marca'])

# %%

# %%






import tkinter as tk
from tkinter import filedialog

root = tk.Tk()
root.withdraw()  # Oculta la ventana principal

carpeta_tubo = filedialog.askdirectory(title="Selecciona la carpeta 'Carpeta tubo'")
print("Carpeta seleccionada:", carpeta_tubo)

fecha_actual = datetime.today()
fecha_mes_anterior = fecha_actual - timedelta(days=fecha_actual.day)
fecha_formateada = fecha_mes_anterior.strftime('%Y-%m')

# %%
fecha_formateada

# %%


meses_espanol = {
    '01': "Enero",
    '02': "Febrero",
    '03': "Marzo",
    '04': "Abril",
    '05': "Mayo",
    '06': "Junio",
    '07': "Julio",
    '08': "Agosto",
    '09': "Septiembre",
    '10': "Octubre",
    '11': "Noviembre",
    '12': "Diciembre"
}

contador = 1


tubo = os.listdir(carpeta_tubo)
for i in tubo:
    print(i)
# for i in tubo:
#     if i[0:4] == fecha_formateada[0:4]:
#         print(i[0:4])
#         carp_año = i
    
# carpeta_tubo_2 = os.listdir(carpeta_tubo + '/' + carp_año)
# for i in carpeta_tubo_2:
#     if meses_espanol.get(fecha_formateada[5:7]) in i:
#         carp_mes = i
#         print(carp_mes)
# carpeta_premium = carpeta_tubo + '/' + carp_año + '/' + carp_mes

carpeta_bmw = carpeta_tubo + '/BMW'
carpeta_ditec = carpeta_tubo + '/Ditec'

carpeta_tubo_bmw = os.listdir(carpeta_bmw)
carpeta_tubo_ditec = os.listdir(carpeta_ditec)

lista_bmw = []
lista_ditec = []

for i in carpeta_tubo_bmw:
    if 'tock' in i:
            lista_bmw.append(i)
            

for i in carpeta_tubo_ditec:
    if 'tock' in i:
        lista_ditec.append(i)


# %%
len(lista_bmw)

# %%


        

# %%



# %%
fecha_formateada[5:7]

# %%
dfs = []
for i in range(0,len(lista_bmw)):

    df_a = pd.read_excel(carpeta_bmw + '/' + lista_bmw[i],sheet_name='Sheet1', dtype={'Material':'str'})
    df_b = pd.read_excel(carpeta_ditec + '/' + lista_ditec[i], sheet_name='Sheet1',dtype={'Material':'str'})
    df = pd.concat([df_a,df_b])
    dfs.append(df)


# %%




# %%


# %%
for df in dfs:
    df = df[['Material','Libre utilización','SEM']]
    print(df.columns.to_list())

for df in dfs:
    max_sem = df['SEM'].max()
    max_sem = max_sem.isocalendar()[1]
    df = pd.merge(df,cadena_de_remplazo, left_on='Material', right_on='Nro_pieza_fabricante_1', how='left')
    df['Cod_Actual_1'] = df['Cod_Actual_1'].fillna(df['Material'])
    df = df.drop('Nro_pieza_fabricante_1', axis=1)
    df_suma = df.groupby(['Cod_Actual_1']).agg({'Libre utilización': 'sum'}).reset_index()
    df_suma = df_suma.rename(columns={'Libre utilización': f'Stock SEM {max_sem}', 'Cod_Actual_1': f'Cod_Actual_1{max_sem}'})
    # df_suma.drop(columns = {'SEM'}, inplace=True)
    suffix = f'_{i}' if i == (len(tubo) - 2) else f'_{i}a'  # Añadir sufijo para la primera fusión
    consolidado_cruce_ventas = consolidado_cruce_ventas.merge(df_suma, left_on='ID2', right_on=f'Cod_Actual_1{max_sem}', how='left')


# %%
consolidado_cruce_ventas

# %%
consolidado_cruce_ventas = consolidado_cruce_ventas.fillna(0)
columns_to_drop = consolidado_cruce_ventas.filter(like='Material').columns
consolidado_cruce_ventas = consolidado_cruce_ventas.drop(columns=columns_to_drop)

# %%

semanas_habiles_por_mes = {
    '2023-10': 5,
    '2023-11': 4,
    '2023-12': 4,
    '2024-01': 5,
    '2024-02': 4,
    '2024-03': 4,
    '2024-04': 5,
    '2024-05': 5,
    '2024-06': 5,
    '2024-07': 4,
    '2024-08': 4,
    '2024-09': 4,
    '2024-10': 5,
    '2024-11': 4,
    '2024-12': 4,
    '2025-01': 5,
    '2025-02': 4,
    '2025-03': 4,
    '2025-04': 4,
    '2025-05': 4,
    '2025-06': 4,
    '2025-07': 4,
    '2025-08': 4,
    '2025-09': 4,
    '2025-10': 4,
    '2025-11': 4,
    '2025-12': 4
}

# %%
consolidado_cruce_ventas

# %%
consolidado_cruce_ventas['Venta'] = consolidado_cruce_ventas['Venta'].apply(lambda x: x if x >= 0 else 0)


consolidado_cruce_ventas['Promedio de Venta 12M'] = consolidado_cruce_ventas['Promedio de Venta 12M'].apply(lambda x: x if x >= 0 else 0)
#consolidado_cruce_ventas['Promedio de Venta 12M'] = consolidado_cruce_ventas['Promedio de Venta 12M']/4
consolidado_cruce_ventas['Promedio de Venta 12M'] = consolidado_cruce_ventas.apply(
    lambda row: row['Promedio de Venta 12M'] / len(lista_bmw),
    axis=1
)

# %%
for column in consolidado_cruce_ventas.columns:
    if 'Stock' in column:
        new_column_name = f'INSTOCK_{column[-2:]}'
        consolidado_cruce_ventas[new_column_name] = consolidado_cruce_ventas.apply(lambda row: 1 if row['Promedio de Venta 12M'] < row[column] else 0, axis=1)


# %%
consolidado_cruce_ventas['Total instock'] = consolidado_cruce_ventas.filter(like='INSTOCK').sum(axis=1)


# %%
consolidado_cruce_ventas[consolidado_cruce_ventas['ID2']=='B07129909320']

# %%

# %%
consolidado_cruce_ventas['MAPE_AJUSTADO'] = abs(consolidado_cruce_ventas['Venta'] - consolidado_cruce_ventas['fc'])

# %%

consolidado_cruce_ventas['ERP'] = consolidado_cruce_ventas.apply(lambda row: row['fc'] - row['Venta'] if row['fc'] > row['Venta'] else 0, axis=1)
consolidado_cruce_ventas['ERN'] = consolidado_cruce_ventas.apply(lambda row: row['Venta'] - row['fc'] if row['Venta'] > row['fc'] else 0, axis=1)
consolidado_cruce_ventas['MAPE_ANASTASIA'] = abs(consolidado_cruce_ventas['Venta'] - consolidado_cruce_ventas['fc_anas'])

# %%

consolidado_cruce_ventas['ERP_ANASTASIA'] = consolidado_cruce_ventas.apply(lambda row: row['fc_anas'] - row['Venta'] if row['fc_anas'] > row['Venta'] else 0, axis=1)
consolidado_cruce_ventas['ERN_ANASTASIA'] = consolidado_cruce_ventas.apply(lambda row: row['Venta'] - row['fc_anas'] if row['Venta'] > row['fc_anas'] else 0, axis=1)
consolidado_cruce_ventas['MAPE_ESTADISTICO'] = abs(consolidado_cruce_ventas['Venta'] - consolidado_cruce_ventas['fc_std'])
consolidado_cruce_ventas['ERP_ESTADISTICO'] = consolidado_cruce_ventas.apply(lambda row: row['fc_std'] - row['Venta'] if row['fc_std'] > row['Venta'] else 0, axis=1)
consolidado_cruce_ventas['ERN_ESTADISTICO'] = consolidado_cruce_ventas.apply(lambda row: row['Venta'] - row['fc_std'] if row['Venta'] > row['fc_std'] else 0, axis=1)


# %%

consolidado_cruce_ventas['WMAPE'] = consolidado_cruce_ventas['MAPE_AJUSTADO'] / consolidado_cruce_ventas['Venta']
consolidado_cruce_ventas['WMAPE.1'] = consolidado_cruce_ventas['MAPE_ANASTASIA'] / consolidado_cruce_ventas['Venta']


consolidado_cruce_ventas['WMAPE.2'] = consolidado_cruce_ventas['MAPE_ESTADISTICO'] / consolidado_cruce_ventas['Venta']

# %%

consolidado_cruce_ventas['Mes'] = hoy.strftime('%Y-%m-%d')

# %%

consolidado_cruce_ventas['Costo Promedio Ponderado'] = pd.to_numeric(consolidado_cruce_ventas['Costo Promedio Ponderado'].astype(str).str.replace(',', '.'), errors='coerce').fillna(0)
consolidado_cruce_ventas['Costo de Venta'] = consolidado_cruce_ventas['Costo Promedio Ponderado'] * consolidado_cruce_ventas['Venta']

# %%

consolidado_cruce_ventas['Forecast en Costo'] = consolidado_cruce_ventas['Costo Promedio Ponderado'] * consolidado_cruce_ventas['fc']
consolidado_cruce_ventas['MAPE Costo']  = abs(consolidado_cruce_ventas['Costo de Venta'] - consolidado_cruce_ventas['Forecast en Costo'])

# %%
consolidado_cruce_ventas['ERP Costo'] = consolidado_cruce_ventas.apply(lambda row: row['Forecast en Costo'] - row['Costo de Venta'] if row['Forecast en Costo'] > row['Costo de Venta'] else 0, axis=1)
consolidado_cruce_ventas['ERN Costo'] = consolidado_cruce_ventas.apply(lambda row: row['Costo de Venta'] - row['Forecast en Costo'] if row['Costo de Venta'] > row['Forecast en Costo'] else 0, axis=1)
consolidado_cruce_ventas['WMAPE Costo'] = consolidado_cruce_ventas['MAPE Costo'] / consolidado_cruce_ventas['Costo de Venta']

# %%





# %%
# %%
consolidado_cruce_ventas = consolidado_cruce_ventas[['ID2', 'Periodo','UE_2','Marca','Familia','Segmentación Inchcape', 'fc', 'fc_anas', 'fc_std',
    'Venta', 'Promedio de Venta 12M', 'Total instock', 'MAPE_AJUSTADO', 'ERP',
    'ERN', 'MAPE_ANASTASIA', 'ERP_ANASTASIA', 'ERN_ANASTASIA', 'MAPE_ESTADISTICO', 'ERP_ESTADISTICO','ERN_ESTADISTICO','WMAPE',
    'WMAPE.1', 'WMAPE.2', 'Mes', 'Costo Promedio Ponderado', 'Costo de Venta', 'Forecast en Costo', 'MAPE Costo',
    'ERP Costo', 'ERN Costo', 'WMAPE Costo', 'Input']]

# %%
rename_dict = {
    'ID2': 'Ultimo eslabon',
    'Periodo': 'Periodo (N-4/N-1)',
    'UE_2': 'SKU ERP',
    'Marca': 'Marca',
    'fc': 'FC Colaborado',
    'fc_anas': 'Forecast Base',
    'fc_std' : 'Forecast Estadistico',
    
    'Venta': 'Venta',
    'Promedio de Venta 12M': 'Prom Venta Real',
    'Total instock': 'Instock VtaCProm',

    'MAPE_AJUSTADO': 'MAPE Colaborado',
    'ERP': 'ERP Colab',
    'ERN': 'ERN Colab',

    'MAPE_ANASTASIA': "MAPE Base'",
    'ERP_ANASTASIA': "ERP Base'",
    'ERN_ANASTASIA': "ERN Base'",

    'MAPE_ESTADISTICO': "MAPE Estadistico",
    'ERP_ESTADISTICO': "ERP Estadistico",
    'ERN_ESTADISTICO': "ERN Estadistico",

  
    'WMAPE': 'WMAPE Colab',
    'WMAPE.1': "WMAPE Base'",
    'WMAPE.2': "WMAPE Estadistico",

    'Mes': 'Mes',
    'Costo Promedio Ponderado': 'Costo Control de Gestión',
    'Costo de Venta': 'Costo de Venta',
    'Forecast en Costo': 'Forecast en Costo',
    'MAPE Costo': 'MAPE Costo',
    'ERP Costo': 'ERP Costo',
    'ERN Costo': 'ERN Costo',
    'WMAPE Costo': 'WMAPE Costo',
    'Input': 'Input'
}



# %%
consolidado_cruce_ventas = consolidado_cruce_ventas.rename(columns=rename_dict)

# Display updated DataFrame



# %%
consolidado_cruce_ventas.columns

# %%


# %%
consolidado_cruce_ventas = consolidado_cruce_ventas[['SKU ERP', 'Ultimo eslabon', 'Familia', 'Marca', 'Segmentación Inchcape', 'Mes',
    'Periodo (N-4/N-1)', 'Input', 'Venta', 'FC Colaborado', 'MAPE Colaborado', 'ERP Colab', 'ERN Colab', 'WMAPE Colab',
    'Forecast Base', "MAPE Base'", "ERP Base'", "ERN Base'", "WMAPE Base'",
    'Forecast Estadistico', 'MAPE Estadistico',"ERP Estadistico","ERN Estadistico","WMAPE Estadistico",
    'Costo Control de Gestión', 'Costo de Venta', 'Forecast en Costo', 'MAPE Costo',
    'ERP Costo', 'ERN Costo', 'WMAPE Costo', 'Prom Venta Real', 'Instock VtaCProm']]
# %%
consolidado_cruce_ventas.to_excel(f'C:/Users/{nombre_usuario}/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras KPI-Reportes/Forecast Error (Python)/OEM Premium/Bases Forecast/{str(hoy.year)}-{str(hoy.month).zfill(2)} Indicador FC Error OEM Premium {nombre_numero_mes}_revision.xlsx', index=False)

print("Proceso finalizado correctamente🎊.")

# %%
codigos_imagen = [
    "B07129909320", "B07147321144", "B11517644809", "B11658623438", "B12415B90997",
    "C07129907398", "C07129907413", "C07129907433", "C07129908598", "C11128520212",
    "2440000004", "2440000005", "2440000006", "2440000008", "2440000010",
    "JJDE26444", "JJDE38685", "LFTC5431", "LLR034389", "LLR048842",
    "LLR079441", "LLR079444", "M07119906043", "M11228648147", "M36122408791",
    "N890225", "NNEUMATICOSDITEC", "O0W20CEPRO", "O0W20M", "O0W30C",
    "O964596G", "O965473G", "Z120-70R19CONTKC70", "Z170-60R17CON", "Z502_148_280_601",
    "P1T0863727", "P8K0253115J", "P95813384300", "P95835140301", "P95835140401",
    "V30622844", "V31213967", "V31316041", "V31341645", "V31370798"
]

# %%



