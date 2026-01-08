# %% [markdown]
# 
def main():
    # %%
    import pandas as pd
    import datetime
    import getpass
    usuario=getpass.getuser()

    print("⏰ Proceso de consolidación de MARA iniciado...")


    # %%
    mes_actual = str(datetime.datetime.today().month).zfill(2)
    año_actual = str(datetime.datetime.today().year)
    # %%


    # %%
    mara_derco = pd.read_csv(f"C:/Users/{usuario}/Inchcape/Planificación y Compras Chile - Documentos/Bases Indicadores en CSV {año_actual}-{mes_actual}/MARA_R3.csv",usecols=['Material_R3','Sector_dsc'], dtype={'Material_R3':'str'})

    # %%
    excluir = ["Subaru", "Repuesto Alternativo", "Rep.Alter.Maquinaria", "JCB", "JAC Truck", "DFSK", "Sector Comun", "Citroen", "Massey Ferguson", "Chevrolet", "Geely", "Landini", "Still", "Accesorios-Car Care", "Komatsu", "Claas", "Foton Pesados", "Neumáticos", "Valtra", "Implemento Agrícola", "Lubricantes", "Foton Ligeros", "NEUMATICOS", "Hangcha", "Europard", "Jacto", "Kverneland", "DS", "Zona Motors", "DFA", "Zongshen Motos", "Linde", "Farmtrac", "Baterías", "Kymco Motos", "JBC", "Piaggio Motos", "TCM", "SYM Motos", "Royal Enfield", "Hesston", "Derco Gas", "Stara", "JLG", "IVECO", "FAW", "Heli", "Yinxiang Motos", "Fiori", "Joylong", "Kesla", "Haval", "Repuestos Autoplanet", "Magni", "LUBRICANTES", "Equipo y Herramienta", "Haojue", "Otros Servicios","Hafei"]

    # %%
    mara_derco = mara_derco[~mara_derco['Sector_dsc'].isin(excluir)]

    # %%
    mara_inchcape = pd.read_excel(f"C:/Users/{usuario}/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras KPI-Reportes/Gerenciamiento MOS/Data Homologación/MARA/1. MARA.xlsx", sheet_name='Sheet1', dtype={'Material':'str'})



    # %%

    mara_inchcape['Marca'] = mara_inchcape['Sector MU'].apply(lambda x: 'Subaru' if x == 1335 else 'DFSK' if x == 1305 else 'Geely' if x == 1344 else 'Harley Davidson' if x==3261 else 'Unknown')

    # %%
    mara_bmw = pd.read_excel(f"C:/Users/{usuario}/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras KPI-Reportes/Gerenciamiento MOS/Data Homologación/MARA/1. MARA_BMW.xlsx",dtype={'Sector MU':'str'} )

    # %%

    mara_bmw['Marca'] = mara_bmw['Sector MU'].apply(lambda x: 'BMW' if x == '72' else 'Mini' if x == '73' else 'Motorrad' if x == '74' else 'Nacional BMW' if x == '75' else 'Unknown')

    # %%
    mara_ditec = pd.read_excel(f"C:/Users/{usuario}/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras KPI-Reportes/Gerenciamiento MOS/Data Homologación/MARA/MARA Ditec.xlsx",sheet_name='Hoja1')

    # %%
    mara_derco.rename(columns={'Material_R3':'Material','Sector_dsc':'Marca'},inplace=True)

    # %%
    mara_derco['Material'] = mara_derco['Material'].astype(str) + 'R3'

    # %%
    mara_derco['Tipo'] = 'OEM/AXS Derco'
    mara_derco = mara_derco[['Material','Marca', 'Tipo']]
    mara_inchcape['Tipo'] = 'OEM Inchcape'
    mara_inchcape = mara_inchcape[['Material','Marca', 'Tipo']]
    mara_ditec['Tipo'] = 'OEM DITEC'
    mara_ditec = mara_ditec[['Material','Marca', 'Tipo']]
    mara_bmw['Tipo'] = 'OEM BMW'
    mara_bmw = mara_bmw[['Material','Marca', 'Tipo']]

    # %%


    # %%
    mara_final = pd.concat([mara_derco,mara_inchcape,mara_bmw,mara_ditec])

    # %%
    import numpy as np
    cond = [mara_final['Marca'] == 'Harley Davidson']
    sel = ['OEM DITEC']
    mara_final['Tipo'] = np.select(cond, sel, mara_final['Tipo'])

    brecha = pd.read_excel(f"C:/Users/{usuario}/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras KPI-Reportes/Gerenciamiento MOS/Data Homologación/MARA/Código Brecha actualizados.xlsx")
    brecha.rename(columns={
        'Último Eslabón':'Material',
        'Sector MU':'Sector'}, inplace=True
    )
   
    brecha = brecha[["Material", "Marca"]]
    mara_final = mara_final[["Material", "Marca","Tipo"]]
    mara_consolidada = pd.concat([mara_final,brecha])
    mara_consolidada=mara_consolidada[mara_consolidada['Material'] != 'nanR3']
    mara_consolidada.to_csv(f"C:/Users/{usuario}/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras KPI-Reportes/Gerenciamiento MOS/Panel PBI/automatizacion/Mara_consolidada.csv")

    print("✅ Proceso de consolidación de MARA finalizado, archivo guardado en:\nPlanificación y Compras KPI-Reportes/Gerenciamiento MOS/Panel PBI/automatizacion/Mara_consolidada.csv")


if __name__ == '__main__':
    main()