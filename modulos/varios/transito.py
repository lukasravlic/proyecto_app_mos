def main():

    import pandas as pd
    import numpy as np
    import os
    import datetime
    import win32com.client
    import getpass
    from pathlib import Path

    # %%
    hoy= datetime.datetime.now()

    # %%



    nueva_carpeta = f"C:/Users/{usuario}/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras KPI-Reportes/Tubo Semanal/{hoy.strftime('%Y-%m-%d')}"
    if not os.path.exists(nueva_carpeta):
        os.makedirs(nueva_carpeta)
    destino_tubo = nueva_carpeta + '/' + f'{hoy.strftime('%Y-%m-%d')} - '

    nueva_carpeta = f"C:/Users/{usuario}/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras KPI-Reportes/Tubo Semanal/{hoy.strftime('%Y-%m-%d')}"
    if not os.path.exists(nueva_carpeta):
        os.makedirs(nueva_carpeta)
    print(nueva_carpeta + '/' + f'{hoy.strftime('%Y-%m-%d')} - ') 






    # ruta_carpeta_base = "C:/Users/{usuario}/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras Maestros/Base Planificable"
    # carpeta_base = os.listdir(ruta_carpeta_base)
    # for archivo in carpeta_base:
    #     if str(hoy.year) in archivo and f'{hoy.month:02d}' in archivo:
    #         ruta = ruta_carpeta_base + '/' + archivo
    #         ruta_2 = os.listdir(ruta)
    #         for archivo in ruta_2:
    #             if 'Base' in archivo:
    #                 print(ruta + '/' + archivo )
    #                 base = pd.read_excel(ruta + '/' + archivo, header=1, engine='openpyxl', sheet_name='Hoja1')
                

    # base_df_prov = base[['Proveedor', 'Pais (Proveedor)']]
    # base_df = base[['Material','Texto breve de material','NomSector_actual','TIPO','Corresponde','Proveedor', 'Pais (Proveedor)']]
    dict = {
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

    columnas = ['Material','Texto breve de material','NomSector_actual','TIPO','Corresponde','Proveedor', 'Pais (Proveedor)']

    base = pd.read_excel(f"C:/Users/{usuario}/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras Maestros/Base Planificable/{str(hoy.year)}-{str(hoy.month).zfill(2)} Base Planificable/Base {dict.get(str(hoy.month).zfill(2))} OEM-AXS v1.XLSX", header=1, engine='openpyxl', sheet_name='Base', usecols=columnas)
    base_df_prov = base[['Proveedor', 'Pais (Proveedor)']]
    base_df = base[['Material','Texto breve de material','NomSector_actual','TIPO','Corresponde','Proveedor', 'Pais (Proveedor)']]

    #BASE PLANIFICABLE


    # %%
    columnas = ['Material','Texto breve de material','NomSector_actual','TIPO','Corresponde','Proveedor', 'Pais (Proveedor)','Familia','Subfamilia']
    base = pd.read_excel(f"C:/Users/{usuario}/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras Maestros/Base Planificable/{str(hoy.year)}-{str(hoy.month).zfill(2)} Base Planificable/Base {dict.get(str(hoy.month).zfill(2))} OEM-AXS v1.XLSX", header=1, engine='openpyxl', sheet_name='Base', usecols=columnas)

    # %%



    usuario = getpass.getuser()
    ruta = f'C:/Users/{usuario}/Inchcape/Planificación y Compras Chile - Documentos/Bases Indicadores en CSV {hoy.year}-{hoy.month:02d}'
    ruta_repo = Path(ruta)

    # %%


    base_df = base_df[['Material','Texto breve de material','NomSector_actual','TIPO','Corresponde']]


    # %%
    columnas= ['Nro_pieza_fabricante_1',	'Cod_Actual_1']

    #columnas= ['Nro_pieza_fabricante_1',	'Cod_Actual_1']
    ruta_cod = ruta_repo.joinpath('COD_ACTUAL.csv')


    # Leer el archivo CSV en un DataFrame
    cod_actual_df = pd.read_csv(ruta_cod)
    cod_actual_df = cod_actual_df[columnas]

    # %%

    base_df_ue = pd.merge(base_df, cod_actual_df, left_on="Material", right_on="Nro_pieza_fabricante_1", how="left")
    base_df_ue['Cod_Actual_1'] = base_df_ue['Cod_Actual_1'].fillna(base_df_ue['Material'])
    base_df_ue = base_df_ue[['Cod_Actual_1','Texto breve de material', 'NomSector_actual', 'TIPO','Corresponde']]
    base_df_ue = base_df_ue.rename(columns={'Cod_Actual_1':'Material'})
    base_df_ue = base_df_ue.drop_duplicates(subset=['Material'])
    #MARA


    # Leer el archivo CSV en un DataFrame


    # %%

    columnas_mara = ['Material_R3','Part_number','Material_dsc','Modelo','Familia', 'Subfamilia', 'Categoría', 'Subcatgería','Sector_dsc']

    ruta_mara = ruta_repo.joinpath('MARA_R3.csv')
    df_mara = pd.read_csv(ruta_mara)

    print('Ruta Mara: ' + '\n' + str(ruta_mara))

    # %%
    df_mara = pd.read_csv(ruta_mara)

    print('Ruta Mara: ' + '\n' + str(ruta_mara))

    # %%


    import win32com.client
    import getpass
    usuario = getpass.getuser()

    # Initialize SAP GUI Scripting
    sap_gui_auto = win32com.client.GetObject("SAPGUI")
    application = sap_gui_auto.GetScriptingEngine

    # Establish connection and session
    connection = application.Children(0)
    session = connection.Children(0)

    # Maximize the window
    session.findById("wnd[0]").maximize()

    # Enter transaction code
    session.findById("wnd[0]/tbar[0]/okcd").text = "me5a"
    session.findById("wnd[0]").sendVKey(0)

    # Set plant code
    session.findById("wnd[0]/usr/ctxtS_WERKS-LOW").text = "0201"
    session.findById("wnd[0]/usr/ctxtS_WERKS-LOW").setFocus()
    session.findById("wnd[0]/usr/ctxtS_WERKS-LOW").caretPosition = 4
    session.findById("wnd[0]").sendVKey(0)

    # Set document type
    session.findById("wnd[0]/usr/ctxtS_BSART-LOW").setFocus()
    session.findById("wnd[0]/usr/ctxtS_BSART-LOW").caretPosition = 0
    session.findById("wnd[0]/usr/btn%_S_BSART_%_APP_%-VALU_PUSH").press()

    # Enter multiple values
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "zsto"
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "zspt"
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "zvor"
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").text = "zipl"
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,4]").text = "zatt"
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,4]").setFocus()
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,4]").caretPosition = 4
    session.findById("wnd[1]").sendVKey(8)

    # Set status
    session.findById("wnd[0]/usr/ctxtS_STATU-LOW").setFocus()
    session.findById("wnd[0]/usr/ctxtS_STATU-LOW").caretPosition = 0
    session.findById("wnd[0]/usr/btn%_S_STATU_%_APP_%-VALU_PUSH").press()
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "a"
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "n"
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").setFocus()
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").caretPosition = 1
    session.findById("wnd[1]").sendVKey(8)

    # Execute
    session.findById("wnd[0]").sendVKey(8)

    # Export to Excel
    session.findById("wnd[0]/tbar[1]/btn[33]").press()
    session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").setCurrentCell(0, "TEXT")
    session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").firstVisibleRow = 30
    session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectedRows = "0"
    session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").clickCurrentCell()
    session.findById("wnd[0]/tbar[1]/btn[43]").press()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:/Users/{usuario}/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras KPI-Reportes/Bases Transito (python)"
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "ME5A_R3.XLSX"
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 8
    session.findById("wnd[1]/tbar[0]/btn[11]").press()

    # Close the session
    session.findById("wnd[0]").sendVKey(3)
    session.findById("wnd[0]").sendVKey(3)




    # %%

    me5a_df_r3 = pd.read_excel("C:/Users/{usuario}/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras KPI-Reportes/Bases Transito (python)/ME5A_R3.XLSX")
    me5a_df_r3 = me5a_df_r3.merge(cod_actual_df, left_on='Material', right_on='Nro_pieza_fabricante_1',how='left')
    me5a_df_r3['Cod_Actual_1'].fillna(me5a_df_r3['Material'], inplace=True)
    me5a_df_r3.drop(columns= 'Nro_pieza_fabricante_1', inplace=True)
    me5a_df_r3.rename(columns={'Cod_Actual_1':'Ult. Eslabon'}, inplace=True)
    me5a_df_r3 = me5a_df_r3[['Solicitud de pedido', 'Clase documento', 'Fecha de solicitud',
        'Pos.solicitud pedido', 'Material','Ult. Eslabon',
        'Texto breve', 'Cantidad solicitada', 'Unidad de medida',
        'Nombre del proveedor', 'Indicador de borrado', 'Status tratamiento',
        'Centro', 'Status tratamiento solicitud pedido', 'Fecha de entrega',
        'Grupo de compras', 'Solicitante', 'Proveedor deseado',
        'Proveedor fijo', 'Reg.info de compras', 'Creado por',
        'Fecha de pedido', 'Nombre del proveedor deseado', 'Pedido',
        'Posición de pedido', 'Proveedor', 'Moneda', 'Precio de valoración',
        'Valor total', 'Petición de oferta',
        'Fecha Petición de oferta', 'Texto bloqueo', 'Cantidad confirmada','Urgencia necesidad'
        ]]

    me5a_df_r3.to_excel(destino_tubo + "ME5A R3.xlsx", index=False)




    # %%
    carpeta_tubo = "C:/Users/{usuario}/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras KPI-Reportes/Tubo Semanal"
    ruta_tubo = os.listdir(carpeta_tubo)[-2]

    # %%
    import win32com.client

    # Initialize SAP GUI scripting engine
    SapGuiAuto = win32com.client.GetObject("SAPGUI")
    application = SapGuiAuto.GetScriptingEngine

    # Get the first connection and session
    connection = application.Children(0)
    session = connection.Children(0)

    # Maximize the SAP window
    session.findById("wnd[0]").maximize()

    # Enter transaction code ME2L and execute
    session.findById("wnd[0]/tbar[0]/okcd").text = "me2l"
    session.findById("wnd[0]").sendVKey(0)

    # Set the "we101" text in the appropriate field
    session.findById("wnd[0]/usr/ctxtSELPA-LOW").text = "we101"

    # Set focus and position for the next field
    session.findById("wnd[0]/usr/ctxtS_BSART-LOW").setFocus()
    session.findById("wnd[0]/usr/ctxtS_BSART-LOW").caretPosition = 0

    # Press the value push button
    session.findById("wnd[0]/usr/btn%_S_BSART_%_APP_%-VALU_PUSH").press()

    # Fill in multiple values in a subsequent popup
    values = ["zsto", "zspt", "zvor", "zipl", "zatt"]
    for i, value in enumerate(values):
        session.findById(f"wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,{i}]").text = value

    # Set focus and caret position
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,4]").setFocus()
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,4]").caretPosition = 4

    # Send VKey 8 to confirm the entries
    session.findById("wnd[1]").sendVKey(8)

    # Set the plant to "0201"
    session.findById("wnd[0]/usr/ctxtS_WERKS-LOW").text = "0201"
    session.findById("wnd[0]/usr/ctxtS_WERKS-LOW").setFocus()
    session.findById("wnd[0]/usr/ctxtS_WERKS-LOW").caretPosition = 4

    # Send VKey 8 to confirm the plant entry
    session.findById("wnd[0]").sendVKey(8)

    # Execute the next steps
    session.findById("wnd[0]").sendVKey(23)
    session.findById("wnd[0]/tbar[1]/btn[33]").press()

    # Interact with the subsequent popup
    session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").setCurrentCell(8, "TEXT")
    session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").firstVisibleRow = 3
    session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectedRows = "8"
    session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").clickCurrentCell()
    session.findById("wnd[0]/tbar[1]/btn[43]").press()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()

    # Set file path and name
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\\Users\\{usuario}\\Inchcape\\Planificación y Compras Chile - Documentos\\Planificación y Compras KPI-Reportes\\Bases Transito (python)\\"
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "me2l_r3.XLSX"
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 7
    session.findById("wnd[1]/tbar[0]/btn[11]").press()

    # Close the windows
    session.findById("wnd[0]").sendVKey(3)
    session.findById("wnd[0]").sendVKey(3)



    import xlwings as xw
    try:
        book = xw.Book("C:/Users/{usuario}/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras KPI-Reportes/Bases Transito (python)/me2l_r3.XLSX")
        book.close()
    except Exception as e:
        print(e)


    # Alternatively, you can access the parts using indexing (e.g., df_parts[0], df_parts[1], etc.)

    # Each part will have a similar number of rows


    # %%
    me2l_df = pd.read_excel("C:/Users/{usuario}/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras KPI-Reportes/Bases Transito (python)/me2l_r3.xlsx", dtype={'Posición':'str', 'Documento compras':'str'})


    # %%
    me2l_df.shape

    # %%

    # #ME2L POR SEPARADO
    # me2l_df = me2l_df.drop(0)
    me2l_df['Por calcular (cantidad)'] = pd.to_numeric(me2l_df['Por calcular (cantidad)'], errors='coerce')
    me2l_df['Por entregar (cantidad)'] = pd.to_numeric(me2l_df['Por entregar (cantidad)'], errors='coerce')

    me2l_df = me2l_df[~((me2l_df['Por calcular (cantidad)']==0) & (me2l_df['Por entregar (cantidad)'] == 0))]

    me2l_df['Material'] = me2l_df['Material'].astype('str')
    me2l_df['Material'] = me2l_df['Material'].apply(lambda x: x.split(".")[0])

    me2l_ue = pd.merge(me2l_df, cod_actual_df, left_on="Material", right_on="Nro_pieza_fabricante_1", how="left")
    me2l_ue['Cod_Actual_1'].fillna(me2l_ue['Material'], inplace=True)
    me2l_ue = me2l_ue[['Documento compras', 'Posición', 'Reparto','Cl.documento compras',
        'Grupo de compras', 'Fecha documento', 'Material','Cod_Actual_1', 'Texto breve','Grupo de artículos',
        'Indicador de borrado','Tipo de posición', 'Tipo de imputación', 'Centro', 'Almacén', 'Cantidad de reparto', 'Cantidad de pedido',
        'Unidad medida pedido', 'Precio neto', 'Moneda', 'Cantidad base',
        'Fecha de entrega','Hora', 'Fecha entrega estad.',  'Cantidad entrada', 'Cantidad de salida', 'Cantidad entregada',
        'Solicitud de pedido', 'Pos.solicitud pedido', 'Indicador creación'
        , 'Nombre del proveedor',
        'Por entregar (cantidad)', 'Por entregar (valor)',
        'Por calcular (cantidad)', 'Por calcular (valor)'
        ]]

    me2l_ue = me2l_ue.rename(columns={'Material':'Material Antiguo'})
    me2l_ue = me2l_ue.rename(columns={'Cod_Actual_1':'Material'})

    me2l_df['Material'] = me2l_df['Material'].astype('str')
    me2l_df['Material'] = me2l_df['Material'].apply(lambda x: x.split(".")[0])

    me2l_ue = pd.merge(me2l_df, cod_actual_df, left_on="Material", right_on="Nro_pieza_fabricante_1", how="left")
    me2l_ue['Cod_Actual_1'].fillna(me2l_ue['Material'], inplace=True)

    me2l_ue = me2l_ue[['Documento compras', 'Posición', 'Reparto','Cl.documento compras',
        'Grupo de compras', 'Fecha documento', 'Material','Cod_Actual_1', 'Texto breve','Grupo de artículos',
        'Indicador de borrado','Tipo de posición', 'Tipo de imputación', 'Centro', 'Almacén', 'Cantidad de reparto', 'Cantidad de pedido',
        'Unidad medida pedido', 'Precio neto', 'Moneda', 'Cantidad base',
        'Fecha de entrega','Hora', 'Fecha entrega estad.',  'Cantidad entrada', 'Cantidad de salida', 'Cantidad entregada',
        'Solicitud de pedido', 'Pos.solicitud pedido', 'Indicador creación',
            'Nombre del proveedor',
        'Por entregar (cantidad)', 'Por entregar (valor)',
        'Por calcular (cantidad)', 'Por calcular (valor)'
        ]]

    me2l_ue = me2l_ue.rename(columns={'Material':'Material Antiguo'})
    me2l_ue = me2l_ue.rename(columns={'Cod_Actual_1':'Material'})

    base_df_prov = base_df_prov.drop_duplicates(subset=['Proveedor'])
    me2l_cruce_sector = pd.merge(me2l_ue, base_df_ue[['Material','Texto breve de material', 'NomSector_actual', 'TIPO','Corresponde']], left_on="Material", right_on="Material", how="left")
    #me2l_ue.to_excel("C:/Users/{usuario}/PROYECTOS DATA/PRUEBAS TRANSITO/me2l_sector.xlsx")
    me2l_cruce_sector.shape
    me2l_cruce_sector = me2l_cruce_sector[['Documento compras', 'Posición', 'Reparto','Cl.documento compras',
        'Grupo de compras', 'Fecha documento', 'Material Antiguo','Material', 'Texto breve','Grupo de artículos',
        'Indicador de borrado','Tipo de posición', 'Tipo de imputación', 'Centro', 'Almacén', 'Cantidad de reparto', 'Cantidad de pedido',
        'Unidad medida pedido', 'Precio neto', 'Moneda', 
        'Fecha de entrega','Hora', 'Fecha entrega estad.',
        'Cantidad entrada', 'Cantidad de salida', 'Cantidad entregada',
        'Solicitud de pedido', 'Pos.solicitud pedido', 'Indicador creación',
        'Nombre del proveedor',
        'Por entregar (cantidad)', 'Por entregar (valor)',
        'Por calcular (cantidad)', 'Por calcular (valor)', 'NomSector_actual',
        'Texto breve de material', 'TIPO', 'Corresponde']]
    me2l_cruce_sector['Cod_Prov'] = me2l_cruce_sector['Nombre del proveedor'].str.split(' ', expand=True)[0]
    me2l_cruce_sector = me2l_cruce_sector.merge(base_df_prov,left_on='Cod_Prov', right_on='Proveedor', how='left')
    me2l_cruce_sector['Posición'] = me2l_cruce_sector['Posición'].astype('str')
    me2l_cruce_sector['AUX'] = me2l_cruce_sector['Documento compras'] + me2l_cruce_sector['Posición']
    me2l_cruce_sector['Origen'] = me2l_cruce_sector['Pais (Proveedor)']
    me2l_cruce_sector = me2l_cruce_sector[['AUX','Documento compras', 'Posición', 'Reparto','Cl.documento compras',
        'Grupo de compras', 'Fecha documento', 'Material Antiguo','Material', 'Texto breve','Grupo de artículos',
        'Indicador de borrado','Tipo de posición', 'Tipo de imputación', 'Centro', 'Almacén', 'Cantidad de reparto', 'Cantidad de pedido',
        'Unidad medida pedido', 'Precio neto', 'Moneda',
        'Fecha de entrega','Hora', 'Fecha entrega estad.',
        'Cantidad entrada', 'Cantidad de salida', 'Cantidad entregada',
        'Solicitud de pedido', 'Pos.solicitud pedido', 'Indicador creación',
            'Nombre del proveedor',
        'Por entregar (cantidad)', 'Por entregar (valor)',
        'Por calcular (cantidad)', 'Por calcular (valor)', 'NomSector_actual','Origen', 'TIPO',  'Corresponde']]
    me2l_base_tr = me2l_cruce_sector
    #[me2l_cruce_sector['NomSector_actual'].notna()]
    me2l_base_tr['Corresponde'].fillna(0, inplace=True)
    me2l_base_tr['Corresponde'].replace({0:1}, inplace=True)




    # %%

    me2l_base_tr.to_excel(destino_tubo + 'ME2L R3.xlsx')
    me2l_cruce_sector = me2l_base_tr
    me2l_cruce_sector['AUX'] = me2l_cruce_sector['Documento compras'] + me2l_cruce_sector['Posición']
    me2l_cruce_sector = me2l_cruce_sector[['AUX','Documento compras', 'Posición', 'Reparto','Cl.documento compras',
        'Grupo de compras', 'Fecha documento', 'Material Antiguo','Material', 'Texto breve','Grupo de artículos',
        'Indicador de borrado','Tipo de posición', 'Tipo de imputación', 'Centro', 'Almacén', 'Cantidad de reparto', 'Cantidad de pedido',
        'Unidad medida pedido', 'Precio neto', 'Moneda',
        'Fecha de entrega','Hora', 'Fecha entrega estad.', 
        'Cantidad entrada', 'Cantidad de salida', 'Cantidad entregada',
        'Solicitud de pedido', 'Pos.solicitud pedido', 'Indicador creación',
            'Nombre del proveedor',
        'Por entregar (cantidad)', 'Por entregar (valor)',
        'Por calcular (cantidad)', 'Por calcular (valor)', 'NomSector_actual','Origen', 'TIPO',  'Corresponde']]
    me2l_base_tr = me2l_cruce_sector
    #[me2l_cruce_sector['NomSector_actual'].notna()]
    bases_oc = me2l_base_tr.groupby(['Documento compras'])['Posición'].count().sort_values(ascending=False).reset_index()


    # Calculate the number of rows in each part
    num_rows_per_part = len(bases_oc) // 3  # Integer division to get an equal split

    # Split the DataFrame into five parts




    # %%
    df_parts = np.array_split(bases_oc, 8)

    # %%
    df_parts[0]['Documento compras'].to_clipboard(header=False, index=False)

    # %%
    import win32com.client

    # Initialize SAP GUI scripting engine
    SapGuiAuto = win32com.client.GetObject("SAPGUI")
    application = SapGuiAuto.GetScriptingEngine

    # Get the first connection and session
    connection = application.Children(0)
    session = connection.Children(0)

    # Maximize the SAP window
    session.findById("wnd[0]").maximize()

    # Enter transaction code ZMM_MONITOR_ORDEN_CL and execute
    session.findById("wnd[0]/tbar[0]/okcd").text = "ZMM_MONITOR_ORDEN_CL"
    session.findById("wnd[0]").sendVKey(0)

    # Select the specific tab
    session.findById("wnd[0]/usr/tabsTABSTRIP_TABB2/tabpMON_OC_OC").select()

    # Set focus and position for the required field
    session.findById("wnd[0]/usr/tabsTABSTRIP_TABB2/tabpMON_OC_OC/ssub%_SUBSCREEN_TABB2:ZMM_MONITO3_SEGUIMIENTO_ORV4CL:1003/ctxtSO_OC-LOW").setFocus()
    session.findById("wnd[0]/usr/tabsTABSTRIP_TABB2/tabpMON_OC_OC/ssub%_SUBSCREEN_TABB2:ZMM_MONITO3_SEGUIMIENTO_ORV4CL:1003/ctxtSO_OC-LOW").caretPosition = 0

    # Press the value push button
    session.findById("wnd[0]/usr/tabsTABSTRIP_TABB2/tabpMON_OC_OC/ssub%_SUBSCREEN_TABB2:ZMM_MONITO3_SEGUIMIENTO_ORV4CL:1003/btn%_SO_OC_%_APP_%-VALU_PUSH").press()

    # Interact with the buttons in the new window
    session.findById("wnd[1]/tbar[0]/btn[24]").press()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/tbar[0]/btn[8]").press()

    # Send VKey 8 to confirm the entries
    session.findById("wnd[0]").sendVKey(8)

    # Press the toolbar button
    session.findById("wnd[0]/usr/cntlCONTROL_ALV_UPPER/shellcont/shell").pressToolbarButton("&MB_VARIANT")

    # Interact with the subsequent popup
    session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").setCurrentCell(4, "TEXT")
    session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectedRows = "4"
    session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").clickCurrentCell()

    # Export the data
    session.findById("wnd[0]/usr/cntlCONTROL_ALV_UPPER/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")
    session.findById("wnd[0]/usr/cntlCONTROL_ALV_UPPER/shellcont/shell").selectContextMenuItem("&XXL")
    session.findById("wnd[1]/tbar[0]/btn[0]").press()

    # Set the file path and name
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = ""
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "monitor_1_r3.XLSX"
    session.findById("wnd[1]/usr/ctxtDY_PATH").setFocus()
    session.findById("wnd[1]/usr/ctxtDY_PATH").caretPosition = 0
    session.findById("wnd[1]").sendVKey(4)

    # Set the directory path in the new window
    session.findById("wnd[2]/usr/ctxtDY_PATH").text = "C:\\Users\\{usuario}\\Inchcape\\Planificación y Compras Chile - Documentos\\Planificación y Compras KPI-Reportes\\Bases Transito (python)"
    session.findById("wnd[2]/usr/ctxtDY_PATH").setFocus()
    session.findById("wnd[2]/usr/ctxtDY_PATH").caretPosition = 131
    session.findById("wnd[2]/tbar[0]/btn[11]").press()

    # Confirm the export
    session.findById("wnd[1]/tbar[0]/btn[11]").press()

    # Close the windows
    session.findById("wnd[0]").sendVKey(3)
    session.findById("wnd[0]").sendVKey(3)


    # %%
    df_parts[1]['Documento compras'].to_clipboard(header=False, index=False)

    # %%
    import win32com.client

    # Initialize SAP GUI scripting engine
    SapGuiAuto = win32com.client.GetObject("SAPGUI")
    application = SapGuiAuto.GetScriptingEngine

    # Get the first connection and session
    connection = application.Children(0)
    session = connection.Children(0)

    # Maximize the SAP window
    session.findById("wnd[0]").maximize()

    # Enter transaction code ZMM_MONITOR_ORDEN_CL and execute
    session.findById("wnd[0]/tbar[0]/okcd").text = "ZMM_MONITOR_ORDEN_CL"
    session.findById("wnd[0]").sendVKey(0)

    # Select the specific tab
    session.findById("wnd[0]/usr/tabsTABSTRIP_TABB2/tabpMON_OC_OC").select()

    # Set focus and position for the required field
    session.findById("wnd[0]/usr/tabsTABSTRIP_TABB2/tabpMON_OC_OC/ssub%_SUBSCREEN_TABB2:ZMM_MONITO3_SEGUIMIENTO_ORV4CL:1003/ctxtSO_OC-LOW").setFocus()
    session.findById("wnd[0]/usr/tabsTABSTRIP_TABB2/tabpMON_OC_OC/ssub%_SUBSCREEN_TABB2:ZMM_MONITO3_SEGUIMIENTO_ORV4CL:1003/ctxtSO_OC-LOW").caretPosition = 0

    # Press the value push button
    session.findById("wnd[0]/usr/tabsTABSTRIP_TABB2/tabpMON_OC_OC/ssub%_SUBSCREEN_TABB2:ZMM_MONITO3_SEGUIMIENTO_ORV4CL:1003/btn%_SO_OC_%_APP_%-VALU_PUSH").press()

    # Interact with the buttons in the new window
    session.findById("wnd[1]/tbar[0]/btn[24]").press()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/tbar[0]/btn[8]").press()

    # Send VKey 8 to confirm the entries
    session.findById("wnd[0]").sendVKey(8)

    # Press the toolbar button
    session.findById("wnd[0]/usr/cntlCONTROL_ALV_UPPER/shellcont/shell").pressToolbarButton("&MB_VARIANT")

    # Interact with the subsequent popup
    session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").setCurrentCell(4, "TEXT")
    session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectedRows = "4"
    session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").clickCurrentCell()

    # Export the data
    session.findById("wnd[0]/usr/cntlCONTROL_ALV_UPPER/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")
    session.findById("wnd[0]/usr/cntlCONTROL_ALV_UPPER/shellcont/shell").selectContextMenuItem("&XXL")
    session.findById("wnd[1]/tbar[0]/btn[0]").press()

    # Set the file path and name
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = ""
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "monitor_2_r3.XLSX"
    session.findById("wnd[1]/usr/ctxtDY_PATH").setFocus()
    session.findById("wnd[1]/usr/ctxtDY_PATH").caretPosition = 0
    session.findById("wnd[1]").sendVKey(4)

    # Set the directory path in the new window
    session.findById("wnd[2]/usr/ctxtDY_PATH").text = "C:\\Users\\{usuario}\\Inchcape\\Planificación y Compras Chile - Documentos\\Planificación y Compras KPI-Reportes\\Bases Transito (python)"
    session.findById("wnd[2]/usr/ctxtDY_PATH").setFocus()
    session.findById("wnd[2]/usr/ctxtDY_PATH").caretPosition = 131
    session.findById("wnd[2]/tbar[0]/btn[11]").press()

    # Confirm the export
    session.findById("wnd[1]/tbar[0]/btn[11]").press()

    # Close the windows
    session.findById("wnd[0]").sendVKey(3)
    session.findById("wnd[0]").sendVKey(3)


    # %%
    # import xlwings as xw
    # try:
    #     book = xw.Book("C:/Users/{usuario}/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras KPI-Reportes/Bases Transito (python)/monitor_2_r3.XLSX")
    #     book.close()
    # except Exception as e:
    #     print(e)

    # %%
    df_parts[2]['Documento compras'].to_clipboard(header=False, index=False)

    # %%


    import win32com.client

    # Initialize SAP GUI scripting engine
    SapGuiAuto = win32com.client.GetObject("SAPGUI")
    application = SapGuiAuto.GetScriptingEngine

    # Get the first connection and session
    connection = application.Children(0)
    session = connection.Children(0)

    # Maximize the SAP window
    session.findById("wnd[0]").maximize()

    # Enter transaction code ZMM_MONITOR_ORDEN_CL and execute
    session.findById("wnd[0]/tbar[0]/okcd").text = "ZMM_MONITOR_ORDEN_CL"
    session.findById("wnd[0]").sendVKey(0)

    # Select the specific tab
    session.findById("wnd[0]/usr/tabsTABSTRIP_TABB2/tabpMON_OC_OC").select()

    # Set focus and position for the required field
    session.findById("wnd[0]/usr/tabsTABSTRIP_TABB2/tabpMON_OC_OC/ssub%_SUBSCREEN_TABB2:ZMM_MONITO3_SEGUIMIENTO_ORV4CL:1003/ctxtSO_OC-LOW").setFocus()
    session.findById("wnd[0]/usr/tabsTABSTRIP_TABB2/tabpMON_OC_OC/ssub%_SUBSCREEN_TABB2:ZMM_MONITO3_SEGUIMIENTO_ORV4CL:1003/ctxtSO_OC-LOW").caretPosition = 0

    # Press the value push button
    session.findById("wnd[0]/usr/tabsTABSTRIP_TABB2/tabpMON_OC_OC/ssub%_SUBSCREEN_TABB2:ZMM_MONITO3_SEGUIMIENTO_ORV4CL:1003/btn%_SO_OC_%_APP_%-VALU_PUSH").press()

    # Interact with the buttons in the new window
    session.findById("wnd[1]/tbar[0]/btn[24]").press()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/tbar[0]/btn[8]").press()

    # Send VKey 8 to confirm the entries
    session.findById("wnd[0]").sendVKey(8)

    # Press the toolbar button
    session.findById("wnd[0]/usr/cntlCONTROL_ALV_UPPER/shellcont/shell").pressToolbarButton("&MB_VARIANT")

    # Interact with the subsequent popup
    session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").setCurrentCell(4, "TEXT")
    session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectedRows = "4"
    session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").clickCurrentCell()

    # Export the data
    session.findById("wnd[0]/usr/cntlCONTROL_ALV_UPPER/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")
    session.findById("wnd[0]/usr/cntlCONTROL_ALV_UPPER/shellcont/shell").selectContextMenuItem("&XXL")
    session.findById("wnd[1]/tbar[0]/btn[0]").press()

    # Set the file path and name
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = ""
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "monitor_3_r3.XLSX"
    session.findById("wnd[1]/usr/ctxtDY_PATH").setFocus()
    session.findById("wnd[1]/usr/ctxtDY_PATH").caretPosition = 0
    session.findById("wnd[1]").sendVKey(4)

    # Set the directory path in the new window
    session.findById("wnd[2]/usr/ctxtDY_PATH").text = "C:\\Users\\{usuario}\\Inchcape\\Planificación y Compras Chile - Documentos\\Planificación y Compras KPI-Reportes\\Bases Transito (python)"
    session.findById("wnd[2]/usr/ctxtDY_PATH").setFocus()
    session.findById("wnd[2]/usr/ctxtDY_PATH").caretPosition = 131
    session.findById("wnd[2]/tbar[0]/btn[11]").press()

    # Confirm the export
    session.findById("wnd[1]/tbar[0]/btn[11]").press()

    # Close the windows
    session.findById("wnd[0]").sendVKey(3)
    session.findById("wnd[0]").sendVKey(3)


    # %%
    # import xlwings as xw
    # try:
    #     book = xw.Book("C:/Users/{usuario}/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras KPI-Reportes/Bases Transito (python)/monitor_3_r3.XLSX")
    #     book.close()
    # except Exception as e:
    #     print(e)

    # %%
    df_parts[3]['Documento compras'].to_clipboard(header=False, index=False)

    # %%


    import win32com.client

    # Initialize SAP GUI scripting engine
    SapGuiAuto = win32com.client.GetObject("SAPGUI")
    application = SapGuiAuto.GetScriptingEngine

    # Get the first connection and session
    connection = application.Children(0)
    session = connection.Children(0)

    # Maximize the SAP window
    session.findById("wnd[0]").maximize()

    # Enter transaction code ZMM_MONITOR_ORDEN_CL and execute
    session.findById("wnd[0]/tbar[0]/okcd").text = "ZMM_MONITOR_ORDEN_CL"
    session.findById("wnd[0]").sendVKey(0)

    # Select the specific tab
    session.findById("wnd[0]/usr/tabsTABSTRIP_TABB2/tabpMON_OC_OC").select()

    # Set focus and position for the required field
    session.findById("wnd[0]/usr/tabsTABSTRIP_TABB2/tabpMON_OC_OC/ssub%_SUBSCREEN_TABB2:ZMM_MONITO3_SEGUIMIENTO_ORV4CL:1003/ctxtSO_OC-LOW").setFocus()
    session.findById("wnd[0]/usr/tabsTABSTRIP_TABB2/tabpMON_OC_OC/ssub%_SUBSCREEN_TABB2:ZMM_MONITO3_SEGUIMIENTO_ORV4CL:1003/ctxtSO_OC-LOW").caretPosition = 0

    # Press the value push button
    session.findById("wnd[0]/usr/tabsTABSTRIP_TABB2/tabpMON_OC_OC/ssub%_SUBSCREEN_TABB2:ZMM_MONITO3_SEGUIMIENTO_ORV4CL:1003/btn%_SO_OC_%_APP_%-VALU_PUSH").press()

    # Interact with the buttons in the new window
    session.findById("wnd[1]/tbar[0]/btn[24]").press()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/tbar[0]/btn[8]").press()

    # Send VKey 8 to confirm the entries
    session.findById("wnd[0]").sendVKey(8)

    # Press the toolbar button
    session.findById("wnd[0]/usr/cntlCONTROL_ALV_UPPER/shellcont/shell").pressToolbarButton("&MB_VARIANT")

    # Interact with the subsequent popup
    session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").setCurrentCell(4, "TEXT")
    session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectedRows = "4"
    session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").clickCurrentCell()

    # Export the data
    session.findById("wnd[0]/usr/cntlCONTROL_ALV_UPPER/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")
    session.findById("wnd[0]/usr/cntlCONTROL_ALV_UPPER/shellcont/shell").selectContextMenuItem("&XXL")
    session.findById("wnd[1]/tbar[0]/btn[0]").press()

    # Set the file path and name
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = ""
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "monitor_4_r3.XLSX"
    session.findById("wnd[1]/usr/ctxtDY_PATH").setFocus()
    session.findById("wnd[1]/usr/ctxtDY_PATH").caretPosition = 0
    session.findById("wnd[1]").sendVKey(4)

    # Set the directory path in the new window
    session.findById("wnd[2]/usr/ctxtDY_PATH").text = "C:\\Users\\{usuario}\\Inchcape\\Planificación y Compras Chile - Documentos\\Planificación y Compras KPI-Reportes\\Bases Transito (python)"
    session.findById("wnd[2]/usr/ctxtDY_PATH").setFocus()
    session.findById("wnd[2]/usr/ctxtDY_PATH").caretPosition = 131
    session.findById("wnd[2]/tbar[0]/btn[11]").press()

    # Confirm the export
    session.findById("wnd[1]/tbar[0]/btn[11]").press()

    # Close the windows
    session.findById("wnd[0]").sendVKey(3)
    session.findById("wnd[0]").sendVKey(3)

    # %%
    df_parts[4]['Documento compras'].to_clipboard(header=False, index=False)

    import win32com.client

    # Initialize SAP GUI scripting engine
    SapGuiAuto = win32com.client.GetObject("SAPGUI")
    application = SapGuiAuto.GetScriptingEngine

    # Get the first connection and session
    connection = application.Children(0)
    session = connection.Children(0)

    # Maximize the SAP window
    session.findById("wnd[0]").maximize()

    # Enter transaction code ZMM_MONITOR_ORDEN_CL and execute
    session.findById("wnd[0]/tbar[0]/okcd").text = "ZMM_MONITOR_ORDEN_CL"
    session.findById("wnd[0]").sendVKey(0)

    # Select the specific tab
    session.findById("wnd[0]/usr/tabsTABSTRIP_TABB2/tabpMON_OC_OC").select()

    # Set focus and position for the required field
    session.findById("wnd[0]/usr/tabsTABSTRIP_TABB2/tabpMON_OC_OC/ssub%_SUBSCREEN_TABB2:ZMM_MONITO3_SEGUIMIENTO_ORV4CL:1003/ctxtSO_OC-LOW").setFocus()
    session.findById("wnd[0]/usr/tabsTABSTRIP_TABB2/tabpMON_OC_OC/ssub%_SUBSCREEN_TABB2:ZMM_MONITO3_SEGUIMIENTO_ORV4CL:1003/ctxtSO_OC-LOW").caretPosition = 0

    # Press the value push button
    session.findById("wnd[0]/usr/tabsTABSTRIP_TABB2/tabpMON_OC_OC/ssub%_SUBSCREEN_TABB2:ZMM_MONITO3_SEGUIMIENTO_ORV4CL:1003/btn%_SO_OC_%_APP_%-VALU_PUSH").press()

    # Interact with the buttons in the new window
    session.findById("wnd[1]/tbar[0]/btn[24]").press()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/tbar[0]/btn[8]").press()

    # Send VKey 8 to confirm the entries
    session.findById("wnd[0]").sendVKey(8)

    # Press the toolbar button
    session.findById("wnd[0]/usr/cntlCONTROL_ALV_UPPER/shellcont/shell").pressToolbarButton("&MB_VARIANT")

    # Interact with the subsequent popup
    session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").setCurrentCell(4, "TEXT")
    session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectedRows = "4"
    session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").clickCurrentCell()

    # Export the data
    session.findById("wnd[0]/usr/cntlCONTROL_ALV_UPPER/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")
    session.findById("wnd[0]/usr/cntlCONTROL_ALV_UPPER/shellcont/shell").selectContextMenuItem("&XXL")
    session.findById("wnd[1]/tbar[0]/btn[0]").press()

    # Set the file path and name
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = ""
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "monitor_5_r3.XLSX"
    session.findById("wnd[1]/usr/ctxtDY_PATH").setFocus()
    session.findById("wnd[1]/usr/ctxtDY_PATH").caretPosition = 0
    session.findById("wnd[1]").sendVKey(4)

    # Set the directory path in the new window
    session.findById("wnd[2]/usr/ctxtDY_PATH").text = "C:\\Users\\{usuario}\\Inchcape\\Planificación y Compras Chile - Documentos\\Planificación y Compras KPI-Reportes\\Bases Transito (python)"
    session.findById("wnd[2]/usr/ctxtDY_PATH").setFocus()
    session.findById("wnd[2]/usr/ctxtDY_PATH").caretPosition = 131
    session.findById("wnd[2]/tbar[0]/btn[11]").press()

    # Confirm the export
    session.findById("wnd[1]/tbar[0]/btn[11]").press()

    # Close the windows
    session.findById("wnd[0]").sendVKey(3)
    session.findById("wnd[0]").sendVKey(3)

    # %%
    df_parts[5]['Documento compras'].to_clipboard(header=False, index=False)

    import win32com.client

    # Initialize SAP GUI scripting engine
    SapGuiAuto = win32com.client.GetObject("SAPGUI")
    application = SapGuiAuto.GetScriptingEngine

    # Get the first connection and session
    connection = application.Children(0)
    session = connection.Children(0)

    # Maximize the SAP window
    session.findById("wnd[0]").maximize()

    # Enter transaction code ZMM_MONITOR_ORDEN_CL and execute
    session.findById("wnd[0]/tbar[0]/okcd").text = "ZMM_MONITOR_ORDEN_CL"
    session.findById("wnd[0]").sendVKey(0)

    # Select the specific tab
    session.findById("wnd[0]/usr/tabsTABSTRIP_TABB2/tabpMON_OC_OC").select()

    # Set focus and position for the required field
    session.findById("wnd[0]/usr/tabsTABSTRIP_TABB2/tabpMON_OC_OC/ssub%_SUBSCREEN_TABB2:ZMM_MONITO3_SEGUIMIENTO_ORV4CL:1003/ctxtSO_OC-LOW").setFocus()
    session.findById("wnd[0]/usr/tabsTABSTRIP_TABB2/tabpMON_OC_OC/ssub%_SUBSCREEN_TABB2:ZMM_MONITO3_SEGUIMIENTO_ORV4CL:1003/ctxtSO_OC-LOW").caretPosition = 0

    # Press the value push button
    session.findById("wnd[0]/usr/tabsTABSTRIP_TABB2/tabpMON_OC_OC/ssub%_SUBSCREEN_TABB2:ZMM_MONITO3_SEGUIMIENTO_ORV4CL:1003/btn%_SO_OC_%_APP_%-VALU_PUSH").press()

    # Interact with the buttons in the new window
    session.findById("wnd[1]/tbar[0]/btn[24]").press()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/tbar[0]/btn[8]").press()

    # Send VKey 8 to confirm the entries
    session.findById("wnd[0]").sendVKey(8)

    # Press the toolbar button
    session.findById("wnd[0]/usr/cntlCONTROL_ALV_UPPER/shellcont/shell").pressToolbarButton("&MB_VARIANT")

    # Interact with the subsequent popup
    session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").setCurrentCell(4, "TEXT")
    session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectedRows = "4"
    session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").clickCurrentCell()

    # Export the data
    session.findById("wnd[0]/usr/cntlCONTROL_ALV_UPPER/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")
    session.findById("wnd[0]/usr/cntlCONTROL_ALV_UPPER/shellcont/shell").selectContextMenuItem("&XXL")
    session.findById("wnd[1]/tbar[0]/btn[0]").press()

    # Set the file path and name
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = ""
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "monitor_6_r3.XLSX"
    session.findById("wnd[1]/usr/ctxtDY_PATH").setFocus()
    session.findById("wnd[1]/usr/ctxtDY_PATH").caretPosition = 0
    session.findById("wnd[1]").sendVKey(4)

    # Set the directory path in the new window
    session.findById("wnd[2]/usr/ctxtDY_PATH").text = "C:\\Users\\{usuario}\\Inchcape\\Planificación y Compras Chile - Documentos\\Planificación y Compras KPI-Reportes\\Bases Transito (python)"
    session.findById("wnd[2]/usr/ctxtDY_PATH").setFocus()
    session.findById("wnd[2]/usr/ctxtDY_PATH").caretPosition = 131
    session.findById("wnd[2]/tbar[0]/btn[11]").press()

    # Confirm the export
    session.findById("wnd[1]/tbar[0]/btn[11]").press()

    # Close the windows
    session.findById("wnd[0]").sendVKey(3)
    session.findById("wnd[0]").sendVKey(3)

    # %%
    df_parts[6]['Documento compras'].to_clipboard(header=False, index=False)

    # %%


    import win32com.client

    # Initialize SAP GUI scripting engine
    SapGuiAuto = win32com.client.GetObject("SAPGUI")
    application = SapGuiAuto.GetScriptingEngine

    # Get the first connection and session
    connection = application.Children(0)
    session = connection.Children(0)

    # Maximize the SAP window
    session.findById("wnd[0]").maximize()

    # Enter transaction code ZMM_MONITOR_ORDEN_CL and execute
    session.findById("wnd[0]/tbar[0]/okcd").text = "ZMM_MONITOR_ORDEN_CL"
    session.findById("wnd[0]").sendVKey(0)

    # Select the specific tab
    session.findById("wnd[0]/usr/tabsTABSTRIP_TABB2/tabpMON_OC_OC").select()

    # Set focus and position for the required field
    session.findById("wnd[0]/usr/tabsTABSTRIP_TABB2/tabpMON_OC_OC/ssub%_SUBSCREEN_TABB2:ZMM_MONITO3_SEGUIMIENTO_ORV4CL:1003/ctxtSO_OC-LOW").setFocus()
    session.findById("wnd[0]/usr/tabsTABSTRIP_TABB2/tabpMON_OC_OC/ssub%_SUBSCREEN_TABB2:ZMM_MONITO3_SEGUIMIENTO_ORV4CL:1003/ctxtSO_OC-LOW").caretPosition = 0

    # Press the value push button
    session.findById("wnd[0]/usr/tabsTABSTRIP_TABB2/tabpMON_OC_OC/ssub%_SUBSCREEN_TABB2:ZMM_MONITO3_SEGUIMIENTO_ORV4CL:1003/btn%_SO_OC_%_APP_%-VALU_PUSH").press()

    # Interact with the buttons in the new window
    session.findById("wnd[1]/tbar[0]/btn[24]").press()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/tbar[0]/btn[8]").press()

    # Send VKey 8 to confirm the entries
    session.findById("wnd[0]").sendVKey(8)

    # Press the toolbar button
    session.findById("wnd[0]/usr/cntlCONTROL_ALV_UPPER/shellcont/shell").pressToolbarButton("&MB_VARIANT")

    # Interact with the subsequent popup
    session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").setCurrentCell(4, "TEXT")
    session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectedRows = "4"
    session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").clickCurrentCell()

    # Export the data
    session.findById("wnd[0]/usr/cntlCONTROL_ALV_UPPER/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")
    session.findById("wnd[0]/usr/cntlCONTROL_ALV_UPPER/shellcont/shell").selectContextMenuItem("&XXL")
    session.findById("wnd[1]/tbar[0]/btn[0]").press()

    # Set the file path and name
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = ""
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "monitor_7_r3.XLSX"
    session.findById("wnd[1]/usr/ctxtDY_PATH").setFocus()
    session.findById("wnd[1]/usr/ctxtDY_PATH").caretPosition = 0
    session.findById("wnd[1]").sendVKey(4)

    # Set the directory path in the new window
    session.findById("wnd[2]/usr/ctxtDY_PATH").text = "C:\\Users\\{usuario}\\Inchcape\\Planificación y Compras Chile - Documentos\\Planificación y Compras KPI-Reportes\\Bases Transito (python)"
    session.findById("wnd[2]/usr/ctxtDY_PATH").setFocus()
    session.findById("wnd[2]/usr/ctxtDY_PATH").caretPosition = 131
    session.findById("wnd[2]/tbar[0]/btn[11]").press()

    # Confirm the export
    session.findById("wnd[1]/tbar[0]/btn[11]").press()

    # Close the windows
    session.findById("wnd[0]").sendVKey(3)
    session.findById("wnd[0]").sendVKey(3)

    # %%
    df_parts[7]['Documento compras'].to_clipboard(header=False, index=False)

    # %%


    import win32com.client

    # Initialize SAP GUI scripting engine
    SapGuiAuto = win32com.client.GetObject("SAPGUI")
    application = SapGuiAuto.GetScriptingEngine

    # Get the first connection and session
    connection = application.Children(0)
    session = connection.Children(0)

    # Maximize the SAP window
    session.findById("wnd[0]").maximize()

    # Enter transaction code ZMM_MONITOR_ORDEN_CL and execute
    session.findById("wnd[0]/tbar[0]/okcd").text = "ZMM_MONITOR_ORDEN_CL"
    session.findById("wnd[0]").sendVKey(0)

    # Select the specific tab
    session.findById("wnd[0]/usr/tabsTABSTRIP_TABB2/tabpMON_OC_OC").select()

    # Set focus and position for the required field
    session.findById("wnd[0]/usr/tabsTABSTRIP_TABB2/tabpMON_OC_OC/ssub%_SUBSCREEN_TABB2:ZMM_MONITO3_SEGUIMIENTO_ORV4CL:1003/ctxtSO_OC-LOW").setFocus()
    session.findById("wnd[0]/usr/tabsTABSTRIP_TABB2/tabpMON_OC_OC/ssub%_SUBSCREEN_TABB2:ZMM_MONITO3_SEGUIMIENTO_ORV4CL:1003/ctxtSO_OC-LOW").caretPosition = 0

    # Press the value push button
    session.findById("wnd[0]/usr/tabsTABSTRIP_TABB2/tabpMON_OC_OC/ssub%_SUBSCREEN_TABB2:ZMM_MONITO3_SEGUIMIENTO_ORV4CL:1003/btn%_SO_OC_%_APP_%-VALU_PUSH").press()

    # Interact with the buttons in the new window
    session.findById("wnd[1]/tbar[0]/btn[24]").press()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/tbar[0]/btn[8]").press()

    # Send VKey 8 to confirm the entries
    session.findById("wnd[0]").sendVKey(8)

    # Press the toolbar button
    session.findById("wnd[0]/usr/cntlCONTROL_ALV_UPPER/shellcont/shell").pressToolbarButton("&MB_VARIANT")

    # Interact with the subsequent popup
    session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").setCurrentCell(4, "TEXT")
    session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectedRows = "4"
    session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").clickCurrentCell()

    # Export the data
    session.findById("wnd[0]/usr/cntlCONTROL_ALV_UPPER/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")
    session.findById("wnd[0]/usr/cntlCONTROL_ALV_UPPER/shellcont/shell").selectContextMenuItem("&XXL")
    session.findById("wnd[1]/tbar[0]/btn[0]").press()

    # Set the file path and name
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = ""
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "monitor_8_r3.XLSX"
    session.findById("wnd[1]/usr/ctxtDY_PATH").setFocus()
    session.findById("wnd[1]/usr/ctxtDY_PATH").caretPosition = 0
    session.findById("wnd[1]").sendVKey(4)

    # Set the directory path in the new window
    session.findById("wnd[2]/usr/ctxtDY_PATH").text = "C:\\Users\\{usuario}\\Inchcape\\Planificación y Compras Chile - Documentos\\Planificación y Compras KPI-Reportes\\Bases Transito (python)"
    session.findById("wnd[2]/usr/ctxtDY_PATH").setFocus()
    session.findById("wnd[2]/usr/ctxtDY_PATH").caretPosition = 131
    session.findById("wnd[2]/tbar[0]/btn[11]").press()

    # Confirm the export
    session.findById("wnd[1]/tbar[0]/btn[11]").press()

    # Close the windows
    session.findById("wnd[0]").sendVKey(3)
    session.findById("wnd[0]").sendVKey(3)

    # %%
    bases_oc['Documento compras'].to_clipboard(header=False, index=False)

    # %%
    SapGuiAuto = win32com.client.GetObject("SAPGUI")
    application = SapGuiAuto.GetScriptingEngine
    connection = application.Children(0)
    session = connection.Children(0)
    #now = datetime.now()

    #session.findById("wnd[0]").maximize
    #session.findById("wnd[0]").maximize
    session.findById("wnd[0]/tbar[0]/okcd").text = "vl06if"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/usr/ctxtIT_LFDAT-LOW").text = ""
    session.findById("wnd[0]/usr/ctxtIT_LFDAT-HIGH").text = ""
    session.findById("wnd[0]/usr/ctxtIT_LFDAT-LOW").setFocus()
    session.findById("wnd[0]/usr/ctxtIT_LFDAT-LOW").caretPosition = 0
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/usr/ctxtIT_LFDAT-HIGH").setFocus()
    session.findById("wnd[0]/usr/ctxtIT_LFDAT-HIGH").caretPosition = 0
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/usr/btn%_IT_EBELN_%_APP_%-VALU_PUSH").press()
    session.findById("wnd[1]/tbar[0]/btn[24]").press()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/tbar[0]/btn[8]").press()
    session.findById("wnd[0]").sendVKey(8)
    session.findById("wnd[0]/tbar[1]/btn[18]").press()
    session.findById("wnd[0]").sendVKey(33)
    session.findById("wnd[1]/usr/lbl[1,7]").setFocus()
    session.findById("wnd[1]/usr/lbl[1,7]").caretPosition = 4
    session.findById("wnd[1]").sendVKey(2)
    session.findById("wnd[0]").sendVKey(43)
    session.findById("wnd[1]/usr/cmbG_LISTBOX").setFocus()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\\Users\\{usuario}\\Inchcape\\Planificación y Compras Chile - Documentos\\Planificación y Compras KPI-Reportes\\Bases Transito (python)"
    session.findById("wnd[1]/tbar[0]/btn[11]").press()
    session.findById("wnd[0]").sendVKey(3)
    session.findById("wnd[0]").sendVKey(3)




    # %%
    # import xlwings as xw
    # try:
    #     book = xw.Book("C:/Users/{usuario}/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras KPI-Reportes/Bases Transito (python)/export.XLSX")
    #     book.close()
    # except Exception as e:
    #     print(e)


    # %%
    lista = os.listdir("C:/Users/{usuario}/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras KPI-Reportes/Bases Transito (python)")

    # %%
    dfs = []
    print('Archivos usados (Monitor)')
    for i in lista:
        if ('monitor_' in i):
            df = pd.read_excel("C:/Users/{usuario}/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras KPI-Reportes/Bases Transito (python)" + '/' +i, dtype={'N° MIRO':'str'},index_col=None)
            print(i)
            print(df.shape)
            dfs.append(df)


    # %%
    df_ee = pd.read_excel("C:/Users/{usuario}/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras KPI-Reportes/Bases Transito (python)/export.XLSX" ,dtype={'Documento compras':'str','Posición modelo':'str'})

    # %%
    print(f'Dimensiones del df1\n: FILAS:{df_ee.shape[0]}\nCOLUMNAS: {df_ee.shape[1]}')

    # %%
    monitor_df = pd.concat(dfs,axis=0, ignore_index=True)

    # %%
    monitor_df.columns

    # %%
    monitor_df.to_excel(destino_tubo + "Monitor R3.xlsx")


    # %%
    df_ee.to_excel(destino_tubo + "Entrega Entrante R3.xlsx")

    # %%
    #monitor_df = monitor_df[['Nro. OC SAP','Clase Doc. OC','Posición OC SAP','Ind. Borrado OC','Clase Doc. SP','Nro. SOLPED','Posición SOLPED','Ind. Borrado SP','Cant. UN. OC','N° MIRO','Posición Factura',
    #'Cant. Factura','Fe. Confirmación','Fe. Compromiso','Fe. ATA','TF de ATA','Cantidad en BO','Ind. Borrado PO']]

    # %%
    monitor_df = monitor_df[monitor_df['Ind. Borrado OC'].isna() &
        monitor_df['Ind. Borrado SP'].isna() &
        monitor_df['Ind. Borrado PO'].isna()]

    # %%


    # %%
    monitor_df['Nro. OC SAP'] = monitor_df['Nro. OC SAP'].astype('str')
    monitor_df['Posición OC SAP'] = monitor_df['Posición OC SAP'].astype('str')


    # %%
    monitor_df['AUX'] = monitor_df['Nro. OC SAP'] + monitor_df['Posición OC SAP']

    # %%
    monitor_df['N° MIRO'] = monitor_df['N° MIRO'].astype('str')

    # %%
    monitor_df['Posición Factura'] = monitor_df['Posición Factura'].astype('str')

    # %%
    monitor_df['AUX_FACTURA'] = monitor_df['N° MIRO'] + monitor_df['Posición Factura']

    # %%
    #revision = monitor_df.groupby('AUX')['Fe. ATA'].nunique().sort_values(ascending=False).reset_index()

    # %%


    # %%
    tabla_1 = monitor_df[monitor_df['Ind. Borrado OC'].isna() &
        monitor_df['Ind. Borrado SP'].isna() &
        monitor_df['Ind. Borrado PO'].isna()].groupby(['AUX', 'AUX_FACTURA']).agg({'Cant. UN. OC':'mean', 'Cant. Factura':'max','Cantidad en BO':'mean'})

    # %%
    tabla_final = tabla_1.groupby(['AUX']).agg({'Cant. UN. OC':'mean','Cant. Factura':'sum', 'Cantidad en BO':'mean'}).reset_index()

    # %%
    print(f'Dimension de tabla final: {tabla_final.shape}')
    print('-'*40)
    print(f'Total Campo OC: {tabla_final['Cant. UN. OC'].sum()}')
    print(f'Total Campo Factura: {tabla_final['Cant. Factura'].sum()}')
    print(f'Total Campo BO: {tabla_final['Cantidad en BO'].sum()}')

    # %%


    # %%
    df_ee = df_ee[['Documento compras','Posición modelo', 'Cantidad entrega', 'Status mov.mcía.']]

    # %%
    df_ee['Documento compras'] = df_ee['Documento compras'].astype('str')
    df_ee['Posición modelo'] = df_ee['Posición modelo'].astype('str')

    # %%
    df_ee.dtypes

    # %%
    df_ee['AUX'] = df_ee['Documento compras'] + df_ee['Posición modelo']
    df_ee.drop(columns=['Documento compras','Posición modelo'])

    # %%
    table_ee = df_ee[df_ee['Status mov.mcía.']=='C'].groupby(['AUX'])['Cantidad entrega'].sum().reset_index()

    # %%
    me2l_15 = me2l_base_tr

    # %%
    monitor_df.drop_duplicates(subset=['AUX'], inplace=True)

    # %%
    me2l_15 = me2l_base_tr

    # %%
    me2l_15 = me2l_15.merge(monitor_df[['AUX','Fe. ATA']], left_on='AUX', right_on='AUX', how = 'left')

    # %%
    me2l_15['Fecha Est. Fact'] = np.where(me2l_15['Fe. ATA'].notnull(), me2l_15['Fe. ATA'] + pd.Timedelta(days=15), me2l_15['Fecha de entrega'])

    # %%
    me2l_15 = me2l_15.merge(tabla_final[['AUX', 'Cant. Factura']], left_on = 'AUX', right_on='AUX',how='left')

    # %%
    me2l_15['Cant. Factura'] = me2l_15['Cant. Factura'] * me2l_15['Corresponde']

    # %%
    me2l_15= me2l_15.merge(table_ee, left_on='AUX', right_on='AUX', how='left')

    # %%
    me2l_15['Cantidad entrega'] = me2l_15['Cantidad entrega']*me2l_15['Corresponde']

    # %%
    me2l_15['Qty por Entregar'] = me2l_15['Por entregar (cantidad)']*me2l_15['Corresponde']

    # %%
    #cambiar esto por la fecha del dia actual

    hoy = datetime.datetime.today()

    # %%


    # %%
    me2l_15['Cantidad entrega'].fillna(0, inplace=True)

    # %%


    # Suponiendo que 'me2l_15' es tu DataFrame y 'AW4', 'BJ4' y 'AT4' son columnas en tu DataFrame
    result = me2l_15['Cant. Factura'] - me2l_15['Cantidad entrega']
    result = np.where(np.isnan(result), me2l_15['Qty por Entregar'], result)

    me2l_15['Qty Fact Corr'] = np.where(result<0, 0, result)

    # %%
    hoy = pd.to_datetime(hoy) 

    # %%
    me2l_15[ 'Fecha Est. Fact'].max()

    # %%
    conditions = [
        (me2l_15['Qty Fact Corr'] > 0) & (me2l_15['Fecha Est. Fact'] >= hoy),
        (me2l_15['Qty Fact Corr'] > 0) & (me2l_15['Fecha Est. Fact'] < hoy)
    ]

    choices = ['Facturado No Vencido', 'Facturado Vencido']

    me2l_15['Status TR Fact'] = np.select(conditions, choices, default='No Facturado')

    # %%
    from datetime import timedelta, datetime

    hoy_31 = hoy  + timedelta(days=31)
    hoy_7 = hoy  + timedelta(days=7)

    # %%
    hoy_31 = pd.to_datetime(hoy_31)
    hoy_7 = pd.to_datetime(hoy_7)


    # %%
    import numpy as np

    # Ventanas dinámicas según grupo
    limite_superior = np.where(
        me2l_15['Grupo de compras'] == 'RR2',
        hoy_7,     # hoy + 7 días
        hoy_31    # hoy + 31 días
    )

    # Condición final
    condition = (
        (me2l_15['Fecha de entrega'] > hoy) &
        (me2l_15['Fecha de entrega'] < limite_superior)
    )

    me2l_15['Fecha BO Mes Actual'] = np.where(condition, 'SI', 'NO')

    # %%
    limite_inferior = np.where(
        me2l_15['Grupo de compras'] == 'RR2',
        hoy - pd.Timedelta(days=7),
        hoy - pd.Timedelta(days=31)
    )

    condition = me2l_15['Fecha de entrega'] < limite_inferior

    me2l_15['TR Vencido'] = np.where(condition, 'SI', 'NO')


    # %%
    #HAcer una condicional aparte para cuando el grupo de articulo sea RR2
    #Si fecha de entrega es mayor a hoy +7, Fecha teorica
    #si es mayor a hoy pero no alcanzan a ser 7 dias Replanificar mas 7 

    # %%
    conditions = [
        (me2l_15['Fecha BO Mes Actual'] == "NO") & (me2l_15['TR Vencido'] == "NO"),
        (me2l_15['TR Vencido'] == "SI"),
    ]



    # Define choices
    choices = [
        "Fecha Teórica",
        "TR Vencido",
    ]

    # Apply np.select
    me2l_15['Statur TR No FAct'] = np.select(conditions, choices, default="Replanificar +45")

    # %%
    #Qty BO
    me2l_15['Cant. BO'] = me2l_15['Cant. Factura'].fillna(0, inplace=True)

    # %%
    #Cant Factura : Cant Fact CORR
    #QTY BO: CHECK

    # %%
    condicion = (me2l_15['Fecha Est. Fact'] < hoy)

    me2l_15['Atraso'] = np.where(condicion, 'SI','NO')

    # %%
    condicion = (me2l_15['Atraso'] == 'SI')

    me2l_15['Fecha Facturacion Final'] = np.where(condicion, hoy_31, me2l_15['Fecha Est. Fact'])



    # %%
    me2l_15['Fecha Facturacion Final'] = pd.to_datetime(me2l_15['Fecha Facturacion Final'])

    # %%
    me2l_15 = me2l_15.merge(monitor_df[['AUX', 'Vía (Texto)']], left_on='AUX', right_on='AUX', how='left')

    # %%
    me2l_15['LT Objetivo'] = np.where(me2l_15['Vía (Texto)'].isin(["Marítimo", "Terrestre"]), 45,
                        np.where(me2l_15['Vía (Texto)'] == "Aéreo", 15,
                            np.where(me2l_15['Vía (Texto)'] == "Courier", 7, 0)))

    # %%
    me2l_15 = me2l_15.merge(monitor_df[['AUX','Fe. Confirmación']], left_on='AUX',right_on='AUX', how='left')

    # %%
    me2l_15['Fecha OC'] = np.where((me2l_15['Fe. Confirmación'] == 0) | (me2l_15['Fe. Confirmación'].isna()), me2l_15['Fecha documento'], me2l_15['Fe. Confirmación'])

    # %%
    import datetime
    today = hoy

    # %%
    today = pd.to_datetime(today)

    # %%
    me2l_15['Fecha OC'] = pd.to_datetime(me2l_15['Fecha OC'], errors='coerce')

    # Calculate the difference
    me2l_15['Dias LT BO'] = (today - me2l_15['Fecha OC']).dt.days - me2l_15['LT Objetivo']

    # Handling potential errors, setting them to 0
    me2l_15['Dias LT BO'] = me2l_15['Dias LT BO'].fillna(0)

    # %%
    import pandas as pd
    import numpy as np

    # Define conditions and choices
    conditions = [
        (me2l_15['Vía (Texto)'].isin(["Marítimo", "Terrestre"])) & (me2l_15['Dias LT BO'].between(1, 30)),
        (me2l_15['Vía (Texto)'].isin(["Marítimo", "Terrestre"])) & (me2l_15['Dias LT BO'].between(31, 60)),
        (me2l_15['Vía (Texto)'].isin(["Marítimo", "Terrestre"])) & (me2l_15['Dias LT BO'].between(61, 90)),
        (me2l_15['Vía (Texto)'].isin(["Marítimo", "Terrestre"])) & (me2l_15['Dias LT BO'] > 90),
        (me2l_15['Vía (Texto)'] == "Aéreo") & (me2l_15['Dias LT BO'].between(1, 15)),
        (me2l_15['Vía (Texto)'] == "Aéreo") & (me2l_15['Dias LT BO'].between(16, 30)),
        (me2l_15['Vía (Texto)'] == "Aéreo") & (me2l_15['Dias LT BO'].between(31, 45)),
        (me2l_15['Vía (Texto)'] == "Courier") & (me2l_15['Dias LT BO'].between(1, 7)),
        (me2l_15['Vía (Texto)'] == "Courier") & (me2l_15['Dias LT BO'].between(8, 15)),
        (me2l_15['Vía (Texto)'] == "Courier") & (me2l_15['Dias LT BO'].between(16, 21)),
        (me2l_15['Vía (Texto)'] == "Courier") & (me2l_15['Dias LT BO'] > 21),
        (me2l_15['Vía (Texto)'] == "Aéreo") & (me2l_15['Dias LT BO'] > 45)
    ]

    choices = [
        "De 1 a 30 días",
        "De 31 a 60 días",
        "De 61 a 90 días",
        "De 91 días a más",
        "De 1 a 15 días",
        "De 16 a 30 días",
        "De 31 a 45 días",
        "De 1 a 7 días",
        "De 8 a 15 días",
        "De 16 a 21 días",
        "De 21 días a más",
        "De 45 días a más"
    ]

    # Assign the values based on conditions
    me2l_15['Categoría LT'] = np.select(conditions, choices, default="Dentro de LT")



    # %%
    import pandas as pd
    import numpy as np

    # As
    BV1 = 45  

    #Definir condiciones
    conditions = [
        me2l_15['Statur TR No FAct'] == "Replanificar +45",
        me2l_15['Statur TR No FAct'] == "Fecha Teórica",
        (me2l_15['Statur TR No FAct'] == "TR Vencido") & ((me2l_15['Categoría LT'] == 0) | (me2l_15['Categoría LT'] == "Dentro de LT")),
        (me2l_15['Statur TR No FAct'] == "TR Vencido") & (me2l_15['Categoría LT'] == "De 31 a 60 días"),
        (me2l_15['Statur TR No FAct'] == "TR Vencido") & (me2l_15['Categoría LT'] == "De 61 a 90 días")
    ]

    choices = [
        hoy + timedelta(days =45),
        me2l_15['Fecha Facturacion Final'],
        hoy + timedelta(days =30),
        hoy + timedelta(days =60),
        hoy + timedelta(days =90)
    ]

    # Asignar valores segun condicion
    me2l_15['Fecha No Fact'] = np.select(conditions, choices, default=hoy + timedelta(days =120))


    # %%
    me2l_15['Fecha No Fact'] = pd.to_datetime(me2l_15['Fecha No Fact'])

    # %%
    me2l_15= me2l_15.merge(tabla_final[['AUX','Cantidad en BO']], left_on='AUX',right_on='AUX', how='left')

    # %% [markdown]
    # AQUI SE DEJA DE TOMAR MONITOR Y SE TOMA ME2L (POR CALCULAR (CANTIDAD))

    # %%
    me2l_15['Cantidad en BO'] = me2l_15['Por calcular (cantidad)']*me2l_15['Corresponde']

    # %%


    # %%
    me2l_15['Cantidad en BO'].fillna(0, inplace=True)

    # %%
    import numpy as np

    # Crear una condición basada en la columna 'TIPO'
    condicion = (me2l_15['TIPO'] == 'OEM') & (me2l_15['Por calcular (cantidad)'] == 0)

    # Actualizar la columna 'Cantidad en BO' basada en la condición
    me2l_15['Cantidad en BO'] = np.where(condicion, 0, me2l_15['Cantidad en BO'])



    # %%
    me2l_15.to_excel(destino_tubo + "Base Transito para Analisis.xlsx")

    # %%
    side_a = me2l_15[['AUX','Status TR Fact',   'Material','Texto breve','Qty Fact Corr','Centro','NomSector_actual','Origen','TIPO','Cl.documento compras','Fecha Facturacion Final','Nombre del proveedor','Grupo de compras']]
    side_b = me2l_15[['AUX','Statur TR No FAct','Material','Texto breve','Cantidad en BO',     'Centro','NomSector_actual','Origen','TIPO','Cl.documento compras','Fecha No Fact',          'Nombre del proveedor','Grupo de compras']]

    # %%
    side_a.rename(columns={'Status TR Fact':'Status','Qty Fact Corr':'Cantidad','Fecha Facturacion Final':'Fecha'}, inplace=True)
    side_b.rename(columns={'Statur TR No FAct':'Status','Cantidad en BO':'Cantidad','Fecha No Fact':'Fecha'}, inplace=True)

    # %%
    tr_final = pd.concat([side_a, side_b])

    # %% [markdown]
    # 

    # %%
    tr_final.to_excel(destino_tubo + "TR FINAL R3.xlsx", index=False)

    # %%
    me5a = pd.read_excel(destino_tubo + "ME5A R3.xlsx", dtype={"Solicitud de pedido":"str","Pos.solicitud pedido":"str", "Fecha de entrega":"str", "Urgencia necesidad":"str"})

    me5a['AUX'] = me5a['Solicitud de pedido'] + me5a['Pos.solicitud pedido']
    me5a['Fecha Estimada'] = pd.to_datetime(me5a['Fecha de entrega'].str[:4]+ '-' + me5a['Fecha de entrega'].str[4:6] + '-' + me5a['Fecha de entrega'].str[6:])

    hoy = datetime.datetime.today()
    hoy = pd.to_datetime(hoy)
    me5a['Atraso'] = np.where(hoy >= me5a['Fecha Estimada'], "SI", "NO")
    me5a['hoy']=pd.to_datetime(hoy)
    me5a['Fecha de llegada'] = me5a.apply(lambda row: row['hoy'] + datetime.timedelta(days=21) if row['Atraso']=="SI" and row['Urgencia necesidad'] in ['1','2'] else row['hoy'] + datetime.timedelta(days=120) if row['Atraso']=="SI" and row['Urgencia necesidad'] == '3' else row['Fecha Estimada'], axis=1)
    me5a['Fecha de llegada'] = pd.to_datetime(me5a['Fecha de llegada'])
    me5a['Status solped'] = np.where(me5a['Fecha Estimada'] == "SI", "Solped Vencida", "Solped en Curso")
    me5a['Semana'] = me5a['Fecha de llegada'].dt.isocalendar().week
    me5a['Año'] = me5a['Fecha de llegada'].dt.isocalendar().year
    me5a['Mes'] = me5a['Fecha de llegada'].dt.month

    # %%
    df_mara.drop_duplicates(subset='Material_R3', inplace=True)

    # %%
    df_mara.shape

    # %%
    me5a = me5a.merge(df_mara[['Material_R3','Sector_dsc']], left_on='Material', right_on='Material_R3', how="left")

    # %%
    me5a = me5a.rename(columns={'Status solped':'Status', 'Cantidad solicitada':'Cantidad', 'Sector_dsc':'NomSector_actual',  'Clase documento':'Cl.documento compras', 'Fecha de llegada':'Fecha'})

    me5a = me5a[['AUX', 'Status', 'Material', 'Texto breve', 'Cantidad', 'Centro',
        'NomSector_actual' ,'Cl.documento compras', 'Fecha',
        'Nombre del proveedor', 'Grupo de compras']]

    # %%
    tr_consolidado = pd.concat([tr_final, me5a])

    # %%
    tr_consolidado.to_excel(destino_tubo + 'TR FINAL R3 Consolidado.xlsx')

    # %%
    import win32com.client
    import getpass
    usuario = getpass.getuser()

    # Initialize SAP GUI Scripting
    sap_gui_auto = win32com.client.GetObject("SAPGUI")
    application = sap_gui_auto.GetScriptingEngine

    # Establish connection and session
    connection = application.Children(0)
    session = connection.Children(0)

    # Maximize the window
    session.findById("wnd[0]").maximize()

    # Enter transaction code
    session.findById("wnd[0]/tbar[0]/okcd").text = "MB52"
    session.findById("wnd[0]").sendVKey(0)

    # Open value help for plant field and fill in the values
    session.findById("wnd[0]/usr/btn%_WERKS_%_APP_%-VALU_PUSH").press()
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "0270"
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "0501"
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").text = "0201"
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").setFocus()
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").caretPosition = 4
    session.findById("wnd[1]").sendVKey(8)

    # Set the variant and execute
    session.findById("wnd[0]/usr/ctxtP_VARI").text = "/LRAVLIC"
    session.findById("wnd[0]/usr/ctxtP_VARI").setFocus()
    session.findById("wnd[0]/usr/ctxtP_VARI").caretPosition = 8
    session.findById("wnd[0]").sendVKey(8)

    # Export the data
    session.findById("wnd[0]/tbar[1]/btn[43]").press()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:/Users/{usuario}/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras KPI-Reportes/Bases Transito (python)"
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "STOCK_R3.XLSX"
    session.findById("wnd[1]/usr/ctxtDY_PATH").setFocus()
    session.findById("wnd[1]/usr/ctxtDY_PATH").caretPosition = 0
    session.findById("wnd[1]/tbar[0]/btn[11]").press()

    # Close the windows
    session.findById("wnd[0]").sendVKey(3)
    session.findById("wnd[0]").sendVKey(3)



    # %%
    import tkinter as tk
    from tkinter import messagebox

    root = tk.Tk()
    root.withdraw()  # Oculta la ventana principal
    messagebox.showinfo("Cambio de Sesión", "Por favor, haga el cambio de sesión antes de proseguir.")
    root.destroy()  # Destruye la ventana principal después de mostrar el mensaje

    # %%
    import win32com.client
    import time

    def main():
        try:
            SapGuiAuto = win32com.client.GetObject("SAPGUI")
            application = SapGuiAuto.GetScriptingEngine
            connection = application.Children(0)
            session = connection.Children(0)
            session.findById("wnd[0]").maximize()
            session.findById("wnd[0]/tbar[0]/okcd").text = "mb52"
            session.findById("wnd[0]").sendVKey(0)
            
            session.findById("wnd[0]/usr/btn%_WERKS_%_APP_%-VALU_PUSH").press()
            session.findById("wnd[1]/tbar[0]/btn[16]").press()
            
            session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "1335"
            session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "1305"
            session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "1344"
            session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").text = "3261"
            
            session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").setFocus()
            session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").caretPosition = 4
            session.findById("wnd[1]").sendVKey(8)
            
            session.findById("wnd[0]/usr/ctxtP_VARI").text = "/STOCK"
            session.findById("wnd[0]/usr/ctxtP_VARI").setFocus()
            session.findById("wnd[0]/usr/ctxtP_VARI").caretPosition = 6
            session.findById("wnd[0]").sendVKey(8)
            
            session.findById("wnd[0]/tbar[1]/btn[43]").press()
            session.findById("wnd[1]/tbar[0]/btn[0]").press()
            
            session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\\Users\\{usuario}\\Inchcape\\Planificación y Compras Chile - Documentos\\Planificación y Compras KPI-Reportes\\Bases Transito (python)"
            session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "STOCK_INP.XLSX"
            session.findById("wnd[1]/usr/ctxtDY_PATH").setFocus()
            session.findById("wnd[1]/usr/ctxtDY_PATH").caretPosition = 0
            session.findById("wnd[1]/tbar[0]/btn[11]").press()

            # Close the windows
            session.findById("wnd[0]").sendVKey(3)
            session.findById("wnd[0]").sendVKey(3)
            
        except Exception as e:
            print(f"An error occurred: {e}")

    if __name__ == "__main__":
        main()


    # %%
    import pandas as pd

    # %%
    stock_r3 = pd.read_excel("C:/Users/{usuario}/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras KPI-Reportes/Bases Transito (python)/STOCK_R3.XLSX",dtype={'Centro':'str'})
    stock_inp = pd.read_excel("C:/Users/{usuario}/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras KPI-Reportes/Bases Transito (python)/STOCK_INP.XLSX", dtype={'Almacén':'str','Centro':'str'})

    # %%
    stock_inp.rename(columns={'Número de material externo largo':'Ult. Eslabon'},inplace=True)

    # %%
    stock_inp = stock_inp[['Ult. Eslabon','Libre utilización', 'Centro', 'Almacén', 'Texto breve de material']]

    # %%
    stock_inp = stock_inp[~stock_inp['Almacén'].isin(['1030','1130'])]

    # %%
    stock_r3 = stock_r3.merge(cod_actual_df, left_on = 'Material', right_on='Nro_pieza_fabricante_1', how = 'left')

    # %%


    # %%
    stock_r3.drop(columns='Nro_pieza_fabricante_1', inplace=True)
    stock_r3.rename(columns={'Cod_Actual_1':'Ult. Eslabon'}, inplace=True)

    # %%


    # %%
    stock_r3['Ult. Eslabon'].fillna(stock_r3['Material'], inplace=True)

    # %%
    stock_r3 = stock_r3.merge(df_mara, left_on='Ult. Eslabon', right_on='Material_R3')

    # %%
    stock_final = pd.concat([stock_r3, stock_inp])

    # %%
    stock_final_cruce = stock_final.merge(base[['Material','Familia','Subfamilia']], left_on='Ult. Eslabon', right_on='Material', how='left')

    # %%
    stock_final_cruce.drop(columns={'Material_y'}, inplace=True)
    stock_final_cruce.rename(columns={'Material_x':'Material'}, inplace=True)
    stock_final_cruce.drop(columns={'Familia_x','Subfamilia_x'}, inplace=True)
    stock_final_cruce.rename(columns={'Familia_y':'Familia','Subfamilia_y':'Subfamilia'}, inplace=True)

    # %%
    condicion = [stock_final_cruce['Centro']=='1305',stock_final_cruce['Centro']=='1335', stock_final_cruce['Centro']=='3261']

    seleccion = ['DFSK','Subaru','Harley Davidson']

    stock_final_cruce['Sector_dsc'] = np.select(condicion,seleccion,stock_final_cruce['Sector_dsc'])

    # %%
    stock_final_cruce['Centro'] = pd.to_numeric(stock_final_cruce['Centro'], errors='coerce').round().astype('Int64')

    # %%
    stock_final_cruce.to_excel(destino_tubo + "Stock R3.xlsx")

    # %%
    tr_consolidado = tr_consolidado[
        (tr_consolidado['Status'].isin(['Facturado No Vencido', 'Facturado Vencido'])) &
        (tr_consolidado['Cl.documento compras'].isin(['ZSTO', 'ZSPT', 'ZIPL', 'ZATT']))
    ]

    # %%

    df_din = tr_consolidado.loc[tr_consolidado.groupby("Material")["Fecha"].idxmin(), ["Material", "Fecha", "Cantidad"]]
    df_din['LT'] = df_din['Fecha'] - hoy


    # %%
    df_din['LT'] = df_din['LT'].dt.days

    # %%
    print(hoy.strftime('%d-%m-%Y'))

    # %%
    df_din.to_excel(f"C:/Users/{usuario}/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras OEM/Calculadora VOR/{hoy.strftime('%d-%m-%Y')} - Calculadora.xlsx")




if __name__ == "__main__":
    main()  