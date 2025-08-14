def main():    

    print("⏰Iniciando descarga de ME5A para VMP...")

    import win32com.client
    import getpass
    import pandas as pd
    import time
    usuario = getpass.getuser()

    try:
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
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = f"C:/Users/{usuario}/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras KPI-Reportes/Descargas Automaticas/Gerenciamiento Solped"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "ME5A_R3.XLSX"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 8
        session.findById("wnd[1]/tbar[0]/btn[11]").press()

        # Close the session
        session.findById("wnd[0]").sendVKey(3)
        session.findById("wnd[0]").sendVKey(3)

        print("✅ Extracción de SAP finalizada.")

        
        time.sleep(5)
        ruta = f"C:/Users/{usuario}/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras KPI-Reportes/Descargas Automaticas/Gerenciamiento Solped/ME5A_R3.XLSX"
        df_me5a = pd.read_excel(ruta)
        columnas = [
            "Solicitud de pedido",
            "Clase documento",
            "Fecha de solicitud",
            "Pos.solicitud pedido",
            "Material",
            "Nº material proveedor",
            "Texto breve",
            "Cantidad solicitada",
            "Unidad de medida",
            "Nombre del proveedor",
            "Indicador de borrado",
            "Status tratamiento",
            "Centro",
            "Status tratamiento solicitud pedido",
            "Fecha de entrega",
            "Grupo de compras",
            "Solicitante",
            "Proveedor deseado",
            "Proveedor fijo",
            "Reg.info de compras",
            "Creado por",
            "Fecha de pedido",
            "Nombre del proveedor deseado",
            "Pedido",
            "Posición de pedido",
            "Proveedor",
            "Moneda",
            "Precio de valoración",
            "Valor total",
            "Cantidad pedida",
            "Petición de oferta",
            "Fecha Petición de oferta",
            "Texto bloqueo",
            "Cantidad confirmada"
        ]
        df_me5a = df_me5a[columnas]
        df_me5a.to_excel(f"C:/Users/{usuario}/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras KPI-Reportes/Descargas Automaticas/Gerenciamiento Solped/ME5A_R3_tratada.xlsx")
        print("🎊 Descarga y tratamiento de ME5A para VMP finalizados correctamente. Archivo descargado en ---> Planificación y Compras KPI-Reportes/Descargas Automaticas/Gerenciamiento Solped.")

    except Exception as e:
        print(f"❌ Error: {str(e)}")

    return "\n".join("")

if __name__ == '__main__':
    print(main())