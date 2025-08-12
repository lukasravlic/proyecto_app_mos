def main():    


    print("⏰Iniciando descarga de ME5A para INP...")

    import win32com.client
    import getpass
    usuario = getpass.getuser()

    try:
        # Initialize SAP GUI Scripting
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
        application = SapGuiAuto.GetScriptingEngine
        connection = application.Children(0)
        session = connection.Children(0)

        # Opcional: maximiza la ventana SAP
        session.findById("wnd[0]").maximize()

        # Transacción ME5A
        session.findById("wnd[0]/tbar[0]/okcd").text = "me5a"
        session.findById("wnd[0]").sendVKey(0)

        # Layout ALV
        session.findById("wnd[0]/usr/ctxtP_LSTUB").text = "alv"
        session.findById("wnd[0]/usr/ctxtP_LSTUB").setFocus()
        session.findById("wnd[0]/usr/ctxtP_LSTUB").caretPosition = 3
        session.findById("wnd[0]").sendVKey(0)

        # Selección de plantas 1305 a 1335
        session.findById("wnd[0]/usr/btn%_S_WERKS_%_APP_%-VALU_PUSH").press()
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "1305"
        session.findById("wnd[1]").sendVKey(0)
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "1335"
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").setFocus()
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").caretPosition = 4
        session.findById("wnd[1]").sendVKey(0)
        session.findById("wnd[1]").sendVKey(8)

        # Filtros S_BANPR
        session.findById("wnd[0]/usr/ctxtS_BANPR-LOW").setFocus()
        session.findById("wnd[0]/usr/ctxtS_BANPR-LOW").caretPosition = 0
        session.findById("wnd[0]/usr/btn%_S_BANPR_%_APP_%-VALU_PUSH").press()
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "a"
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "n"
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "b"
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").setFocus()
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").caretPosition = 1
        session.findById("wnd[1]").sendVKey(8)

        # Ejecutar reporte
        session.findById("wnd[0]").sendVKey(8)

        # Confirmar layout ALV (otra vez)
        session.findById("wnd[0]/usr/ctxtP_LSTUB").text = "Alv"
        session.findById("wnd[0]/usr/ctxtP_LSTUB").setFocus()
        session.findById("wnd[0]/usr/ctxtP_LSTUB").caretPosition = 3
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]").sendVKey(8)

        # Exportar reporte a Excel
        session.findById("wnd[0]/tbar[1]/btn[33]").press()  # Lista -> Exportar
        session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").currentCellColumn = "TEXT"
        session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").firstVisibleRow = 0
        session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectedRows = "0"
        session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").clickCurrentCell()

        # Confirmar exportar a archivo local
        session.findById("wnd[0]/tbar[1]/btn[43]").press()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = f"C:/Users/{usuario}/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras KPI-Reportes/Descargas Automaticas/Gerenciamiento Solped"
        session.findById("wnd[1]/usr/ctxtDY_PATH").setFocus()
        session.findById("wnd[1]/usr/ctxtDY_PATH").caretPosition = 0
        session.findById("wnd[1]/tbar[0]/btn[11]").press()

        # Cerrar ventanas
        session.findById("wnd[0]").sendVKey(3)
        session.findById("wnd[0]").sendVKey(3)
        print("🎊 Descarga y tratamiento de ME5A para INP finalizados correctamente. Archivo descargado en ---> Planificación y Compras KPI-Reportes/Descargas Automaticas/Gerenciamiento Solped.")

    except Exception as e:
        print(f"❌ Error: {str(e)}")

    return "\n".join("")

if __name__ == '__main__':
    print(main())