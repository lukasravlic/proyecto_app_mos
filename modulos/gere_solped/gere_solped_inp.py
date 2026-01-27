def main():    


    print("⏰Iniciando descarga de ME5A para INP...")

    import win32com.client
    import getpass
    usuario = getpass.getuser()

    try:
        # Initialize SAP GUI Scripting
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
        application = SapGuiAuto.GetScriptingEngine
        
        # Get the first connection and the first session
        connection = application.Children(0)
        session = connection.Children(0)

        # Maximize Window and enter Transaction ME5A
        session.findById("wnd[0]").maximize()
        session.findById("wnd[0]/tbar[0]/okcd").text = "me5a"
        session.findById("wnd[0]").sendVKey(0)

        # Set Scope of List to 'ALV'
        session.findById("wnd[0]/usr/ctxtP_LSTUB").text = "alv"
        session.findById("wnd[0]").sendVKey(0)

        # Open Multiple Selection for Plants (S_WERKS)
        session.findById("wnd[0]/usr/btn%_S_WERKS_%_APP_%-VALU_PUSH").press()
        
        # Enter Plant 1305
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "1305"
        session.findById("wnd[1]").sendVKey(0)
        
        # Enter Plant 1335
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "1335"
        session.findById("wnd[1]").sendVKey(0)
        
        # Execute Multiple Selection (F8) and then Execute Report (F8)
        session.findById("wnd[1]").sendVKey(8)
        session.findById("wnd[0]").sendVKey(8)

        # Layout selection and Export
        session.findById("wnd[0]/tbar[1]/btn[33]").press()
        
        # Selecting the first layout row in the grid
        grid = session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell")
        grid.currentCellColumn = "TEXT"
        grid.selectedRows = "0"
        grid.clickCurrentCell()

        # Export to Local File
        session.findById("wnd[0]/tbar[1]/btn[43]").press()
        session.findById("wnd[1]/tbar[0]/btn[0]").press() # Confirm format
        
        # Set File Path (Ensure directory exists)
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = f"C:/Users/{usuario}/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras KPI-Reportes/Descargas Automaticas/Gerenciamiento Solped"
        session.findById("wnd[1]/tbar[0]/btn[11]").press() # Replace/Save

        # Go back to main menu
        session.findById("wnd[0]").sendVKey(3)
        session.findById("wnd[0]").sendVKey(3)
        print("🎊 Descarga y tratamiento de ME5A para INP finalizados correctamente. Archivo descargado en ---> Planificación y Compras KPI-Reportes/Descargas Automaticas/Gerenciamiento Solped.")

    except Exception as e:
        print(f"❌ Error: {str(e)}")

    return "\n".join("")

if __name__ == '__main__':
    print(main())