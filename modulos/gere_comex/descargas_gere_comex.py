# %%
def main():
    import pandas as pd
    import numpy as np
    import os
    import datetime
    import win32com.client
    import time
    import getpass
    from datetime import timedelta

    usuario = getpass.getuser()

    inicio = time.time()


    # %%

    # Define the directory path
    carpeta_fechas = f"C:/Users/{usuario}/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras KPI-Reportes/3.Actualización diaria fechas DT/Actualización Diaria fechas Dts R3 2026"

    # Extract today's date in the specified format
    hoy = datetime.datetime.today().strftime('%d-%m-%Y')
    columnas = ['Nro. DT', 'Vía (Texto)']

    # Function to find and read the file that contains today's date
    def find_file_with_today_date(date_str):
        for archivo in os.listdir(carpeta_fechas):
            if date_str in archivo:  # Check if today's date is in the file name
                ruta = os.path.join(carpeta_fechas, archivo)
                try:
                    df_fechas = pd.read_excel(ruta, sheet_name='Data', dtype={'Nro. DT': 'str'})
                    print(f"Archivo encontrado: {ruta}")
                    return df_fechas
                except Exception as e:
                    print(f"Error al leer el archivo: {ruta}. Detalles: {e}")
                    return None
        print(f"No se encontró un archivo que contenga la fecha {date_str}.")
        return None

    # Find the file with today's date
    df_fechas = find_file_with_today_date(hoy)

    # Continue with your further processing
    if df_fechas is not None:
        # Realiza las operaciones con df_fechas
        pass
    else:
        # Maneja el caso en el que no se encontró el archivo
        pass



    #df_fechas = pd.read_excel(f"C:/Users/{usuario}/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras KPI-Reportes/Actualización diaria fechas DT/Actualización Diaria fechas Dts R3/02-01-2025 Actualizacion fechas diaria  Dts OEM (002).xlsx", sheet_name='Data', dtype = {'Nro. DT':'str'})

    # %%
    filtro = df_fechas[df_fechas['Vía (Texto)'].isin(['Maritimo','Terrestre'])]['Nro. DT']

    # %%
    filtro.to_clipboard(index=False, header=False)



    # %%
    import win32com.client

    try:
        # Initialize SAP GUI scripting
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
        application = SapGuiAuto.GetScriptingEngine
        connection = application.Children(0)
        session = connection.Children(0)
    except Exception as e:
        print(f"Error obtaining SAP GUI session: {str(e)}")
        exit(1)

    try:
        # Maximize the window
        session.findById("wnd[0]").maximize()

        # Execute transaction
        session.findById("wnd[0]/tbar[0]/okcd").text = "ZMM_MONITOR_ORDEN_CL"
        session.findById("wnd[0]").sendVKey(0)

        # Navigate to specific tab
        session.findById("wnd[0]/usr/tabsTABSTRIP_TABB2/tabpMON_DT").select()
        session.findById("wnd[0]/usr/tabsTABSTRIP_TABB2/tabpMON_DT/ssub%_SUBSCREEN_TABB2:ZMM_MONITO3_SEGUIMIENTO_ORV4CL:1005/btn%_SO_TKNUM_%_APP_%-VALU_PUSH").press()

        # Handle pop-up windows
        session.findById("wnd[1]/tbar[0]/btn[24]").press()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
        session.findById("wnd[0]").sendVKey(8)

        # Load variant
        session.findById("wnd[0]/usr/cntlCONTROL_ALV_UPPER/shellcont/shell").pressToolbarContextButton("&MB_VARIANT")
        session.findById("wnd[0]/usr/cntlCONTROL_ALV_UPPER/shellcont/shell").selectContextMenuItem("&LOAD")

        # Select variant row
        session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").setCurrentCell(2, "TEXT")
        session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectedRows = "2"
        session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").clickCurrentCell()

        # Export to Excel
        session.findById("wnd[0]/usr/cntlCONTROL_ALV_UPPER/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")
        session.findById("wnd[0]/usr/cntlCONTROL_ALV_UPPER/shellcont/shell").selectContextMenuItem("&PC")
        session.findById("wnd[1]").close()

        # Handle export pop-up windows
        session.findById("wnd[0]/usr/cntlCONTROL_ALV_UPPER/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")
        session.findById("wnd[0]/usr/cntlCONTROL_ALV_UPPER/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")
        session.findById("wnd[0]/usr/cntlCONTROL_ALV_UPPER/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")
        session.findById("wnd[0]/usr/cntlCONTROL_ALV_UPPER/shellcont/shell").selectContextMenuItem("&XXL")
        session.findById("wnd[1]/tbar[0]/btn[0]").press()

        # Set file path and name
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = f"C:\\Users\\{usuario}\\Inchcape\\Planificación y Compras Chile - Documentos\\Planificación y Compras KPI-Reportes\\Descargas Automaticas\\Gerenciamiento Comex"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "dts.XLSX"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = len("dts.XLSX")
        session.findById("wnd[1]/tbar[0]/btn[11]").press()

        # Close SAP windows
        session.findById("wnd[0]").sendVKey(3)
        session.findById("wnd[0]").sendVKey(3)

    except Exception as e:
        print(f"Error during SAP GUI interaction: {str(e)}")











    # %%
    import xlwings as xw
    try:
        book = xw.Book(f"C:/Users/{usuario}/Codigos/automatizacion_gere_comex/dts.XLSX")
        book.close()
    except Exception as e:
        print(e)

    # %%
    import win32com.client

    try:
        # Initialize SAP GUI scripting
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
        application = SapGuiAuto.GetScriptingEngine
        if not application:
            raise Exception("Error obtaining SAP GUI application")

        connection = application.Children(0)
        if not connection:
            raise Exception("Error obtaining SAP GUI connection")

        session = connection.Children(0)
        if not session:
            raise Exception("Error obtaining SAP GUI session")

        # Connect to WScript if available
        try:
            WScript = win32com.client.Dispatch("WScript")
            WScript.ConnectObject(session, "on")
            WScript.ConnectObject(application, "on")
        except Exception as e:
            print(f"Error connecting to WScript: {str(e)}")

        # Maximize the window
        session.findById("wnd[0]").maximize()

        # Execute transaction
        session.findById("wnd[0]/tbar[0]/okcd").text = "zmm_seguim_comex_cl"
        session.findById("wnd[0]").sendVKey(0)

        # Press buttons
        session.findById("wnd[0]/usr/btn%_P_TKNUM_%_APP_%-VALU_PUSH").press()
        session.findById("wnd[1]/tbar[0]/btn[24]").press()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
        session.findById("wnd[0]").sendVKey(8)

        # Export to Excel
        session.findById("wnd[0]/usr/cntlALV_COMEX/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")
        session.findById("wnd[0]/usr/cntlALV_COMEX/shellcont/shell").selectContextMenuItem("&XXL")
        session.findById("wnd[1]/tbar[0]/btn[0]").press()

        # Set file path and name
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = f"C:\\Users\\{usuario}\\Inchcape\\Planificación y Compras Chile - Documentos\\Planificación y Compras KPI-Reportes\\Descargas Automaticas\\Gerenciamiento Comex"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "comex.XLSX"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 5
        session.findById("wnd[1]/tbar[0]/btn[11]").press()

        # Close SAP windows
        session.findById("wnd[0]").sendVKey(3)
        session.findById("wnd[0]").sendVKey(3)

    except Exception as e:
        print(f"Error during SAP GUI interaction: {str(e)}")

if __name__ == "__main__":
    main()


