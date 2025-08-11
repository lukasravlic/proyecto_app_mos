import webview
import subprocess
import os
import sys

#SCRIPT FORECAST
import fc_consolidado_axs
import fc_consolidado_mainstream
import fc_consolidado_premium
import union_fc

#SCRIPT VENTAS
import ventas_mainstream
import ventas_axs
import ventas_premium
import union_venta

import consolidacion_mara



class Api:
    # No guardes el window como atributo, solo métodos públicos

    # Método genérico para ejecutar scripts basado en el ID recibido del frontend
    def execute_script(self, script_id):
        # FC Mainstream
        if script_id == '1':
            output = fc_consolidado_mainstream.main()
            return output
        #FC AXS
        elif script_id == '2':
            output = fc_consolidado_axs.main()
            return output
        #FC Premium
        elif script_id == '3':
            output = fc_consolidado_premium.main()
            return output
        #Consolidacion FC 
        elif script_id == '4':
           output = union_fc.main()
           return output
        elif script_id == '5':
            output = ventas_mainstream.main()
            return output
        elif script_id == '6':
            output = ventas_axs.main()
            return output
        elif script_id == '7':
            output = ventas_premium.main()
            return output
        elif script_id == '8':
            output = union_venta.main()
            return output
        elif script_id == '9':
            output = consolidacion_mara.main()
            return output
        else:
            print(f"Script con ID '{script_id}' no reconocido. No se ejecutó nada.")
            return f"Error: Script con ID '{script_id}' no reconocido."

if __name__ == '__main__':
    if getattr(sys, 'frozen', False):
        BASE_DIR = sys._MEIPASS
    else:
        BASE_DIR = os.path.dirname(os.path.abspath(__file__))

    ruta_absoluta_html = os.path.join(BASE_DIR, "index.html")
    api = Api()
    window = webview.create_window("Panel de Scripts", ruta_absoluta_html, js_api=api)
    webview.start()