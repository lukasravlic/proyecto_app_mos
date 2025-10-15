import webview
import subprocess
import os
import sys

#SCRIPT FORECAST
import modulos.mos.forecast.fc_consolidado_axs as fc_consolidado_axs
import modulos.mos.forecast.fc_consolidado_mainstream as fc_consolidado_mainstream
import modulos.mos.forecast.fc_consolidado_premium as fc_consolidado_premium
import modulos.mos.consolidaciones.union_fc as union_fc

#SCRIPT VENTAS
import modulos.mos.venta.ventas_mainstream as ventas_mainstream
import modulos.mos.venta.ventas_axs as ventas_axs
import modulos.mos.venta.ventas_premium as ventas_premium
import modulos.mos.consolidaciones.union_venta as union_venta

#SCRIPTS GERE SOLPED
import modulos.gere_solped.gere_solped_vmp as gere_solped_vmp
import modulos.gere_solped.gere_solped_inp as gere_solped_inp
import modulos.gere_solped.consolidado_gere_solped as consolidado_solped

#SCRIPT MARA/MOS
import modulos.mos.mara.consolidacion_mara as consolidacion_mara

#SCRIPT DISPONIBILIDAD FUTURA
import modulos.dispo_futura.oem.dispo_futura_oem as dispo_futura_oem
import modulos.dispo_futura.axs.dispo_futura_axs as dispo_futura_axs








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
        elif script_id == "GS1":
            output = gere_solped_vmp.main()
            return output
        elif script_id == "GS2":
            output = gere_solped_inp.main()
            return output
        elif script_id == "GS3":
            output = consolidado_solped.main()
            return output
        elif script_id == "dfOem":
            output = dispo_futura_oem.main()
            return output
        elif script_id == "dfAxs":
            output = dispo_futura_axs.main()
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