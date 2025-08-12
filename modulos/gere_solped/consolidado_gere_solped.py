def main():    

    import getpass
    import pandas as pd

    usuario = getpass.getuser()

    df_inp = pd.read_excel(f"C:/Users/{usuario}/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras KPI-Reportes/Descargas Automaticas/Gerenciamiento Solped/EXPORT.XLSX", sheet_name="Sheet1")
    df_r3 = pd.read_excel(f"C:/Users/{usuario}/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras KPI-Reportes/Descargas Automaticas/Gerenciamiento Solped/ME5A_R3.XLSX",sheet_name="Sheet1")

    # %%





    # %%
    df_inp.rename(columns={'Fecha orden compra':'Fecha de pedido'}, inplace=True)

    # %%
    df_inp = df_inp[[ 'Solicitud de pedido',
                    'Clase documento',
                    'Fecha de solicitud', 
                    'Pos.solicitud pedido', 
                    'Material', 
                    'Nº material proveedor',
                    'Texto breve', 
                    'Cantidad solicitada', 
                    'Unidad de medida',
                    'Nombre del proveedor',
                    'Indicador de borrado',
                    'Status tratamiento',
                    'Centro',
                    'Status tratamiento solicitud pedido',
                    'Fecha de entrega',
                    'Grupo de compras',
                    'Solicitante',
                    'Proveedor deseado',
                    'Proveedor fijo',
                    'Reg.info de compras',
                    'Creado por',
                    'Fecha de pedido',
                    'Nombre del proveedor deseado',
                    'Pedido',
                    'Posición de pedido',
                    'Proveedor',
                    'Moneda',
                    'Precio de valoración',
                    'Valor total',
                    'Cantidad pedida',
                    'Texto bloqueo',
                    'Cantidad confirmada']]
                    



    # %%
    me5a_consolidado = pd.concat([ df_r3, df_inp], ignore_index=True)

    #me5a_consolidado = me5a_consolidado[me5a_consolidado['Fecha de pedido'].dt.year > 2024]
    me5a_consolidado = me5a_consolidado[me5a_consolidado['Texto bloqueo'].isna()]

    # %%
    me5a_consolidado.to_excel(f"C:/Users/{usuario}/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras KPI-Reportes/Descargas Automaticas/Gerenciamiento Solped/gere_solped_consolidado.xlsx", sheet_name="Sheet1")

    print("🎊 Descargas de SAP consolidadas para ambos legacies! Archivo final exportado en ---> Planificación y Compras KPI-Reportes/Descargas Automaticas/Gerenciamiento Solped.")


if __name__ == '__main__':
    print(main())