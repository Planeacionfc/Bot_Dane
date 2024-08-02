import os
import pandas as pd
from openpyxl import load_workbook
from openpyxl import Workbook
import time
from sap_library.sap_scripting import Sap, Transaction, DatabaseProcessing, TransactionThread, Utils


# def procesar_archivo_ke30():
#     USER = os.getenv('DANIELA').split(",")
#     sap_instance = Sap(USER[0], USER[1], conection="PRD [PRODUCTIVO]") 
#     ke30 = Transaction(sap_instance.get_session(), "ke30" ) 
#     ke30.open_transaction()

#     # Configuración de filtros y operaciones en la transacción KE30
#     ke30.sap_session.findById("wnd[1]/usr/ctxtRKEA2-ERKRS").text = "PRBE"
#     ke30.sap_session.findById("wnd[1]/usr/ctxtRKEA2-ERKRS").caretPosition = 4
#     ke30.sap_session.findById("wnd[1]/tbar[0]/btn[0]").press()
#     ke30.sap_session.findById("wnd[0]/shellcont/shell").selectedNode = "000000001002"
#     ke30.sap_session.findById("wnd[0]/shellcont/shell").doubleClickNode("000000001002")
#     ke30.sap_session.findById("wnd[0]/usr/ctxtPAR_02").text = "006.2024"  # Cambiar fecha
#     ke30.sap_session.findById("wnd[0]/usr/ctxtPAR_02").caretPosition = 8
#     ke30.sap_session.findById("wnd[0]/tbar[1]/btn[8]").press()
#     ke30.sap_session.findById("wnd[0]/mbar/menu[2]/menu[9]").select()
#     ke30.sap_session.findById("wnd[0]/usr/tabsTABSTRIP/tabpTAB2/ssubSUB2:SAPLPIVB:1020/cntlCCONTROL_LAYOUT/shellcont/shell").selectedRows = "0"
#     ke30.sap_session.findById("wnd[0]/usr/tabsTABSTRIP/tabpTAB2/ssubSUB2:SAPLPIVB:1020/cntlCCONTROL_LAYOUT/shellcont/shell").clickCurrentCell()
#     ke30.sap_session.findById("wnd[0]/shellcont/shell").pressToolbarButton("SHOWBUT")
#     ke30.sap_session.findById("wnd[0]/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")
#     ke30.sap_session.findById("wnd[0]/shellcont/shell").selectContextMenuItem("&XXL")
#     ke30.sap_session.findById("wnd[1]/tbar[0]/btn[0]").press()
#     ke30.sap_session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\\Users\\prac.planeacionfi\\OneDrive - Prebel S.A BIC\\Escritorio\\Practi-SOFÍA\\bot_Dane"  # Ruta para guardar archivo
#     ke30.sap_session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "KE30.XLSX"  # Nombre archivo
#     ke30.sap_session.findById("wnd[1]/usr/ctxtDY_PATH").setFocus()
#     ke30.sap_session.findById("wnd[1]/usr/ctxtDY_PATH").caretPosition = 86
#     ke30.sap_session.findById("wnd[1]/tbar[0]/btn[0]").press()
#     ke30.sap_session.findById("wnd[0]").close()
#     ke30.sap_session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()

#     # Leer el archivo Excel
#     df = pd.read_excel('KE30.XLSX')

#     # Sumar las columnas M, F, G, H, J y restar I para cada fila
#     df['Ingreso Neto'] = df[['Ventas Brutas', 'Dctos Bonificaciones', 'Descuento de ventas', 'Descuento mercadeo', 'Devoluciones']].sum(axis=1) - df['Descuento ProntoPago']

#     # Guardar el DataFrame modificado en un nuevo archivo Excel
#     df.to_excel('KE30.xlsx', index=False)

# # Llamar a la función para ejecutar el proceso
# procesar_archivo_ke30()


# def procesar_archivo_mb51():
#     USER = os.getenv('DANIELA').split(",")
#     sap_instance= Sap(USER[0],USER[1],conection="PRD [PRODUCTIVO]") 

#     Mb51=Transaction(sap_instance.get_session(),"Mb51" ) 
#     Mb51.open_transaction()



#     Mb51.sap_session.findById("wnd[0]/usr/ctxtWERKS-LOW").text = "1000"
#     Mb51.sap_session.findById("wnd[0]/usr/ctxtBWART-LOW").setFocus()
#     Mb51.sap_session.findById("wnd[0]/usr/ctxtBWART-LOW").caretPosition = 0
#     Mb51.sap_session.findById("wnd[0]/usr/btn%_BWART_%_APP_%-VALU_PUSH").press()
#     Mb51.sap_session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "102"
#     Mb51.sap_session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "122"
#     Mb51.sap_session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").text = "123"
#     Mb51.sap_session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,4]").text = "101"
#     Mb51.sap_session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,4]").setFocus()
#     Mb51.sap_session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,4]").caretPosition = 3
#     Mb51.sap_session.findById("wnd[1]/tbar[0]/btn[8]").press()
#     Mb51.sap_session.findById("wnd[0]/usr/ctxtBUDAT-LOW").text = "01.06.2024" #cambiar intervalo de fecha
#     Mb51.sap_session.findById("wnd[0]/usr/ctxtBUDAT-HIGH").text = "31.06.2024"
#     Mb51.sap_session.findById("wnd[0]/usr/ctxtBUDAT-HIGH").setFocus()
#     Mb51.sap_session.findById("wnd[0]/usr/ctxtBUDAT-HIGH").caretPosition = 10
#     Mb51.sap_session.findById("wnd[0]/tbar[1]/btn[8]").press()

#     Mb51.sap_session.findById("wnd[0]/tbar[1]/btn[48]").press()
#     Mb51.sap_session.findById("wnd[0]/mbar/menu[3]/menu[2]/menu[1]").select()
#     Mb51.sap_session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").currentCellRow = 23
#     Mb51.sap_session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").firstVisibleRow = 23
#     Mb51.sap_session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectedRows = "23"
#     Mb51.sap_session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").clickCurrentCell()

#     Mb51.sap_session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[1]").select()
#     Mb51.sap_session.findById("wnd[1]/tbar[0]/btn[0]").press()
#     Mb51.sap_session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\\Users\\prac.planeacionfi\\OneDrive - Prebel S.A BIC\\Escritorio\\Practi-SOFÍA\\bot_Dane"
#     Mb51.sap_session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "MB51.xlsx" #Nombre archivo
#     Mb51.sap_session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 4
#     Mb51.sap_session.findById("wnd[1]/tbar[0]/btn[0]").press()

#         # Carga el archivo Excel
#     df = pd.read_excel('MB51.XLSX')

#     # Elimina filas con valores nulos en la columna 'orden'
#     df = df.dropna(subset=['Material'])

#     # Elimina todas las filas donde la columna 'orden' contiene un asterisco (*)
#     df = df[~(df['Orden'].astype(str).str.contains('\*') | df['Orden'].astype(str).str.startswith('100'))]

#     # Guarda el DataFrame modificado en un nuevo archivo Excel
#     df.to_excel('MB51.xlsx', index=False)


#     Mb51.sap_session.findById("wnd[0]/tbar[0]/btn[15]").press()
#     Mb51.sap_session.findById("wnd[0]/tbar[0]/btn[15]").press()
#     Mb51.sap_session.findById("wnd[0]/tbar[0]/btn[15]").press()

# procesar_archivo_mb51()

def procesar_archivo_mc9():

    USER = os.getenv('DANIELA').split(",")
    sap_instance= Sap(USER[0],USER[1],conection="PRD [PRODUCTIVO]") 

    MC9=Transaction(sap_instance.get_session(),"Mc.9" ) 
    MC9.open_transaction()

    # CENTRO 1000

    #CATEGORIA 1001
    MC9.sap_session.findById("wnd[0]/usr/ctxtSL_WERKS-LOW").text = "1000"
    MC9.sap_session.findById("wnd[0]/usr/ctxtSL_BKLAS-LOW").text = "1001"
    MC9.sap_session.findById("wnd[0]/usr/ctxtSL_SPMON-LOW").text = "06.2024" #Cambiar fecha
    MC9.sap_session.findById("wnd[0]/usr/ctxtSL_SPMON-HIGH").text = "06.2024"
    MC9.sap_session.findById("wnd[0]/usr/radML").select()
    MC9.sap_session.findById("wnd[0]/usr/radML").setFocus()
    MC9.sap_session.findById("wnd[0]/tbar[1]/btn[8]").press()

    MC9.sap_session.findById("wnd[0]/mbar/menu[6]/menu[5]/menu[2]/menu[2]").select()
    MC9.sap_session.findById("wnd[1]/tbar[0]/btn[0]").press()
    MC9.sap_session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\\Users\\prac.planeacionfi\\OneDrive - Prebel S.A BIC\\Escritorio\\Practi-SOFÍA\\bot_Dane"
    MC9.sap_session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "1000_1001.txt"
    MC9.sap_session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 4
    MC9.sap_session.findById("wnd[1]/tbar[0]/btn[0]").press()

    #CATEGORIA 1023
    MC9.sap_session.findById("wnd[0]/tbar[0]/btn[15]").press()
    MC9.sap_session.findById("wnd[1]/usr/btnSPOP-OPTION2").press()
    MC9.sap_session.findById("wnd[0]/usr/ctxtSL_BKLAS-LOW").text = "1023"
    MC9.sap_session.findById("wnd[0]/usr/ctxtSL_BKLAS-LOW").setFocus()
    MC9.sap_session.findById("wnd[0]/usr/ctxtSL_BKLAS-LOW").caretPosition = 4
    MC9.sap_session.findById("wnd[0]/tbar[1]/btn[8]").press()

    MC9.sap_session.findById("wnd[0]/mbar/menu[6]/menu[5]/menu[2]/menu[2]").select()
    MC9.sap_session.findById("wnd[1]/tbar[0]/btn[0]").press()
    MC9.sap_session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\\Users\\prac.planeacionfi\\OneDrive - Prebel S.A BIC\\Escritorio\\Practi-SOFÍA\\bot_Dane"
    MC9.sap_session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "1000_1023.txt"
    MC9.sap_session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 9
    MC9.sap_session.findById("wnd[1]/tbar[0]/btn[0]").press()

    #CATEGORIA 1003
    MC9.sap_session.findById("wnd[0]/tbar[0]/btn[15]").press()
    MC9.sap_session.findById("wnd[1]/usr/btnSPOP-OPTION2").press()
    MC9.sap_session.findById("wnd[0]/usr/ctxtSL_BKLAS-LOW").text = "1003"
    MC9.sap_session.findById("wnd[0]/usr/ctxtSL_BKLAS-LOW").setFocus()
    MC9.sap_session.findById("wnd[0]/usr/ctxtSL_BKLAS-LOW").caretPosition = 4
    MC9.sap_session.findById("wnd[0]/tbar[1]/btn[8]").press()

    MC9.sap_session.findById("wnd[0]/mbar/menu[6]/menu[5]/menu[2]/menu[2]").select()
    MC9.sap_session.findById("wnd[1]/tbar[0]/btn[0]").press()
    MC9.sap_session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\\Users\\prac.planeacionfi\\OneDrive - Prebel S.A BIC\\Escritorio\\Practi-SOFÍA\\bot_Dane"
    MC9.sap_session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "1000_1003.txt"
    MC9.sap_session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 13
    MC9.sap_session.findById("wnd[1]/tbar[0]/btn[0]").press()




    # CENTRO 3000

    #CATEGORIA 1001

    MC9.sap_session.findById("wnd[0]/tbar[0]/btn[15]").press()
    MC9.sap_session.findById("wnd[1]/usr/btnSPOP-OPTION2").press()
    MC9.sap_session.findById("wnd[0]/usr/ctxtSL_WERKS-LOW").text = "3000"
    MC9.sap_session.findById("wnd[0]/usr/ctxtSL_BKLAS-LOW").text = "1001"
    MC9.sap_session.findById("wnd[0]/usr/ctxtSL_BKLAS-LOW").setFocus()
    MC9.sap_session.findById("wnd[0]/usr/ctxtSL_BKLAS-LOW").caretPosition = 4
    MC9.sap_session.findById("wnd[0]/tbar[1]/btn[8]").press()

    MC9.sap_session.findById("wnd[0]/mbar/menu[6]/menu[5]/menu[2]/menu[2]").select()
    MC9.sap_session.findById("wnd[1]/tbar[0]/btn[0]").press()
    MC9.sap_session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\\Users\\prac.planeacionfi\\OneDrive - Prebel S.A BIC\\Escritorio\\Practi-SOFÍA\\bot_Dane"
    MC9.sap_session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "3000_1001.txt"
    MC9.sap_session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 13
    MC9.sap_session.findById("wnd[1]/tbar[0]/btn[0]").press()

    #CATEGORIA 1023

    MC9.sap_session.findById("wnd[0]/tbar[0]/btn[15]").press()
    MC9.sap_session.findById("wnd[1]/usr/btnSPOP-OPTION2").press()
    MC9.sap_session.findById("wnd[0]/usr/ctxtSL_BKLAS-LOW").text = "1023"
    MC9.sap_session.findById("wnd[0]/usr/ctxtSL_BKLAS-LOW").setFocus()
    MC9.sap_session.findById("wnd[0]/usr/ctxtSL_BKLAS-LOW").caretPosition = 4
    MC9.sap_session.findById("wnd[0]/tbar[1]/btn[8]").press()

    MC9.sap_session.findById("wnd[0]/mbar/menu[6]/menu[5]/menu[2]/menu[2]").select()
    MC9.sap_session.findById("wnd[1]/tbar[0]/btn[0]").press()
    MC9.sap_session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\\Users\\prac.planeacionfi\\OneDrive - Prebel S.A BIC\\Escritorio\\Practi-SOFÍA\\bot_Dane"
    MC9.sap_session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "3000_1023.txt"
    MC9.sap_session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 13
    MC9.sap_session.findById("wnd[1]/tbar[0]/btn[0]").press()

    #CATEGORIA 1003
    MC9.sap_session.findById("wnd[0]/tbar[0]/btn[15]").press()
    MC9.sap_session.findById("wnd[1]/usr/btnSPOP-OPTION2").press()
    MC9.sap_session.findById("wnd[0]/usr/ctxtSL_BKLAS-LOW").text = "1003"
    MC9.sap_session.findById("wnd[0]/usr/ctxtSL_BKLAS-LOW").setFocus()
    MC9.sap_session.findById("wnd[0]/usr/ctxtSL_BKLAS-LOW").caretPosition = 4
    MC9.sap_session.findById("wnd[0]/tbar[1]/btn[8]").press()

    MC9.sap_session.findById("wnd[0]/mbar/menu[6]/menu[5]/menu[2]/menu[2]").select()
    MC9.sap_session.findById("wnd[1]/tbar[0]/btn[0]").press()
    MC9.sap_session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\\Users\\prac.planeacionfi\\OneDrive - Prebel S.A BIC\\Escritorio\\Practi-SOFÍA\\bot_Dane"
    MC9.sap_session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "3000_1003.txt"
    MC9.sap_session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 13
    MC9.sap_session.findById("wnd[1]/tbar[0]/btn[0]").press()

    
    # Lista de nombres de archivos que deseas combinar
    archivos = ['1000_1001.txt', '1000_1003.txt', '1000_1023.txt', '3000_1001.txt', '3000_1003.txt', '3000_1023.txt']

    # Lista para almacenar DataFrames
    dfs = []

    # Iterar sobre los nombres de archivos
    for archivo in archivos:
        # Leer el archivo de texto línea por línea
        with open(archivo, 'r') as f:
            lineas = f.readlines()
        
        # Dividir cada línea por el delimitador '|' y almacenar los datos en una lista de listas
        datos = [linea.strip().split('|') for linea in lineas]
        
        # Crear un DataFrame a partir de los datos
        df = pd.DataFrame(datos)
        
        # Agregar el DataFrame a la lista de DataFrames
        dfs.append(df)

    # Concatenar todos los DataFrames en uno solo
    df_final = pd.concat(dfs, ignore_index=True)

    # Escribir el DataFrame en un archivo Excel
    df_final.to_excel('archivo_final_MC9.xlsx', index=False)


    MC9.sap_session.findById("wnd[0]/tbar[0]/btn[15]").press()
    MC9.sap_session.findById("wnd[1]/usr/btnSPOP-OPTION2").press()
    MC9.sap_session.findById("wnd[0]/tbar[0]/btn[15]").press()

procesar_archivo_mc9()

def procesar_archivo_me2n():
    USER = os.getenv('DANIELA').split(",")
    sap_instance= Sap(USER[0],USER[1],conection="PRD [PRODUCTIVO]") 

    ME2N=Transaction(sap_instance.get_session(),"ME2N" ) 
    ME2N.open_transaction()

    ME2N.sap_session.findById("wnd[0]/usr/btn%_S_LIFNR_%_APP_%-VALU_PUSH").press()
    ME2N.sap_session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "3000003340"  #Provedores
    ME2N.sap_session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "3000000491"
    ME2N.sap_session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "3000004007"
    ME2N.sap_session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").setFocus()
    ME2N.sap_session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").caretPosition = 10
    ME2N.sap_session.findById("wnd[1]/tbar[0]/btn[8]").press()
    ME2N.sap_session.findById("wnd[0]/usr/ctxtS_BEDAT-LOW").text = "01.06.2024"  #Cambiar intervalo de fechas
    ME2N.sap_session.findById("wnd[0]/usr/ctxtS_BEDAT-HIGH").text = "31.06.2024"
    ME2N.sap_session.findById("wnd[0]/usr/ctxtS_BEDAT-HIGH").setFocus()
    ME2N.sap_session.findById("wnd[0]/usr/ctxtS_BEDAT-HIGH").caretPosition = 8
    ME2N.sap_session.findById("wnd[0]/tbar[1]/btn[8]").press()

    ME2N.sap_session.findById("wnd[0]/tbar[1]/btn[33]").press()
    ME2N.sap_session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").currentCellRow = 1
    ME2N.sap_session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectedRows = "1"
    ME2N.sap_session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").clickCurrentCell()
    ME2N.sap_session.findById("wnd[0]/tbar[1]/btn[43]").press()
    ME2N.sap_session.findById("wnd[1]/tbar[0]/btn[0]").press()
    ME2N.sap_session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\\Users\\prac.planeacionfi\\OneDrive - Prebel S.A BIC\\Escritorio\\Practi-SOFÍA\\bot_Dane"
    ME2N.sap_session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "ME2N.XLSX"  #Nombre archivo
    ME2N.sap_session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 9
    ME2N.sap_session.findById("wnd[1]/tbar[0]/btn[0]").press()

    ME2N.sap_session.findById("wnd[0]/tbar[0]/btn[15]").press()
    ME2N.sap_session.findById("wnd[0]/tbar[0]/btn[15]").press()