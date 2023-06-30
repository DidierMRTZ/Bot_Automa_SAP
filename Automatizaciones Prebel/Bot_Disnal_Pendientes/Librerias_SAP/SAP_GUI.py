# Librerias SAP GUI

# Importing the Libraries
import win32com.client
from datetime import datetime
import subprocess
from time import sleep
import os
import pywintypes
from pywinauto.application import Application
import pyautogui
#Iniciar sesión
# Input= Usuario y contraseña y output= session

password=None
def SessionSAP(user,password):
   path = "C:\\Program Files (x86)\\SAP\\FrontEnd\\SAPgui\\saplogon.exe"
   subprocess.Popen(path)
   sleep(3)
   SapGuiAuto = win32com.client.GetObject('SAPGUI')
   application = SapGuiAuto.GetScriptingEngine
   Connection = application.OpenConnection("PRD [PRODUCTIVO]", True)
   Session = Connection.Children(0)
   Session.findById("wnd[0]/usr/txtRSYST-BNAME").text = user
   Session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = password
   Session.findById("wnd[0]/tbar[0]/btn[0]").press()
   #Aqui es por si aparece una ventana adicional
   try:
      Session.findById("wnd[1]/usr/radMULTI_LOGON_OPT2").select()
      Session.findById("wnd[1]/usr/radMULTI_LOGON_OPT2").setFocus()
      Session.findById("wnd[1]/tbar[0]/btn[0]").press()
      return Session
   except:
      return Session

#Buscar transaccion LX03 general datos de entrada Transaccion y Session

def Search_table_Variant(session,Variant):
    session.findById("wnd[0]/tbar[1]/btn[17]").press()
    variants =session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell")
    row=variants.RowCount  #Count total rows
    lis=[variants.GetCellValue(i,"VARIANT") for i in range(0,row)]  #send variant apply GetCellValue
    indice=[indice for indice, dato in enumerate(lis) if dato == Variant]
    variants.selectedRows = indice[0]
    session.findById("wnd[1]/tbar[0]/btn[2]").press()
    return(Variant)


def Search_COGI(Transsaccion,Variant,session):
    session.StartTransaction(Transsaccion)
    session.findById("wnd[0]/tbar[1]/btn[17]").press()
    Search_table_Variant(session,Variant)
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    try:
        table=session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell")
        return(table)
    except:
        None

#Search can´t variant
def Search(Transsacion,session,provision):
    session.StartTransaction(Transsacion)
    variant=Search_table_Variant(session,provision)   #Recived variant
    session.findById("wnd[0]/tbar[1]/btn[8]").press()

def Search_MB52(Transsaccion,session,provision,variant):
    session.StartTransaction(Transsaccion)
    session.findById("wnd[0]/tbar[1]/btn[17]").press()
    session.findById("wnd[1]/usr/txtV-LOW").text = provision
    session.findById("wnd[1]/usr/txtENAME-LOW").text = variant
    session.findById("wnd[1]/usr/txtV-LOW").caretPosition = 8
    session.findById("wnd[1]/tbar[0]/btn[8]").press()
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    return(session)

def Search_COOISPI(Transsaccion,session,provision,variant,disposicion,DateIni=None,DateFin=None):
    session.StartTransaction(Transsaccion)
    session.findById("wnd[0]/tbar[1]/btn[17]").press()
    session.findById("wnd[1]/usr/txtV-LOW").text = variant
    session.findById("wnd[1]/usr/txtENAME-LOW").text = provision
    session.findById("wnd[1]/tbar[0]/btn[8]").press()
    session.findById("wnd[0]/usr/ssub%_SUBSCREEN_TOPBLOCK:PPIO_ENTRY:1100/ctxtPPIO_ENTRY_SC1100-ALV_VARIANT").text = disposicion
    if DateIni!=None and DateFin!=None:
        session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_ECKST-LOW").text = DateIni
        session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_ECKST-HIGH").text = DateFin
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        return(session)
    else:
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        return(session)


def Search_LX03(Transsaccion,session):
    session.StartTransaction(Transsaccion)
    session.findById("wnd[0]/usr/ctxtS1_LGNUM").text = "pro"
    session.findById("wnd[0]/usr/ctxtS1_LGNUM").caretPosition = 3
    session.findById("wnd[0]/tbar[1]/btn[17]").press()
    session.findById("wnd[1]/tbar[0]/btn[8]").press()
    session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "0"
    session.findById("wnd[1]/tbar[0]/btn[2]").press()
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    return(session)


def Search_ZPP57(Transsaccion,session):
    session.StartTransaction(Transsaccion)
    session.findById("wnd[0]/usr/ctxtSP$00001-LOW").text = "1000"
    session.findById("wnd[0]/usr/btn%_SP$00003_%_APP_%-VALU_PUSH").press()        #Boton para pasar los componentes
    session.findById("wnd[1]/tbar[0]/btn[24]").press()                            #Pegar materiales
    session.findById("wnd[1]/tbar[0]/btn[8]").press()
    session.findById("wnd[0]/tbar[1]/btn[8]").press()


#Search CO60

def Search_CO60(session,variant,orden=None):
    Ordenes=[] #Arreglar
    Info=[]  #Arreglar
    session.StartTransaction("CO60") 
    session.findById("wnd[0]/tbar[1]/btn[17]").press()
    variants =session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell")
    row=variants.RowCount  #Count total rows
    lis=[variants.GetCellValue(i,"VARIANT") for i in range(0,row)] 
    indice=[indice for indice, dato in enumerate(lis) if dato == variant]
    variants.selectedRows = indice[0]
    session.findById("wnd[1]/tbar[0]/btn[2]").press()
    if orden==None:
        None
    else:
        try:
            session.findById("wnd[0]/usr/ctxtS_AUFNR-LOW").text = orden
            session.findById("wnd[0]/tbar[1]/btn[8]").press()
        except:
            None
    session.findById("wnd[0]/tbar[1]/btn[5]").press()
    session.findById("wnd[0]/usr/shellcont/shell")
    pyautogui.click()
    try:
        app = Application().connect(title="Process Manufacturing Cockpit VVALENCIAO")  #Puede cambiar la ventana
        dlg = app.top_window()
        dlg.child_window(title="Control  Container", class_name="Shell Window Class").click_input()
        pyautogui.press('tab')
        pyautogui.press('enter')
        # Obtener las dimensiones de la pantalla
        screen_width, screen_height = pyautogui.size()
        # Mover el cursor al centro de la pantalla
        pyautogui.moveTo(screen_width/2, screen_height/2)
        # Desplazarse hacia abajo utilizando la función scroll
        pyautogui.scroll(-30600)
        sleep(1)
        dlg.child_window(title="Control  Container", class_name="Shell Window Class").click_input()
        pyautogui.press('tab')
        pyautogui.press('tab')
        try:
            pyautogui.write("VVALENCIAO")
            pyautogui.press('enter')
            session.findById("wnd[1]/usr/pwdSIGN_POPUP_STRUC-PASSWORD").text= password
            #session.findById("wnd[1]/tbar[0]/btn[0]").press()  #Boton Check para cerrar orden 
            session.findById("wnd[1]/tbar[0]/btn[12]").press()  #Boton Cancelar 
        except:
            Ordenes.append(orden) #Arreglar
            Info.append(orden)    #Arreglar
    except:
        print("No encontro La orden"+" "+str(orden))

#Exportar datos a TXT input= (Name=Nombre del documento,session=Engine)

def Export_TXT2(Name,session,Ruta=None):
    try:
        try:
            session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").pressToolbarButton("&NAVIGATION_PROFILE_TOOLBAR_EXPAND")
            session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")
            session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").selectContextMenuItem("&PC")
        except:
            session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")
            session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").selectContextMenuItem("&PC")
    except pywintypes.com_error:
        try:
            session.findById("wnd[0]/tbar[1]/btn[45]").press()
        except:
            session.findById("wnd[0]/tbar[1]/btn[9]").press ()
    finally:
        session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        if Ruta==None:
            session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\\Users\\prac.ingindustrial2\\OneDrive - Prebel S.A\\Escritorio\\SAP\\Archivos_CSV\\"
        else:
            session.findById("wnd[1]/usr/ctxtDY_PATH").text = Ruta
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = str(Name) + ".txt"
        session.findById("wnd[1]/usr/ctxtDY_FILE_ENCODING").text = "4310"
        session.findById("wnd[1]/tbar[0]/btn[11]").press()


#Boxlist Orden search arange multiple 
def Boxlist_Orden(session):
    try:
        try:
            session.findById("wnd[0]/usr/btn%_SO_AUFNR_%_APP_%-VALU_PUSH").press()
        except:
            session.findById("wnd[0]/usr/btn%_AUFNR_%_APP_%-VALU_PUSH").press()
    except:
        try:
            session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/btn%_S_PAUFNR_%_APP_%-VALU_PUSH").press()
        except:
            None
#Boxlist Orden search arange multiple 
def Boxlist_Material(session):
    try:
        try:
            session.findById("wnd[0]/usr/btn%_SO_MATNR_%_APP_%-VALU_PUSH").press()
        except:
            session.findById("wnd[0]/usr/btn%_QL_MATNR_%_APP_%-VALU_PUSH").press()
    except:
        try:
            session.findById("wnd[0]/usr/btn%_MATNR_%_APP_%-VALU_PUSH").press()
        except:
            session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/btn%_S_MATNR_%_APP_%-VALU_PUSH").press()

def Search_Ordenes_COOISPI(Transsaccion,Series,provision,session):      #(column Dataframe)
    session.StartTransaction(Transsaccion)
    Series=Series.to_clipboard(index=False, header=False)
    session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/btn%_S_PAUFNR_%_APP_%-VALU_PUSH").press()
    session.findById("wnd[1]/tbar[0]/btn[24]").press()
    session.findById("wnd[1]/tbar[0]/btn[8]").press()
    session.findById("wnd[0]/usr/ssub%_SUBSCREEN_TOPBLOCK:PPIO_ENTRY:1100/ctxtPPIO_ENTRY_SC1100-ALV_VARIANT").text = provision
    session.findById("wnd[0]/tbar[1]/btn[8]").press()



def Close_session(session):
    session.findById("wnd[0]").close()
    session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()




"""------------------------------------------ START Search ZD110----------------------------------------------------"""
def Search_ZSD110(Transsaccion,variant,provision,session):  #Optiene la tabla al buscar la transaccion
    session.StartTransaction(Transsaccion)
    session.findById("wnd[0]/tbar[1]/btn[17]").press()
    Varians_ZSD110=session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell")
    Rows_Varians_ZSD110=Varians_ZSD110.RowCount
    # List Variants
    List_Varian_ZSD110=[Varians_ZSD110.GetCellValue(i,"VARIANT") for i in range(0,Rows_Varians_ZSD110)] 
    indice=[indice for indice, dato in enumerate(List_Varian_ZSD110) if dato == variant]
    Varians_ZSD110.selectedRows = indice[0]
    session.findById("wnd[1]/tbar[0]/btn[2]").press()
    session.findById("wnd[0]/usr/ctxtPA_LAYOU").text = provision
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    tabla_zsd110=session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell")     #Select table cont
    return(tabla_zsd110)


# Descargar channels de Entrega y pedido
def Download_ZSD110_Channels(tabla_zsd110,Download_channel,session):
    """
    - tabla_zsd110: Tabla buscada en transaccion ZSD110
    - Download_channel: Descarga channels Entrega de interes
    - session: Retorna session
    """
    row_table=tabla_zsd110.RowCount
    List_Channels=[]
    for i in range(0,row_table-1):  #Se resta uno por los subtotales
        channelID=tabla_zsd110.GetCellValue(i,"VTWEG")
        tabla_zsd110.selectedRows = str(i)
        try:
            if channelID in Download_channel:
                session.findById("wnd[0]/tbar[1]/btn[8]").press()
                session.findById("wnd[1]/tbar[0]/btn[45]").press()
                session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select()
                session.findById("wnd[1]/tbar[0]/btn[0]").press()
                session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\\Users\\prac.ingindustrial2\\OneDrive - Prebel S.A\\Escritorio\\SAP\\Archivos_CSV\\"   #CAMBIAR RUTA
                session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "Canal"+str(channelID)+"_ENT.txt"
                session.findById("wnd[1]/usr/ctxtDY_FILE_ENCODING").text = "4310"
                session.findById("wnd[1]/tbar[0]/btn[11]").press()
                session.findById("wnd[1]").close()
            session.findById("wnd[0]/tbar[1]/btn[7]").press()   
            session.findById("wnd[1]/tbar[0]/btn[45]").press()
            session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select()
            session.findById("wnd[1]/tbar[0]/btn[0]").press()
            session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\\Users\\prac.ingindustrial2\\OneDrive - Prebel S.A\\Escritorio\\SAP\\Archivos_CSV\\"   #CAMBIAR RUTA
            session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "Canal"+str(channelID)+"_PEN.txt"
            session.findById("wnd[1]/usr/ctxtDY_FILE_ENCODING").text = "4310"
            session.findById("wnd[1]/tbar[0]/btn[11]").press()
            session.findById("wnd[1]").close()
            List_Channels.append(channelID)
            print("Se encontro detalle Channel ",channelID)
        except:
            print("No se encontro ",channelID)
    return(List_Channels)


# Buscar ZSD037

def Search_Pedidos_ZSD037(Transsaccion,Series,session,provision=None):      #(column Dataframe)
    """
    Transsaccion: Transsacion a buscar
    Series: Columna del dataframe que quiero copiar
    session: session del usuario
    provision: disposicion de interes
    """
    session.StartTransaction(Transsaccion)
    Series=Series.to_clipboard(index=False, header=False)
    session.findById("wnd[0]/usr/btn%_SP$00011_%_APP_%-VALU_PUSH").press()
    session.findById("wnd[1]/tbar[0]/btn[24]").press()
    session.findById("wnd[1]/tbar[0]/btn[8]").press()
    if provision!=None:
        session.findById("wnd[0]/usr/ctxt%LAYOUT").text = provision
    session.findById("wnd[0]/tbar[1]/btn[8]").press()




def Search_ZSD035D(Transsaccion,session,provision,date,variant):
    """
    Transsaccion: transaccion a buscar
    session: mantener seesion iniciada
    provision: Disposición
    date: dia de interes a buscar
    variant: Variante a visualizar
    """
    session.StartTransaction(Transsaccion)
    session.findById("wnd[0]/tbar[1]/btn[17]").press()
    session.findById("wnd[1]/usr/txtV-LOW").text = provision
    session.findById("wnd[1]/usr/txtENAME-LOW").text = variant
    session.findById("wnd[1]/usr/txtV-LOW").caretPosition = 8
    session.findById("wnd[1]/tbar[0]/btn[8]").press()
    if date==None:
        None
    else:
        session.findById("wnd[0]/usr/ctxtSP$00024-LOW").text = date
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    try:
        table=session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell")
        return(table)
    except:
        return(session)
    


def Search_ZSD10(transsaccion,Series,variant,provision,session):
    """
    Transsaccion: Transsacion a buscar
    Series: Columna del dataframe que quiero copiar
    session: session del usuario
    variant: Variante a buscar
    provision: disposicion de interes
    session: session activa
    """
    session.StartTransaction(transsaccion)
    session.findById("wnd[0]/tbar[1]/btn[17]").press()
    session.findById("wnd[1]/usr/txtV-LOW").text = provision
    session.findById("wnd[1]/usr/txtENAME-LOW").text = variant
    session.findById("wnd[1]/usr/txtV-LOW").caretPosition = 8
    session.findById("wnd[1]/tbar[0]/btn[8]").press()
    Series=Series.to_clipboard(index=False, header=False)
    session.findById("wnd[0]/usr/btn%_SP$00013_%_APP_%-VALU_PUSH").press()
    session.findById("wnd[1]/tbar[0]/btn[24]").press()
    session.findById("wnd[1]/tbar[0]/btn[8]").press()
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    try:
        tabla=session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell")
        return(tabla)
    except:
        return(session)
    

def Search_ZSD76(Transsaccion,session,provision,variant):
    """
    DETALLE DE ENTREGAS
    Transsaccion: Transsacion a buscar
    session: session activa
    provision: disposicion de interes
    variant: Variante a buscar
    """
    session.StartTransaction(Transsaccion)
    session.findById("wnd[0]/tbar[1]/btn[17]").press()
    session.findById("wnd[1]/usr/txtV-LOW").text = provision
    session.findById("wnd[1]/usr/txtENAME-LOW").text = variant
    session.findById("wnd[1]/usr/txtV-LOW").caretPosition = 8
    session.findById("wnd[1]/tbar[0]/btn[8]").press()
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    return(session)


def Search_Pedidos_ZSD127(Transsaccion,Series,session,Variant,provision=None):     
    """
    SEGUIMIENTO DE ESTADO DE PEDIDOS
    Transsaccion: Transsacion a buscar
    Series: Columna del dataframe que quiero copiar
    session: session del usuario
    Variant: Variante de visualizacion
    provision: disposicion o Layout
    """
    session.StartTransaction(Transsaccion)
    Search_table_Variant(session,Variant)
    Series=Series.to_clipboard(index=False, header=False)
    session.findById("wnd[0]/usr/btn%_SP$00004_%_APP_%-VALU_PUSH").press()
    session.findById("wnd[1]/tbar[0]/btn[24]").press()
    session.findById("wnd[1]/tbar[0]/btn[8]").press()
    if provision!=None:
        session.findById("wnd[0]/usr/ctxt%LAYOUT").text = provision
    session.findById("wnd[0]/tbar[1]/btn[8]").press()

def Search_VA05N(Transsaccion,session,provision,variant):
    session.StartTransaction(Transsaccion)
    session.findById("wnd[0]/tbar[1]/btn[17]").press()
    session.findById("wnd[1]/usr/txtV-LOW").text = provision
    session.findById("wnd[1]/usr/txtENAME-LOW").text = variant
    session.findById("wnd[1]/usr/txtV-LOW").caretPosition = 8
    session.findById("wnd[1]/tbar[0]/btn[8]").press()
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    return(session)

def Search_ZSD79(Transsaccion,Series,session):      #(column Dataframe)
    """
    Transsaccion: Transsacion a buscar
    Series: Columna del dataframe que quiero copiar
    session: session del usuario
    provision: disposicion de interes
    """
    session.StartTransaction(Transsaccion)
    Series=Series.to_clipboard(index=False, header=False)
    session.findById("wnd[0]/usr/btn%_SO_BSTKD_%_APP_%-VALU_PUSH").press()
    session.findById("wnd[1]/tbar[0]/btn[24]").press()
    session.findById("wnd[1]/tbar[0]/btn[8]").press()
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    table=session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell")
    return(table)

def Export_TXT(Name,session,Ruta=None):
    try:
        session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[2]").select()
    except:
        pass
    finally:
        session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        if Ruta==None:
            session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\\Users\\prac.ingindustrial2\\OneDrive - Prebel S.A\\Escritorio\\SAP\\Archivos_CSV\\"
        else:
            session.findById("wnd[1]/usr/ctxtDY_PATH").text = Ruta
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = str(Name) + ".txt"
        session.findById("wnd[1]/usr/ctxtDY_FILE_ENCODING").text = "4310"
        session.findById("wnd[1]/tbar[0]/btn[11]").press()

def Search_Table_ZSD79(table,session):
    """
    -table: Tabla ZSD79
    -session: session
    (Busca filtro especial de colores)
    """
    #Columna Pedido concluido 
    table.SelectColumn("LFGSK")
    #Columna entregas concluidas
    table.SelectColumn("WBSTK")
    session.findById("wnd[0]/mbar/menu[1]/menu[3]").select() 
    session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").text = "@0A@"  #Rojo Pedido
    session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN002-LOW").text = "@08@"  #Verde Entrega
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    Row=table.RowCount
    dic={"Nº Pedido cliente":[],"Clase Orden":[],"PrimFecEnt":[],"ÚltEntrega":[]}
    for i in range(0,Row):
        #Nº Pedido cliente Col "BSTNK"
        #Clase Orden Col "AUART"
        #PrimFecEnt Col "AUDAT"
        #ÚltEntrega Col "VDATU"
        Pedido_cliente,Clase_Orden,PrimFecEnt,ÚltEntrega=table.GetCellValue(i,"BSTNK"),table.GetCellValue(2,"AUART"),table.GetCellValue(2,"AUDAT"),table.GetCellValue(2,"VDATU") 
        dic["Nº Pedido cliente"].append(Pedido_cliente)
        dic["Clase Orden"].append(Clase_Orden)
        dic["PrimFecEnt"].append(PrimFecEnt)
        dic["ÚltEntrega"].append(ÚltEntrega)
    return(pd.DataFrame(dic))