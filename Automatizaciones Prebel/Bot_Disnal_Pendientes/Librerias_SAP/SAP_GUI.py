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
            session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\\Users\\prac.ingindustrial2\\OneDrive - Prebel S.A\\Documentos 1\\SAP\\SAP GUI\\"
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
