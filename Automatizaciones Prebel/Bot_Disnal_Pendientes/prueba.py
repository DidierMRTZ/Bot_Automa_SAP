from Librerias_SAP import SAP_GUI, Funtions
import win32com.client
import pandas as pd
from time import sleep
import json


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


"""----------------------------Inciar session----------------------------------------------------"""
# Insert User name and password
Keys=pd.read_excel("C:\\Users\\prac.ingindustrial2\\OneDrive - Prebel S.A\\Escritorio\\SAP\\Claves\\Keys.xlsx")
user=Keys["User"][0]
password=Keys["Password"][0]
# Initialize session
session=SAP_GUI.SessionSAP(user,password)

"""----------------- Search ZD110 -------------------------------------------------------"""
Transsaccion_ZSD110="zsd110"
Varian_FIRME="PEN-FIRME"
Varian_MERCADEO="PEN-MERCADEO"
Varian_DISNAL_TOTAL="PENDIENTES"
provision_ZSD110="/REVISIONPEN"

"""-----------------------------------BUSCAR y DESCARGAR TRANSSACCION ZD110---------------------------------------------------------------- """
# Pendiente Firme
Name_ZSD110_FIRME="Pendiente_Firme"  # Revisar suele no guardar
SAP_GUI.Search_ZSD110(Transsaccion_ZSD110,Varian_FIRME,provision_ZSD110,session)
SAP_GUI.Export_TXT2(Name_ZSD110_FIRME,session)