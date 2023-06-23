import SAP_GUI
import win32com.client
import pandas as pd
from datetime import datetime
from datetime import timedelta
import numpy as np
import re
import calendar
from time import sleep



"""######################################################FUNCIONES##############################################################################"""

"""-------------------------------------FUNCION PARA LIMPIAR ARCHIVOS----------------------------------------------------------"""
def Clean_Columns(List):
    Set_Columns=[i.strip() for i in List.columns]   #Alert in list
    clean_column=[i if "Unnamed" in i else None for i in List.columns]
    clean_column = list(filter(lambda x: x is not None, clean_column)) 
    List=List.set_axis(Set_Columns, axis=1).drop(clean_column,axis=1)  #Drop Unammed: 0
    return(List)

"""--------------------------------FUNCION PARA LIMPIAR NUMEROS A FLOAT--------------------------------------------------------"""

def Clean_column_number(column):
    column=column.apply(lambda x: float(str(x).strip().replace(",","")))
    return(column)


"""-----------------------------Funcion para cambiar los estados en Cooispi------------------------------------------------------------------"""
#### Esta funcion debe estar abieta la COOISPI PASO 1 ###########################
def Change_Estate_Cooispi(session,Data):
    table=session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell")
    for i,j in Data.iterrows():
        try:
            if not "LIB. NOTP ENTR PREC" in j["Status de sistema"]:
                table.SetCurrentCell(i,"AUFNR")
                session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").pressToolbarButton("OBAE")
                session.findById("wnd[0]/usr/tabsTABSTRIP_5115/tabpKOWE").select()
                session.findById("wnd[0]/usr/tabsTABSTRIP_5115/tabpKOWE/ssubSUBSCR_5115:SAPLCOKO:5190/chkAFPOD-ELIKZ").selected = True
                # session.findById("wnd[0]/tbar[0]/btn[15]").press()  #Select botton Cnacel
                session.findById("wnd[0]/tbar[0]/btn[11]").press()
                try:
                    session.findById("wnd[1]/usr/btnSPOP-OPTION2").press() #boton Back
                except:
                    None
                # print(i,j["Status de sistema"])
        except:
            continue
    session.findById("wnd[0]").sendVKey(5)  #Press F5



def UPGRADE_CORR(Transsaccion,Data_CORR,Ordenes_Cooispi,Data_cooispi_validacion,session,Data_fases):
    c=0
    for i,j in Data_CORR.iterrows():
        if j['Ctd.teór.']>=j['Ctd entreg']:   # Compara DATA transformada CORR
            print(Ordenes_Cooispi[i],"------------------------------------------------")
            try:
                session.StartTransaction(Transsaccion)  # inicio Transsacción
                session.findById("wnd[0]/usr/subCOL_SUG_TICKET1:SAPLCORU:5807/tblSAPLCORUTABCDEF_0807/ctxtAFRD-AUFNR[1,0]").text = str(Ordenes_Cooispi[i]).strip()   #Paso ordenes una a una limpiadas
                session.findById("wnd[0]/usr/subCOL_SUG_TICKET1:SAPLCORU:5807/tblSAPLCORUTABCDEF_0807/txtAFRD-LMNGA[8,0]").text = str(Data_cooispi_validacion['Ctd entreg'][i]).strip()  #Copio data original datos SAP
                session.findById("wnd[0]/usr/subCOL_SUG_TICKET1:SAPLCORU:5807/tblSAPLCORUTABCDEF_0807/ctxtAFRD-MEINH[9,0]").text = str(Data_cooispi_validacion.Unidad[i]).strip()        #Copio data original datos SAP
                session.findById("wnd[0]").sendVKey(0)  #Press Enter
                #Entrego las fases
                session.findById("wnd[0]/usr/subCOL_TICKET1:SAPLCORU:5808/tblSAPLCORUTABCNTR_0808/txtAFRUD-VORNR[3,0]").text=Data_fases[0][0]
                session.findById("wnd[0]/usr/subCOL_TICKET1:SAPLCORU:5808/tblSAPLCORUTABCNTR_0808/txtAFRUD-VORNR[3,1]").text=Data_fases[0][1]
                session.findById("wnd[0]/usr/subCOL_TICKET1:SAPLCORU:5808/tblSAPLCORUTABCNTR_0808/txtAFRUD-VORNR[3,2]").text=Data_fases[0][2]
                session.findById("wnd[0]/usr/subCOL_TICKET1:SAPLCORU:5808/tblSAPLCORUTABCNTR_0808/txtAFRUD-VORNR[3,3]").text=Data_fases[0][3]
                session.findById("wnd[0]/usr/subCOL_TICKET1:SAPLCORU:5808/tblSAPLCORUTABCNTR_0808/txtAFRUD-VORNR[3,4]").text=Data_fases[0][4]
                session.findById("wnd[0]/usr/subCOL_TICKET1:SAPLCORU:5808/tblSAPLCORUTABCNTR_0808/txtAFRUD-VORNR[3,5]").text=Data_fases[0][5]
                session.findById("wnd[0]/usr/subCOL_TICKET1:SAPLCORU:5808/tblSAPLCORUTABCNTR_0808/txtAFRUD-VORNR[3,6]").text=Data_fases[0][6]
                session.findById("wnd[0]/usr/subCOL_TICKET1:SAPLCORU:5808/tblSAPLCORUTABCNTR_0808/txtAFRUD-VORNR[3,7]").text=Data_fases[0][7]
                session.findById("wnd[0]/usr/subCOL_TICKET1:SAPLCORU:5808/tblSAPLCORUTABCNTR_0808/txtAFRUD-VORNR[3,8]").text=Data_fases[0][8]
                session.findById("wnd[0]").sendVKey(9)  #Press F9
                print("SE actualizo "+str(Ordenes_Cooispi[i].strip()))
                try:
                    session.findById("wnd[0]").sendVKey(5)  #Press F5

                    session.findById("wnd[1]/usr/chkTCORU-VSSLE").selected = True
                    session.findById("wnd[1]/usr/chkTCORU-VSSZT").selected = True
                    session.findById("wnd[1]/tbar[0]/btn[0]").press()
                    session.findById("wnd[0]/tbar[0]/btn[11]").press() ### Boton para actualizar
                    
                    # #   En este caso por si queremos probar
                    # session.findById("wnd[0]/tbar[0]/btn[15]").press()
                    # sleep(2)
                    # session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
                except:
                    print("posiblemente se firmo "+Ordenes_Cooispi[i].strip())
                    session.findById("wnd[1]/usr/btnSPOP-OPTION2").press()
            except:
                session.findById("wnd[1]").close()    #Sale ventana y la cierra
                continue
        else:
            print(Ordenes_Cooispi[i].strip(),Data_cooispi_validacion['Ctd.teór.'][i],Data_cooispi_validacion['Ctd entreg'][i],Data_cooispi_validacion.Unidad[i],"ES MENOR")
            c=c+1
    print("Menores",c)

"""######################################################FUNCIONES##############################################################################"""
#### hashdfhsdfk


def UDGRADE_ESTATE(dateIni,dateFin):
    """-------------------------------------------------INICIAR SESSION-------------------------------------------------------------------------"""
    user="vvalenciao"
    password="Abril2023%"
    session=SAP_GUI.SessionSAP(user,password)

    """-------------------------------------------------BUSCAR EN LA COOISPI-------------------------------------------------------------------------"""
    dateIni=dateIni   #Fecha inicial
    dateFin=dateFin                      #Fecha Final

    #Search Cooispi
    provisionCooispi="CAMBIO ESTADO"      #Provision
    variantCooispi="vvalenciao"           #variante
    TranssaccionCooispi="cooispi"         #Transsacion
    disposicion="ACTUALIZAR"              #Layout

    SAP_GUI.Search_COOISPI(TranssaccionCooispi,session,variantCooispi,provisionCooispi,disposicion,dateIni,dateFin)
    """-----------------------------------------------READ AND LOAD ARCHIVO----------------------------------------------"""
    Name_COOISPI="Validadcion"  # Revisar suele no guardar
    SAP_GUI.Export_TXT2(Name_COOISPI,session)

    # Pass the route and read file 
    url_cooispi_validacion="C:\\Users\\prac.ingindustrial2\\OneDrive - Prebel S.A\\Documentos 1\\SAP\\SAP GUI\\"+str(Name_COOISPI)+".txt"
    Data_cooispi_validacion=pd.read_csv(url_cooispi_validacion,skiprows=1,delimiter="\t")


    """--------------------------------LIMPIA COLUMNAS Y ELIMINA FILAS CON VALORES VACIOS-------------------------------------------"""
    Data_cooispi_validacion=Clean_Columns(Data_cooispi_validacion)
    Data_cooispi_validacion=Data_cooispi_validacion[~Data_cooispi_validacion.Material.isnull()]

    """--------------------------------------------REMPLAZO COLUMNAS DE INTERES--------------------------------------------------------------"""
    Estandarizar_columnas={'Ctd.teórica':"Ctd.teór.",'Ctd entreg':"Ctd entreg"}
    Data_cooispi_validacion=Data_cooispi_validacion.rename(columns=Estandarizar_columnas)
    """--------------------------------------------REMPLAZO COLUMNAS DE INTERES--------------------------------------------------------------"""


    # COLUMNAS COMPARATIVAS CAN TEORICA VS CANTIDAD ENTREGADA
    Data_CORR=Data_cooispi_validacion[["Ctd.teór.","Ctd entreg"]]
    Data_CORR["Ctd.teór."]=Clean_column_number(Data_CORR["Ctd.teór."])
    Data_CORR["Ctd entreg"]=Clean_column_number(Data_CORR["Ctd entreg"])

    """---------------------------------------------APLICO FUNCION CAMBIAR ESTADOS PASO 1 #############------------------------------------------"""
    ##Aplyy funtion 
    Change_Estate_Cooispi(session,Data_cooispi_validacion)

    """---------------------------------------------------------Transformo datos ordenes---------------------------------------------------"""

    Ordenes_Cooispi=Data_cooispi_validacion.Orden.apply(lambda x: re.findall("\d*(?=.)",str(x))[0])

    """-----------------------------------Creo las fases y las copio al portapapeles--------------------------------------------------------"""
    #Ordenes_Cooispi
    fases={0:"0012",1:"0031",2:"0032",3:"0033",4:"0034",5:"0035",6:"0036",7:"0037",8:"0057"}
    Data_fases=pd.DataFrame(fases.values(),fases.keys())
    Data_fases[0].to_clipboard(header=False,index=False)


    """-------------------------------------COMPROBAR LOS LAS COLUMNAS EN LA COR CANT ENTREGADA y NOTIFICADA ------------------------------------"""
    #Parametros DATA_CORR transformada CORR, Transsacción, Ordenes_Cooispi,
    Transsaccion="CORR"

    UPGRADE_CORR(Transsaccion,Data_CORR,Ordenes_Cooispi,Data_cooispi_validacion,session,Data_fases)


    """-------------------------------------COMPROBAR LOS LAS COLUMNAS ------------------------------------"""
    Transsaccion="cooispi"
    Series=Ordenes_Cooispi
    provision="/JPZ"
    SAP_GUI.Search_Ordenes_COOISPI(Transsaccion,Ordenes_Cooispi,provision,session)


    """-----------------------------------------------READ AND LOAD ARCHIVO COMPARATIVO----------------------------------------------"""
    Comprobante_Cooispi="Comprobacion"  # Revisar suele no guardar
    SAP_GUI.Export_TXT2(Comprobante_Cooispi,session)

    # Pass the route and read file 
    url_cooispi_comprobante="C:\\Users\\prac.ingindustrial2\\OneDrive - Prebel S.A\\Documentos 1\\SAP\\SAP GUI\\"+str(Comprobante_Cooispi)+".txt"
    Data_cooispi_comprobante=pd.read_csv(url_cooispi_comprobante,skiprows=1,delimiter="\t")

    """--------------------------------------------------------Limpiar Datos y dar formato---------------------------------------------------------------"""
    Data_cooispi_comprobante=Clean_Columns(Data_cooispi_comprobante).rename(columns=Estandarizar_columnas)
    Data_cooispi_comprobante=Data_cooispi_comprobante[~Data_cooispi_comprobante.Material.isnull()]
    Data_cooispi_comprobante.Orden=Data_cooispi_comprobante.Orden.astype(int)


    Data_cooispi_comprobante=Data_cooispi_comprobante[['Orden','Texto breve material', 'Ctd.teór.', 'Ctd entreg', 'Ctd.notif.',
        'Unidad','Status de sistema']]


    """----------------------------------------------FUNCIONES PARA ENVIAR CORREOS--------------------------------------------"""
    def send_emails(*args,emails="",htmlbody="",subject=""):
        email=emails
        outlook=win32com.client.Dispatch("outlook.application")
        mail=outlook.CreateItem(0)
        mail.Subject=subject+" "+datetime.now().strftime('%#d %b %Y %H:%M')
        mail.To=email
        mail.HTMLBody=htmlbody.format(*args)
        mail.Send()


    def style_df(df,Diferente):
        return df.style \
            .set_table_styles([{'selector': "table,tr,th,td", 'props': [("border", "1px solid"), ('color', '#000'),("text-align","center")]}]) \
            .highlight_between(subset=["Status de sistema"],color='#FF5733',left=Diferente,right=Diferente) \

    def style_df_stand(df):
        return df.style \
            .set_table_styles([{'selector': "table,tr,th,td", 'props': [("border", "1px solid"), ('color', '#000'),("text-align","center")]}]) 


    """Send email"""

    correos="prac.ingindustrial2@prebel.com.co;vanessa.valencia@prebel.com.co"
    now=datetime.now()


    html="""
        <h2 style="text-align: center"> REPORTE ACTUALIZACIÓN DE ESTADOS y CANTIDAD NOTIFICADA</h2>
        <p> Por medio del presente informe se evidencia las ordenes actualizadas respecto al estado del sistema y cantidad notificada.</p>

        <h4 style="color: red;" > Advertencia: Se evidencia en color rojo las ordenes que posiblemente no estan actualizadas.</h4>
        
        <div"> {0} </div>
    """


    """-------------------------------------- BUSCO LAS QUE SEA DIFERENTE-------------------------------------------------"""

    indice=[indice for indice, dato in enumerate(Data_cooispi_comprobante["Status de sistema"]) if 'LIB. NOTP ENTR' not in dato]

    try:
        Diferente=Data_cooispi_comprobante["Status de sistema"][indice[0]]
        Send = style_df(Data_cooispi_comprobante,Diferente)  #Style between LI and LS     #Data de firmas
        send_emails(Send.to_html(),emails=correos,htmlbody=html,subject="REPORTE ACTUALIZACION DE ESTADOS")
        print("Se envio Correo 1")
    except:
        Send = style_df_stand(Data_cooispi_comprobante)
        send_emails(Send.to_html(),emails=correos,htmlbody=html,subject="REPORTE ACTUALIZACION DE ESTADOS")
        print("Se envio Correo 2")
    return(Data_cooispi_comprobante,Ordenes_Cooispi,Data_cooispi_validacion)
    #SAP_GUI.Close_session(session)


"""---------------------------- AQUI UTILIZO LAS HORAS PARA COMPARAR--------------------"""
day=datetime.now().date()
Día_Máximo=calendar.monthrange(day.year,day.month)[1]  #Saco maximo día del mes
dia_ma=str(day.year)+"-"+str(day.month)+"-"+str(Día_Máximo)
Día_Máximo=datetime.strptime(dia_ma,"%Y-%m-%d").date()
Día_Máximo_day=Día_Máximo.day
print(day.strftime("%A")+str(day.day)+"---------------------------"+str(Día_Máximo))

#nota ver el mes de Marzo-Abril
Día_Máximo


"""------------------------------Calculamos mes anterior--------------------------"""
# Calcula el primer día del mes actual
primer_dia_mes_actual =datetime(day.year, day.month, 1)
# Resta un día al primer día del mes actual
ultimo_dia_mes_anterior = primer_dia_mes_actual - timedelta(days=1)
# Obtén el mes anterior
mes_anterior = datetime(ultimo_dia_mes_anterior.year, ultimo_dia_mes_anterior.month, 1)

Dia_mes_anterior=mes_anterior.day

dia_mes_anterior=calendar.monthrange(mes_anterior.year,mes_anterior.month)[1] 



Dos_dias_atras=day-timedelta(days=2)
tres_dias_atras=day-timedelta(days=3)
Cuatro_dias_atras=day-timedelta(days=4)
if day.day==1 or day.day==Día_Máximo_day:
    print("No corre")
    print("Hoy es "+str(day.day)+" "+"del mes "+str(day.month)+" "+"del año "+str(day.year)+" ------- "+day.strftime("%A"))
elif (Dos_dias_atras.strftime("%A") in ['Thursday','Friday','Saturday']) and (Dos_dias_atras.day in [dia_mes_anterior,1]):  #Aca es del mes anterior ojo
    print("No corre")
    print("Hoy es "+str(day.day)+" "+"del mes "+str(day.month)+" "+"del año "+str(day.year),"------- "+day.strftime("%A"))
    print("El ultimo día es "+str(Dos_dias_atras.day)+" "+"del mes "+str(Dos_dias_atras.month)+" "+"del año "+str(Dos_dias_atras.year),"------- "+Dos_dias_atras.strftime("%A"))
elif (day.day==2):
    print("Corre Normal")
    Tiempo_atras=day-timedelta(days=2)
    print("Hoy es "+str(day.day)+" "+"del mes "+str(day.month)+" "+"del año "+str(day.year),"------- "+day.strftime("%A"))
    print("Al dia actual le resto 2 "+str(Dos_dias_atras.day)+" "+"del mes "+str(Dos_dias_atras.month)+" "+"del año "+str(Dos_dias_atras.year),"------- "+Dos_dias_atras.strftime("%A"))
    dateIni=(datetime.now()-timedelta(days=4)).strftime("%d-%m-%Y").replace("-",".")   #Fecha inicial
    dateFin=datetime.now().strftime("%d-%m-%Y").replace("-",".")                       #Fecha Final
    UDGRADE_ESTATE(dateIni,dateFin)
elif (Cuatro_dias_atras.strftime("%A") in ['Thursday','Friday'] and Cuatro_dias_atras.day==dia_mes_anterior):
    print(day.day)
    print("Al dia actual le resto 4 y corre")
    dateIni=(datetime.now()-timedelta(days=4)).strftime("%d-%m-%Y").replace("-",".")   #Fecha inicial
    dateFin=datetime.now().strftime("%d-%m-%Y").replace("-",".")                       #Fecha Final
    UDGRADE_ESTATE(dateIni,dateFin)
elif (tres_dias_atras.strftime("%A") in ['Saturday'] and tres_dias_atras.day==dia_mes_anterior):
    print(day.day)
    print("Al dia actual le resto 3 y corre")
    dateIni=(datetime.now()-timedelta(days=3)).strftime("%d-%m-%Y").replace("-",".")   #Fecha inicial
    dateFin=datetime.now().strftime("%d-%m-%Y").replace("-",".")                       #Fecha Final
    UDGRADE_ESTATE(dateIni,dateFin)
else:
    print(day)
    print(Dos_dias_atras.strftime("%A"))
    print("Corre Normal")
    dateIni=(datetime.now()-timedelta(days=2)).strftime("%d-%m-%Y").replace("-",".")   #Fecha inicial
    dateFin=datetime.now().strftime("%d-%m-%Y").replace("-",".")                       #Fecha Final
    print(dateIni,dateFin)
    UDGRADE_ESTATE(dateIni,dateFin)
    print("Entra aqui ------")





