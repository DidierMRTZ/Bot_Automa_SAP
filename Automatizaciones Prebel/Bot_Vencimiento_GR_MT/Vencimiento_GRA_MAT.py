from Librerias_SAP import SAP_GUI
import win32com.client
import pandas as pd
import subprocess
from time import sleep
from datetime import datetime
from datetime import timedelta

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

def default_column(default_columns,dataframe):     #Parametros (default_columns: Columnas predeterminadas,dataframe:)
    diccionay_default_column={}
    if len(default_columns)==len(dataframe.columns):
        for i in range(0,len(default_columns)):
            diccionay_default_column[dataframe.columns[i]]=default_columns[i]   
        dataframe=dataframe.rename(columns=diccionay_default_column)   #Remplazo las columnas con las de default
    else:
        None   #asumo columnas originales como estandar
    return(dataframe)

def Date_Null(dataframe,date):
    filter=(date!="00.00.0000")
    dataframe=dataframe.loc[filter]
    return(dataframe)


"""----------------------------Inciar session----------------------------------------------------"""
# Insert User name and password
Keys=pd.read_excel("C:\\Users\\prac.ingindustrial2\\OneDrive - Prebel S.A\\Escritorio\\SAP\\Claves\\Keys.xlsx")
user=Keys["User"][1]
password=Keys["Password"][1]
# Initialize session
Iniciar=SAP_GUI.SessionSAP(user,password)

# Start transsación
Transsacion="mb52"
provision="GRANELES"   
variant="AUXNVOSPT"
Buscar_MB52_GRANELES=SAP_GUI.Search_MB52(Transsacion,Iniciar,provision,variant)


Name_GRANELES="GRANELES"  # Name file
SAP_GUI.Export_TXT2(Name_GRANELES,Iniciar)

Ruta_GRANELES="C:\\Users\\prac.ingindustrial2\\OneDrive - Prebel S.A\\Documentos 1\\SAP\\SAP GUI\\"+str(Name_GRANELES)+".txt"
data_GRANELES = pd.read_csv(Ruta_GRANELES, skiprows=1, delimiter='\t')

data_GRANELES=Clean_Columns(data_GRANELES)

default_column_MB52=['Material', 'Texto breve de material', 'Gpo.artíc.', 'Ce.', 'Alm.',
       'Lote', 'UMB', 'Libre utilización', 'En control calidad',
       'Valor libre util.', 'Valor en insp.cal.', 'Cad./FPC']   #Disposicion PROXI A VENCER

data_GRANELES=default_column(default_column_MB52,data_GRANELES) 


data_GRANELES=Date_Null(data_GRANELES,data_GRANELES['Cad./FPC'])

"""-------------------------Filtro para eliminar subtotales y combertir fechas-------------------------------"""
data_GRANELES=data_GRANELES[data_GRANELES.Material.notnull()]
data_GRANELES["Cad./FPC"]=data_GRANELES["Cad./FPC"].apply(lambda x: datetime.strptime(x,'%d.%m.%Y'))  #Change columns str to datetime

Data_Denominacion=Clean_Columns(pd.read_csv("C:\\Users\\prac.ingindustrial2\\OneDrive - Prebel S.A\\Documentos 1\\SAP\\SAP GUI\\Denominacion_Articulos.txt",delimiter='\t'))[["Grupo art.","Denom.gr.artíc."]]

data_GRANELES=pd.merge(Data_Denominacion,data_GRANELES,how="right",left_on="Grupo art.",right_on="Gpo.artíc.")


"""---------------------------------Filtro graneles vencidos-------------------"""
now=datetime.now()
month_vencido=now-timedelta(days=30) 
mask_vencido=(data_GRANELES["Cad./FPC"]<=now) & (data_GRANELES["Cad./FPC"]>=month_vencido)
data_GRANELES_vencidos=data_GRANELES.loc[mask_vencido]
#Sort values
data_GRANELES_vencidos=data_GRANELES_vencidos.sort_values(by=["Cad./FPC"],ascending=False)
data_GRANELES_vencidos=data_GRANELES_vencidos.reset_index(drop=True)
data_GRANELES_vencidos["Cad./FPC"]=data_GRANELES_vencidos["Cad./FPC"].dt.date   #convet datetime to date 



# filter dates between now a 30 month later
now=datetime.now()
month=timedelta(days=6*30)   # Change limiter day(30 days= 1 month)
month=now+month
mask=(data_GRANELES["Cad./FPC"]>=now) & (data_GRANELES["Cad./FPC"]<=month)
data_GRANELES=data_GRANELES.loc[mask]
#Sort values
data_GRANELES=data_GRANELES.sort_values(by=["Cad./FPC","Libre utilización"])
data_GRANELES=data_GRANELES.reset_index(drop=True)
data_GRANELES["Cad./FPC"]=data_GRANELES["Cad./FPC"].dt.date   #convet datetime to date 


"""------------DATAFRAME GRANALES PARA INFORME-----------------------------------------------------------------------"""

data_Graneles_informe=data_GRANELES[['Material', 'Texto breve de material', 'Gpo.artíc.','Denom.gr.artíc.', 'Ce.', 'Alm.',
       'Lote',  'Libre utilización', 'En control calidad','UMB', 'Valor libre util.', 'Valor en insp.cal.','Cad./FPC']]


data_Graneles_Vencidos_informe=data_GRANELES_vencidos[['Grupo art.', 'Denom.gr.artíc.', 'Material', 'Texto breve de material',
       'Gpo.artíc.', 'Ce.', 'Alm.', 'Lote', 'UMB',
       'Cad./FPC']]


"""-----------------------------GENERO SUBTOTALES------------------------------------------------------------"""

data_GRANELES_Libre=Clean_column_number(data_GRANELES["Valor libre util."]).sum()
data_GRANELES_Calidad=Clean_column_number(data_GRANELES['Valor en insp.cal.']).sum()
Total_Graneles=data_GRANELES_Libre+data_GRANELES_Calidad
Total_Graneles="${:,.2f}".format(Total_Graneles)


""" Search MB52 raw material"""

# Start transsación
Transsacion="mb52"
provision="MATERIA PRIMA"  
variant="AUXNVOSPT"
Buscar_MB52_MATERIAS=SAP_GUI.Search_MB52(Transsacion,Iniciar,provision,variant)
# Save file



Name_Materias="MATERIAS"  # Revisar suele no guardar
SAP_GUI.Export_TXT2(Name_Materias,Iniciar)
# Pass the route and read file 
Ruta_Materias="C:\\Users\\prac.ingindustrial2\\OneDrive - Prebel S.A\\Documentos 1\\SAP\\SAP GUI\\"+str(Name_Materias)+".txt"
data_Materias=pd.read_csv(Ruta_Materias,delimiter="\t")
# Change, drop and standardize columns
data_Materias=Clean_Columns(data_Materias)
data_Materias=default_column(default_column_MB52,data_Materias) 
data_Materias=Date_Null(data_Materias,data_Materias['Cad./FPC'])


"""-------------------------Filtro para eliminar subtotales y combertir fechas-------------------------------"""
data_Materias=data_Materias[data_Materias.Material.notnull()]
data_Materias["Cad./FPC"]=data_Materias["Cad./FPC"].apply(lambda x: datetime.strptime(str(x),'%d.%m.%Y'))  #Change columns str to datetime

data_Materias=pd.merge(Data_Denominacion,data_Materias,how="right",left_on="Grupo art.",right_on="Gpo.artíc.")


"""---------------------------Buscamos las clases de material------------- """
Data_Material_Total=data_Materias
Lista_tipo_Materias=data_Materias["Gpo.artíc."].drop_duplicates().reset_index(drop=True)


"""--------------------------------------MATERIAS PRIMAS VENCIDAS-------------------------"""

mask_vencido=(data_Materias["Cad./FPC"]<=now) & (data_Materias["Cad./FPC"]>=month_vencido)
data_Materias_vencidos=data_Materias.loc[mask_vencido]
#Sort values
data_Materias_vencidos=data_Materias_vencidos.sort_values(by=["Cad./FPC"],ascending=False)
data_Materias_vencidos=data_Materias_vencidos.reset_index(drop=True)
data_Materias_vencidos["Cad./FPC"]=data_Materias_vencidos["Cad./FPC"].dt.date   #convet datetime to date 


mask=(data_Materias["Cad./FPC"]>=now) & (data_Materias["Cad./FPC"]<=month)
data_Materias=data_Materias.loc[mask]
data_Materias=data_Materias.sort_values(by=["Cad./FPC","Libre utilización"]).reset_index(drop=True)
data_Materias["Cad./FPC"]=data_Materias["Cad./FPC"].dt.date

#Close Session
SAP_GUI.Close_session(Iniciar)


"""------------DATAFRAME MATERIAS PARA INFORME-----------------------------------------------------------------------"""
data_Materia_informe=data_Materias[['Material', 'Texto breve de material', 'Gpo.artíc.','Denom.gr.artíc.', 'Ce.', 'Alm.',
       'Lote',  'Libre utilización', 'En control calidad','UMB', 'Valor libre util.', 'Valor en insp.cal.','Cad./FPC']]

data_Materia_informe_vencidos=data_Materias_vencidos[['Grupo art.', 'Denom.gr.artíc.', 'Material', 'Texto breve de material',
       'Gpo.artíc.', 'Ce.', 'Alm.', 'Lote', 'UMB',
       'Cad./FPC']]

"""-----------------------------GENERO SUBTOTALES------------------------------------------------------------"""

data_Materia_Libre=Clean_column_number(data_Materias["Valor libre util."]).sum()
data_Materia_Calidad=Clean_column_number(data_Materias['Valor en insp.cal.']).sum()
Total_Materia=data_Materia_Libre+data_Materia_Calidad
Total_Materia="${:,.2f}".format(Total_Materia)

"""Send email"""

correos="prac.ingindustrial2@prebel.com.co"

def send_emails(*args,emails="",htmlbody="",subject=""):
    email=emails
    outlook=win32com.client.Dispatch("outlook.application")
    mail=outlook.CreateItem(0)
    mail.Subject=subject+" "+datetime.now().strftime('%#d %b %Y %H:%M')
    mail.To=email
    mail.HTMLBody=htmlbody.format(*args)
    mail.Send()

def style_df(df):
    return df.style \
        .set_table_styles([{'selector': "table,tr,th,td", 'props': [("border", "1px solid"), ('color', '#000'),("text-align","center")]}]) \
        .highlight_between(subset=["Cad./FPC"],color='#FF5733',left=now.date(),right=(now.date()+timedelta(days=7))) \

html="""
    <h2 style="text-align: center">REPORTE CADUCIDAD DE GRANELES</h2>
    <p> Por medio del presente informe se evidencia los graneles segun el tipo de material que estan proximos a vencer durante un mes a partir de la 
        fecha actual, ademas de las cantidades con su respectiva unidad de medida que estan en libre utilización.</p>

    <h4 style="color: black;" > Advertencia: El valor económico potencial en riesgo de vencimiento es de: </h4>
    
    <h1 style="color: red;" > {1} </h1>

    <div"> {0} </div>

    <h4 style="color: red;" > Advertencia: Los graneles a vencer en los proximos 7 días se resaltan en rojo</h4>

    <p> Anticipo sinceros agradecimientos. </p>
 """

html2="""
    <h2 style="text-align: center"> REPORTE CADUCIDAD DE MATERIAS PRIMAS</h2>
    <p> Por medio del presente informe se evidencia las materias primas segun el tipo de material que estan proximos a vencer durante un mes a partir de la 
        fecha actual, ademas de las cantidades con su respectiva unidad de medida que estan en libre utilización.</p>

    <h4 style="color: black;" > Advertencia: El valor económico potencial en riesgo de vencimiento es de: </h4>
    
    <h1 style="color: red;" > {1} </h1>
    
    <div"> {0} </div>

    <h4 style="color: red;" > Advertencia: Las materias primas proximas a vencer en los proximos 7 días se resaltan en rojo</h4>

    <p> Anticipo sinceros agradecimientos. </p>
 """

Send = style_df(data_Graneles_informe)  #Style between LI and LS
Send2=style_df(data_Materia_informe)     #Style between LI and LS



# Definir el diccionario de formato
try:
    formato_Graneles = {'Ce.': '{:.0f}',"Alm.":'{:.0f}'}
    formato_material= {'Ce.': '{:.0f}',"Alm.":'{:.0f}',"Lote":'{:.0f}'}
    # Aplicar el formato a la columna 'Altura'
    Send = Send.format(formato_Graneles)
    Send2=Send2.format(formato_material)
except:
    None

send_emails(Send.to_html(),Total_Graneles,emails=correos,htmlbody=html,subject="REPORTE GRANELES PROXIMOS A VENCER")
send_emails(Send2.to_html(),Total_Materia,emails=correos,htmlbody=html2,subject="REPORTE MATERIAS PRIMAS PROXIMOS A VENCER")


"""-------------------------------------Send email material y granel vencidos--------------------------------"""

correos="prac.ingindustrial2@prebel.com.co;juan.murillo@prebel.com.co"

def send_emails(*args,emails="",htmlbody="",subject=""):
    email=emails
    outlook=win32com.client.Dispatch("outlook.application")
    mail=outlook.CreateItem(0)
    mail.Subject=subject+" "+datetime.now().strftime('%#d %b %Y %H:%M')
    mail.To=email
    mail.HTMLBody=htmlbody.format(*args)
    mail.Send()

def style_df(df):
    return df.style \
        .set_table_styles([{'selector': "table,tr,th,td", 'props': [("border", "1px solid"), ('color', '#000'),("text-align","center")]}]) \

html="""
    <h2 style="text-align: center">REPORTE CADUCIDAD DE GRANELES</h2>
    <p> Por medio del presente informe se evidencia los graneles y materias primas vencidas, para su correspondiente análisis. </p>

    <h2 style="color: black;" > GRANELES VENCIDOS</h4>

    <div"> {0} </div>

    <h2 style="color: black;" > MATERIA PRIMA VENCIDA </h4>
    

    <div"> {1} </div>


    <h4> Anticipo sinceros agradecimientos. </h4>
 """
formato_Graneles = {'Ce.': '{:.0f}',"Alm.":'{:.0f}'}


Send = style_df(data_Graneles_Vencidos_informe)  #Style between LI and LS
Send2=style_df(data_Materia_informe_vencidos)     #Style between LI and LS
Send=Send.format(formato_Graneles)
Send2=Send2.format(formato_Graneles)
send_emails(Send.to_html(),Send2.to_html(),emails=correos,htmlbody=html,subject="REPORTE GRANELES Y MARIA PRIMA VENCIDO")
