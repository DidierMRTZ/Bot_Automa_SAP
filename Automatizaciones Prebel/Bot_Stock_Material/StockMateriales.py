import SAP_GUI
import win32com.client
import pandas as pd
import subprocess
from time import sleep
from datetime import datetime
from datetime import timedelta


""" Search MB52 bulk"""
# Insert User name and password
user="auxnvospt"
password="Lanzamientos11"
# Initialize session
Iniciar=SAP_GUI.SessionSAP(user,password)

""" Filtro de día"""
now=datetime.now()


""" Search MB52 raw MATERIA PRIMA (MA)"""
# Start transsación
Transsacion_MA="mb52"
provision_MA="MATERIA PRIMA"      #"MATERIA PRIMA"  
variant_MA="AUXNVOSPT"
Buscar_MA=SAP_GUI.Search_MB52(Transsacion_MA,Iniciar,provision_MA,variant_MA)
# Save file
Name_MB52_MA="MP"  # Revisar suele no guardar
SAP_GUI.Export_TXT2(Name_MB52_MA,Iniciar)
# Pass the route and read file 
url_MB52_MA="C:\\Users\\prac.ingindustrial2\\OneDrive - Prebel S.A\\Documentos 1\\SAP\\SAP GUI\\"+str(Name_MB52_MA)+".txt"
Data_Mb52_MA=pd.read_csv(url_MB52_MA,delimiter="\t")

"""----------------------------------------------CLEAN DATA-----------------------------------------------------------"""
# Change, drop and standardize columns
Columns_MA=[i.strip() for i in Data_Mb52_MA.columns]   #Alert in list
Data_Mb52_MA=Data_Mb52_MA.set_axis(Columns_MA, axis=1).drop(["Unnamed: 0"],axis=1)  #Drop Unammed: 0
#normalizo estas columnas
Data_Mb52_MA["Libre utilización"]=Data_Mb52_MA["Libre utilización"].apply(lambda x: x.strip())
Data_Mb52_MA["En control calidad"]=Data_Mb52_MA["En control calidad"].apply(lambda x: x.strip())


""" Search MB52 raw MATERIAL DE EMPAQUE (ME)"""
# Start transsación
Transsacion_ME="mb52"
provision_ME="MATERIAL EMP"      #"MATERIA PRIMA"  
variant_ME="AUXNVOSPT"
Buscar_ME=SAP_GUI.Search_MB52(Transsacion_ME,Iniciar,provision_ME,variant_ME)
# Save file
Name_MB52_ME="ME"  # Revisar suele no guardar
SAP_GUI.Export_TXT2(Name_MB52_ME,Iniciar)
# Pass the route and read file 
url_MB52_ME="C:\\Users\\prac.ingindustrial2\\OneDrive - Prebel S.A\\Documentos 1\\SAP\\SAP GUI\\"+str(Name_MB52_ME)+".txt"
Data_Mb52_ME=pd.read_csv(url_MB52_ME,delimiter="\t")

"""----------------------------------------------CLEAN DATA ME-----------------------------------------------------------"""
# Change, drop and standardize columns
Columns_ME=[i.strip() for i in Data_Mb52_ME.columns]   #Alert in list
Data_Mb52_ME=Data_Mb52_ME.set_axis(Columns_ME, axis=1).drop(["Unnamed: 0"],axis=1)  #Drop Unammed: 0
#normalizo estas columnas
Data_Mb52_ME["Libre utilización"]=Data_Mb52_ME["Libre utilización"].apply(lambda x: x.strip())
Data_Mb52_ME["En control calidad"]=Data_Mb52_ME["En control calidad"].apply(lambda x: x.strip())


"""--------------------------------------------COPY ORDENES MB52 MA------------------------------------------------------"""
Ordenes_MB52_MA=Data_Mb52_MA["Material"].drop_duplicates()
componentes=Ordenes_MB52_MA.to_clipboard(index=False, header=False)

"""------------------------------------------------BUSCAR EN ZPP57 MA----------------------------------------------------"""
Transsaccion_ZPP57_MA="zpp57"

SAP_GUI.Search_ZPP57(Transsaccion_ZPP57_MA,Iniciar)

Name_ZPP57_MA="Stock_MA"
SAP_GUI.Export_TXT2(Name_ZPP57_MA,Iniciar)

"""-----------------------------------------MATERIA PRIMA ZPP57 -----------------------------------------------------"""

url_ZPP57_MA="C:\\Users\\prac.ingindustrial2\\OneDrive - Prebel S.A\\Documentos 1\\SAP\\SAP GUI\\"+str(Name_ZPP57_MA)+".txt"
Data_ZPP57_MA=pd.read_csv(url_ZPP57_MA,skiprows=5,delimiter="\t")

# Change, drop and standardize columns
Columns_ZPP57_MA=[i.strip() for i in Data_ZPP57_MA.columns]   #Alert in list
Data_ZPP57_MA=Data_ZPP57_MA.set_axis(Columns_ZPP57_MA,axis=1).drop(["Unnamed: 0","Unnamed: 1"],axis=1)

#Limpio los datos
Data_ZPP57_MA = Data_ZPP57_MA[(Data_ZPP57_MA.Material.notnull())][['Material','Componente','SM']]

"""--------------------------------------------COPY ORDENES MB52 ME------------------------------------------------------"""
Ordenes_MB52_ME=Data_Mb52_ME["Material"].drop_duplicates()
componentes=Ordenes_MB52_ME.to_clipboard(index=False, header=False)

Transsaccion_ZPP57_ME="zpp57"
SAP_GUI.Search_ZPP57(Transsaccion_ZPP57_ME,Iniciar)

Name_ZPP57_ME="Stock_ME"
SAP_GUI.Export_TXT2(Name_ZPP57_ME,Iniciar)

"""-----------------------------------------MATERIA PRIMA ZPP57 -----------------------------------------------------"""

url_ZPP57_ME="C:\\Users\\prac.ingindustrial2\\OneDrive - Prebel S.A\\Documentos 1\\SAP\\SAP GUI\\"+str(Name_ZPP57_ME)+".txt"
Data_ZPP57_ME=pd.read_csv(url_ZPP57_ME,skiprows=5,delimiter="\t")

# Change, drop and standardize columns
Columns_ZPP57_ME=[i.strip() for i in Data_ZPP57_ME.columns]   #Alert in list
Data_ZPP57_ME=Data_ZPP57_ME.set_axis(Columns_ZPP57_ME,axis=1).drop(["Unnamed: 0","Unnamed: 1"],axis=1)

#Limpio los datos
Data_ZPP57_ME = Data_ZPP57_ME[(Data_ZPP57_ME.Material.notnull())][['Material','Componente','SM']]

SAP_GUI.Close_session(Iniciar)

#Dataframe encuentra los que no estan en la MB52

Sin_existencia_MA=pd.merge(Data_Mb52_MA,Data_ZPP57_MA,how="left",left_on="Material",right_on="Componente")
Sin_existencia_MA[['Material_x','Material_y',"Componente"]]

# Filtro de SM usando ZPP57
filter_SM_MA=Data_ZPP57_MA.SM.isin([2,3,5,6,7,8])
# Saco una tabla del anterior filtro por componente
table_MA=Data_ZPP57_MA[filter_SM_MA].Componente


"""------------------------------------------------- Filtro final ---------------------------------------------------"""
#Filtro para material con estado SM enZPP57 
Filter_Not_MB52_MA=Sin_existencia_MA.Material_x.isin(table_MA)

filter_MA=(Sin_existencia_MA.Material_y.isnull() | Filter_Not_MB52_MA)  # Material_x son los que estan en la MB52 pero no estan en la ZPP57
DATA_FINAL_MA=Sin_existencia_MA[filter_MA][['Material_x','Material_y',"Libre utilización","En control calidad","UMB","SM"]].reset_index(drop=True)

"""------------------------------------------- LIMPIO LOS DATOS Y ELIMINA DUPLICADOS-----------------------------------"""
DATA_FINAL_MA=DATA_FINAL_MA.drop("Material_y",axis=1).drop_duplicates()
Filter_stock_MA=(DATA_FINAL_MA["Libre utilización"]!="0")&(DATA_FINAL_MA["Libre utilización"]!="0")
DATA_FINAL_MA=DATA_FINAL_MA[Filter_stock_MA].reset_index(drop=True).sort_values("Libre utilización",ascending=False)

"""------------------------------------------- ESTETICA DEL DATAFRAME MA--------------------------------------------------"""

diccio={2.0:"02",5.0:"05",6.0:"07",7.0:"07"}
column={"Material_x":"Material"}
# Reemplazar los valores en la columna 'valores'
DATA_FINAL_MA["SM"]=DATA_FINAL_MA.SM.replace(diccio)
DATA_FINAL_MA.rename(columns = {'Material_x':'Material'}, inplace = True)


"""------------------------------------------- DATOS PARA MATERIAL DE EMPAQUE-------------------------------------------------"""

#Dataframe encuentra los que no estan en la MB52

Sin_existencia_ME=pd.merge(Data_Mb52_ME,Data_ZPP57_ME,how="left",left_on="Material",right_on="Componente")
Sin_existencia_ME[['Material_x','Material_y',"Componente"]]

# Filtro de SM usando ZPP57
filter_SM_ME=Data_ZPP57_ME.SM.isin([2,3,5,6,7,8])
# Saco una tabla del anterior filtro por componente
table_ME=Data_ZPP57_ME[filter_SM_ME].Componente

"""------------------------------------------------- Filtro final ---------------------------------------------------"""
#Filtro para material con estado SM enZPP57 
Filter_Not_MB52_ME=Sin_existencia_ME.Material_x.isin(table_ME)
filter_ME=(Sin_existencia_ME.Material_y.isnull() | Filter_Not_MB52_ME)  # Material_x son los que estan en la MB52 pero no estan en la ZPP57
DATA_FINAL_ME=Sin_existencia_ME[filter_ME][['Material_x','Material_y',"Libre utilización","En control calidad","UMB","SM"]].reset_index(drop=True)

"""------------------------------------------- LIMPIO LOS DATOS Y ELIMINA DUPLICADOS ME-----------------------------------"""
DATA_FINAL_ME=DATA_FINAL_ME.drop("Material_y",axis=1).drop_duplicates()
Filter_stock_ME=(DATA_FINAL_ME["Libre utilización"]!="0")&(DATA_FINAL_ME["Libre utilización"]!="0")
DATA_FINAL_ME=DATA_FINAL_ME[Filter_stock_ME].sort_values("Libre utilización",ascending=False).reset_index(drop=True)

"""------------------------------------------- ESTETICA DEL DATAFRAME ME--------------------------------------------------"""

diccio={2.0:"02",5.0:"05",6.0:"07",7.0:"07"}
column={"Material_x":"Material"}
# Reemplazar los valores en la columna 'valores'
DATA_FINAL_ME["SM"]=DATA_FINAL_ME.SM.replace(diccio)
DATA_FINAL_ME.rename(columns = {'Material_x':'Material'}, inplace = True)




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

html="""
    <h2 style="text-align: center">REPORTE STOCK MATERIAS PRIMAS</h2>
    <p> Por medio del presente informe se evidencia las materias primas pertenecientes al stock en almacen donde no apararecen
        en la ZPP57 o tiene los siguientes estados: Bloq.material obsoleto, Bloq.material anulado,Fin Abastecimiento, 
        Bloqueo Regulatorio, Obsoleto-Sin Uso, Obsoleto-Banco Productos.</p>

    <div">{0}</div>

    <p> Anticipo sinceros agradecimientos. </p>
 """

html1="""
    <h2 style="text-align: center">REPORTE STOCK MATERIAL DE EMPAQUE</h2>
    <p> Por medio del presente informe se evidencia los materiales de empaque pertenecientes al stock en almacen donde no apararecen
        en la ZPP57 o tiene los siguientes estados: Bloq.material obsoleto, Bloq.material anulado,Fin Abastecimiento, 
        Bloqueo Regulatorio, Obsoleto-Sin Uso, Obsoleto-Banco Productos.</p>
    <div">{0}</div>

    <p> Anticipo sinceros agradecimientos. </p>
 """

Send=style_df(DATA_FINAL_MA)     #Style between LI and LS

Send2=style_df(DATA_FINAL_ME)     #Style between LI and LS

send_emails(Send.to_html(),emails=correos,htmlbody=html,subject="REPORTE MATERIAS PRIMAS PROXIMOS A VENCER")
send_emails(Send2.to_html(),emails=correos,htmlbody=html1,subject="REPORTE MATERIAS PRIMAS PROXIMOS A VENCER")