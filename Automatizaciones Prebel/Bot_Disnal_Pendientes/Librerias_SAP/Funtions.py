import re
import json

"""----------------------------------------------CLEAN COLUMN AND DATA-----------------------------------------------------------"""
def Clean_Columns(List):
    Set_Columns=[i.strip() for i in List.columns]   #Alert in list
    clean_column=[i if "Unnamed" in i else None for i in List.columns]
    clean_column = list(filter(lambda x: x is not None, clean_column)) 
    List=List.set_axis(Set_Columns, axis=1).drop(clean_column,axis=1)  #Drop Unammed: 0
    return(List)

"""------------------------------------------FUNCION PARA CAMBIAR A NUMEROS-------------------------------------------------------"""

def Clean_num(x):
    x = float(str(x).strip().replace(',',''))
    return(x)

"""------------------------------------------FUNCION PARA COMPLETAR los 10 con 00 al inicio---------------------------------------"""
def Complete_00(lista):
    Nueva_Lista_Ordenes=[]
    for i in lista:
        while len(str(i))<10:
            i="0"+i
            # print(len(str(i)),i)
        Nueva_Lista_Ordenes.append(str(i))
    return(Nueva_Lista_Ordenes)

def default_column(default_columns,dataframe):     #Parametros (default_columns: Columnas predeterminadas,dataframe:)
    diccionay_default_column={}
    if len(default_columns)==len(dataframe.columns):
        for i in range(0,len(default_columns)):
            diccionay_default_column[dataframe.columns[i]]=default_columns[i]   
        dataframe=dataframe.rename(columns=diccionay_default_column)   #Remplazo las columnas con las de default
    else:
        None   #asumo columnas originales como estandar
    return(dataframe)


#Limpiar colummna a numeros
def Clean_Num_List(*args):
    lista=[]
    for arg in args:
        args=arg.apply(lambda x: float(str(x).strip().replace(',','')))
        lista.append(args)
    if len(lista)==1:
        return lista[0]
    else:
        return(tuple(lista))

#Limpiar colummna a numeros
def Clean_int_to_str(*args):
    lista=[]
    for arg in args:
        args=arg.astype(int).astype(str)
        lista.append(args)
    if len(lista)==1:
        return lista[0]
    else:
        return(tuple(lista))


# Ecluir Agenda de pedidos exito

def Datos_Agenda(data_pedido,Agenda):   #(data_pedido: columna dataframe a transformar, Agenda: Datos a encontrar)
    Exluidos_Entrega=set()
    for i in data_pedido:
        if (str(i)[:4] in Agenda) or (str(i)[:3] in Agenda):
            Exluidos_Entrega.add(i)
    return(Exluidos_Entrega)


def Search_agenda(data_Pedidos,agenda):
    conjunto_agenda=set()
    for i in data_Pedidos:
        if re.findall("(\d*)-",str(i))!=[] and (re.findall("(\d*)-",str(i))[0] in agenda):
            conjunto_agenda.add(i)
    return(conjunto_agenda)

# lista json de channles

def list_to_json(List_Channels,path=None):
    """
    - List_Channels: Recibe lista de canales
    - path: ruta del archivo .json para exportar (Default: None no exporta)
    """
    dic={}
    for i in List_Channels: dic["Channels "+i]=i
    if path==None:
        return(json.dumps(dic))
    else:
        with open(path, "w") as archivo:
            # Escribir datos en formato JSON
            json.dump(dic, archivo)
        return(json.dumps(dic))