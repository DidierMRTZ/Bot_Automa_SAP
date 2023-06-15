from Librerias_SAP import SAP_GUI, Funtions
import win32com.client
import pandas as pd
import subprocess
from time import sleep
from datetime import datetime
from datetime import timedelta

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