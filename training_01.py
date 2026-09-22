
from PIL import Image
import streamlit as st
import numpy as np 
import pandas as pd 
from datetime import date, datetime, timedelta

Today = date.today()

URL="RAZZAK_OIL_REPORT_21_SEP_26.xlsx"


RAZZAK_File=pd.ExcelFile(URL)
Excel_Sheets=RAZZAK_File.sheet_names
Sheet_Check_List=[]
for i in range (1,len(Excel_Sheets)):
  try:
    pd.read_excel(URL, sheet_name=Excel_Sheets[i])
    A=True
  except:
    A=False;
  Sheet_Check_List.append(A)
Excel_Check_Report= pd.DataFrame(Excel_Sheets,columns=["Excel Sheets"])
Excel_Check_Report["Sheet Existence"]= pd.Series(Sheet_Check_List)
Excel_Check_Report

df_REPORT=pd.read_excel(URL, sheet_name="REPORT",usecols='AU:BB',skiprows=(2))
df_REPORT.columns=df_REPORT.columns.str.upper()
df_REPORT.columns=df_REPORT.columns.str.replace(".1","")
df_REPORT.columns=df_REPORT.columns.str.replace("\n","")
df_REPORT.columns=df_REPORT.columns.str.strip()
df_REPORT.index = range(1, len(df_REPORT) + 1)

df_WT=pd.read_excel(URL, sheet_name="WELL TESTS",usecols='AO:BH',skiprows=1)
df_WT.columns=df_WT.columns.str.upper()
df_WT.columns=df_WT.columns.str.replace(".1","")
df_WT.columns=df_WT.columns.str.replace("\n","")
df_WT.columns=df_WT.columns.str.strip()
df_WT.index = range(1, len(df_WT) + 1)

df_WD=pd.read_excel(URL, sheet_name="WELL DATA",usecols='CW:DR',skiprows=3)
df_WD.columns=df_WD.columns.str.upper()
df_WD.columns=df_WD.columns.str.replace(".1","")
df_WD.columns=df_WD.columns.str.replace("\n","")
df_WD.columns=df_WD.columns.str.strip()
df_WD.index = range(1, len(df_WD) + 1)


df_WFW=pd.read_excel(URL, sheet_name="WATER FLOOD WELLS",usecols='AY:BN',skiprows=1)
df_WFW.columns=df_WFW.columns.str.upper()
df_WFW.columns=df_WFW.columns.str.replace(".1","")
df_WFW.columns=df_WFW.columns.str.replace("\n","")
df_WFW.columns=df_WFW.columns.str.strip()
df_WFW=df_WFW.drop([0],axis=0)
df_WFW.index = range(1, len(df_WFW) + 1)

df_T=pd.read_excel(URL, sheet_name="TANKS",usecols='B:N',skiprows=1)
df_T.columns=df_T.columns.str.upper()
df_T.columns=df_T.columns.str.replace(".1","")
df_T.columns=df_T.columns.str.replace("\n","")
df_T.columns=df_T.columns.str.strip()
df_T.index = range(1, len(df_T) + 1)

df_DFL=pd.read_excel(URL, sheet_name="DFL SHOTS",usecols='Z:AJ',skiprows=1)
df_DFL.columns=df_DFL.columns.str.upper()
df_DFL.columns=df_DFL.columns.str.replace(".1","")
df_DFL.columns=df_DFL.columns.str.replace("\n","")
df_DFL.columns=df_DFL.columns.str.strip()
df_DFL.index = range(1, len(df_DFL) + 1)


df_Data_Input=pd.read_excel(URL, sheet_name="DATA INPUT",usecols='B:I',skiprows=1)
df_Data_Input.columns=df_Data_Input.columns.str.upper()
df_Data_Input.columns=df_Data_Input.columns.str.replace(".1","")
df_Data_Input.columns=df_Data_Input.columns.str.replace("\n","")
df_Data_Input.columns=df_Data_Input.columns.str.strip()
df_Data_Input.index = range(1, len(df_Data_Input) + 1)

df_List=pd.read_excel(URL, sheet_name="WELL LIST",usecols='A:B',skiprows=1)
df_List.columns=df_List.columns.str.upper()
df_List.columns=df_List.columns.str.replace(".1","")
df_List.columns=df_List.columns.str.replace("\n","")
df_List.columns=df_List.columns.str.strip()
df_List.index = range(1, len(df_List) + 1)

st.markdown(" <center>  <h1> RAZZAZK OIL REPORT ANALYSIS </h1> </font> </center> </h1> ",
            unsafe_allow_html=True)

st.write("Hello")
Oil_Prod=pd.read_excel(URL, sheet_name="REPORT", header = None)
Oil_Prod.iloc[17,7]
Oil_Prod.iloc[17,6]
Oil_Var= int(Oil_Prod.iloc[17,7]-Oil_Prod.iloc[17,6])/int(Oil_Prod.iloc[17,7])

Oil_Var*100
if Oil_Var*100 <4.99:
  st.write("Yallahwy El7gny Ya 3my Mansour")



