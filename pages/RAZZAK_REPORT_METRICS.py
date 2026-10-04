from PIL import Image
import numpy as np 
import streamlit as st
import matplotlib.pyplot as plt
st.set_page_config(page_title="RZAZZAK METRICS",page_icon="🛢️",)
import pandas as pd 

from datetime import date, timedelta
from datetime import datetime as dt

Today = date.today()

Today_Report = Today.strftime('%d_%m_%Y')
YT_Report = dt.strptime(Today_Report, "%d_%m_%Y") - timedelta(days=1)
YT_Report=YT_Report.strftime("%d_%m_%Y")


im = Image.open("KPC.jpg")
image = np.array(im)

imm = Image.open("Apache.jpg")
imagee = np.array(imm)
col1, col2 = st.columns([1, 2])
with col2:
    st.image(imagee, width = 150)
with col1:
    st.image(image, width = 150)
  
Deduction = []

URL = "RAZZAK_OIL_REPORT_"+"22-_09_2026"+'.xlsx'
url = "RAZZAK_OIL_REPORT_"+"21-_09_2026"+'.xlsx'


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
st.markdown(" <center>  <h1> RAZZAZK OIL REPORT ANALYSIS </h1> </font> </center> </h1> ",
            unsafe_allow_html=True)







