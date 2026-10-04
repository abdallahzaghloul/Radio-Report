from PIL import Image
import numpy as np 
import streamlit as st
import matplotlib.pyplot as plt
st.set_page_config(page_title="RZAZZAK METRICS",page_icon="🛢️",)
import pandas as pd 

from datetime import date, timedelta
from datetime import datetime as dt
st.metric(label="Total Users", value="1,245")

Today = date.today()


im = Image.open("KPC.jpg")
image = np.array(im)

st.image(image, width = 150)

st.markdown(" <center>  <h1> RAZZAZK OIL REPORT METRICS </h1> </font> </center> </h1> ",
            unsafe_allow_html=True)









#st.session_state["RRE"]=RAZZAK_REPORT_Evalulation
#st.session_state["OnLine"]=Well_Data[(Well_Data['STATUS']!='SI')].shape[0]
#st.session_state["SI"]=Well_Data[(Well_Data['STATUS']=='SI')].shape[0]
#st.session_state["DRR"]=URL
#st.session_state["YRR"]=url
#URL = "RAZZAK_OIL_REPORT_22_09_2026.xlsx"

RRE = st.session_state["RRE"]
URL = st.session_state["DRR"]
url = st.session_state["YRR"]
SI_Wells_No = st.session_state["SI"]
On_Line_Wells_No = st.session_state["OnLine"]


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








