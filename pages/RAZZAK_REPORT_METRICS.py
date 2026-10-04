from PIL import Image
import numpy as np 
import streamlit as st
import matplotlib.pyplot as plt
st.set_page_config(page_title="RZAZZAK METRICS",page_icon="🛢️",)
import pandas as pd 

from datetime import date, timedelta
from datetime import datetime as dt

Today = date.today()


im = Image.open("KPC.jpg")
image = np.array(im)

st.image(image, width = 150)

st.markdown(" <center>  <h1> RAZZAZK OIL REPORT METRICS </h1> </font> </center> </h1> ",
            unsafe_allow_html=True)








URL = "RAZZAK_OIL_REPORT_22_09_2026.xlsx"


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




import streamlit as st

# Display a simple big number with a label
st.metric(label="Total Users", value="1,245")


st.text_input("Edit shared text:", key="shared_text")

current_x = st.session_state["X"]
current_x 




