from PIL import Image
import streamlit as st
import numpy as np 
import pandas as pd 

from datetime import date, timedelta
from datetime import datetime as dt

Today = date.today()

im = Image.open("KPC.jpg")
image = np.array(im)

st.image(image,width =150)

url="RAZZAK_OIL_REPORT_21_SEP_26.xlsx"
URL="RAZZAK_OIL_REPORT_22_SEP_26.xlsx"

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

TANKS = pd.read_excel(URL, sheet_name="TANKS", header = None)
TANKSS =pd.read_excel(url, sheet_name="TANKS", header = None)

WF= pd.read_excel(URL, sheet_name="WATER FLOOD WELLS")

MRZK_INJ=WF.iloc[97,8]
WRZK_INJ=WF.iloc[98,8]
ERZK_INJ=WF.iloc[99,8]
NRQ_INJ=WF.iloc[100,8]


MRZK_REQ=WF.iloc[97,9]
WRZK_REQ=WF.iloc[98,9]
ERZK_REQ=WF.iloc[99,9]
NRQ_REQ=WF.iloc[100,9]

# Variance = ((Actual - Forecast) / Forecast) * 100

if abs((MRZK_INJ-MRZK_REQ)/MRZK_REQ) < .05:
  print("Allah Allah Ya Wla")
elif abs((MRZK_INJ-MRZK_REQ)/MRZK_REQ) > .05:
  print("5od ya koskos")

if abs((WRZK_INJ-WRZK_REQ)/WRZK_REQ) < .05:
  print("Allah Allah Ya Wla")
elif abs((WRZK_INJ-WRZK_REQ)/WRZK_REQ) > .05:
  print("5od ya koskos")

if abs((ERZK_INJ-ERZK_REQ)/ERZK_REQ) < .05:
  print("Allah Allah Ya Wla")
elif abs((ERZK_INJ-ERZK_REQ)/ERZK_REQ) > .05:
  print("5od ya koskos")

if abs((NRQ_INJ-NRQ_REQ)/NRQ_REQ) < .05:
  print("Allah Allah Ya Wla")
elif abs((NRQ_INJ-NRQ_REQ)/NRQ_REQ) > .05:
  print("5od ya koskos")

Well_Data= pd.read_excel(URL, sheet_name="WELL DATA",usecols='B:AP',skiprows=1, nrows=145)

Well_Data.columns=Well_Data.columns.str.upper()
Well_Data.columns=Well_Data.columns.str.replace(".1","")
Well_Data.columns=Well_Data.columns.str.replace("\n","")
Well_Data.columns=Well_Data.columns.str.strip()
Well_Data=Well_Data.drop([0,1], axis=0)
Well_Data.index = range(1, len(Well_Data) + 1)


#(Well_Data['STATUS']=='SI').sum()

#(Well_Data['STATUS']!='SI').sum()


#(Well_Data['STATUS']).count()

#(Well_Data['STATUS']=='SI').sum() + (Well_Data['STATUS']!='SI').sum() == (Well_Data['STATUS']).count()


#Well_Data.columns


Well_Data['LAST WELL TEST']=pd.to_datetime(Well_Data['LAST WELL TEST'])
Well_Data['LAST WELL TEST']=Well_Data['LAST WELL TEST'].dt.strftime('%d-%m-%Y')

Tested_Today= int(Well_Data[Well_Data['LAST WELL TEST']=='22-09-2026']['LAST WELL TEST'].count())

Tested_Today_Per = float(Tested_Today/(Well_Data['STATUS']!='SI').sum())*100



st.markdown(" <center>  <h1> RAZZAZK OIL REPORT ANALYSIS </h1> </font> </center> </h1> ",
            unsafe_allow_html=True)


st.markdown(" <right>  <h1> Alerts </h1> </font> </right> </h1> ",
            unsafe_allow_html=True)
st.write("Hello")


st.markdown('<p style="color:blue; font-size:24px;">This is blue and 24 pixels big!</p>',
    unsafe_allow_html=True)

df = pd.DataFrame(
    {"A": [10, 20, 30], "B": [15, 25, 35], "C": [100, 200, 300]}
)

# Apply a solid blue background with white text
styled_df = df.style.set_properties(
    **{"background-color": "#1E3A8A", "color": "white"}
)

# Render in Streamlit
st.dataframe(styled_df)




import streamlit as st

st.set_page_config(page_title="My Cool App",page_icon="🎈",)


# Evaluation of Oil Variance

Oil_Prod=pd.read_excel(URL, sheet_name="REPORT", header = None)
Oil_Var= int(Oil_Prod.iloc[17,7]-Oil_Prod.iloc[17,6])/int(Oil_Prod.iloc[17,7])
Oil_Var_Per= abs(int(Oil_Prod.iloc[17,7]-Oil_Prod.iloc[17,6])/int(Oil_Prod.iloc[17,7]))*100 

Evalution = []
if Oil_Var_Per<= 5:
  Evaluation.append(2)
else:
  Evaluation = [10]

# Tanks Evaluation 




