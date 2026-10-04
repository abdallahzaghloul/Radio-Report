from PIL import Image
import numpy as np 
import streamlit as st
import matplotlib.pyplot as plt
st.set_page_config(page_title="RZAZZAK REPORT ",page_icon="🛢️",)
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


if st.button("Default RAZZAK Report Dates"):    
    URL = "RAZZAK_OIL_REPORT_"+Today_Report+'.xlsx'
    url = "RAZZAK_OIL_REPORT_"+YT_Report+'.xlsx'




Start_Date = st.text_input("Enter Start Date [30-08-2025]")
Start_Date =Start_Date.replace('-','_') 


End_Date = st.text_input("Enter End Date [31-08-2025]")
End_Date =End_Date.replace('-','_')
URL = "RAZZAK_OIL_REPORT_"+End_Date+'.xlsx'


if st.button("Selective RAZZAK Report Dates"):
    URL = "RAZZAK_OIL_REPORT_"+End_Date+'.xlsx'
    url = "RAZZAK_OIL_REPORT_"+Start_Date+'.xlsx'


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

#df_REPORT=pd.read_excel(URL, sheet_name="REPORT",usecols='AU:BB',skiprows=(2))
#df_REPORT.columns=df_REPORT.columns.str.upper()
#df_REPORT.columns=df_REPORT.columns.str.replace(".1","")
#df_REPORT.columns=df_REPORT.columns.str.replace("\n","")
#df_REPORT.columns=df_REPORT.columns.str.strip()
#df_REPORT.index = range(1, len(df_REPORT) + 1)

#df_WT=pd.read_excel(URL, sheet_name="WELL TESTS",usecols='AO:BH',skiprows=1)
#df_WT.columns=df_WT.columns.str.upper()
#df_WT.columns=df_WT.columns.str.replace(".1","")
#df_WT.columns=df_WT.columns.str.replace("\n","")
#df_WT.columns=df_WT.columns.str.strip()
#df_WT.index = range(1, len(df_WT) + 1)

#df_WD=pd.read_excel(URL, sheet_name="WELL DATA",usecols='CW:DR',skiprows=3)
#df_WD.columns=df_WD.columns.str.upper()
#df_WD.columns=df_WD.columns.str.replace(".1","")
#df_WD.columns=df_WD.columns.str.replace("\n","")
#df_WD.columns=df_WD.columns.str.strip()
#df_WD.index = range(1, len(df_WD) + 1)


#df_WFW=pd.read_excel(URL, sheet_name="WATER FLOOD WELLS",usecols='AY:BN',skiprows=1)
#df_WFW.columns=df_WFW.columns.str.upper()
#df_WFW.columns=df_WFW.columns.str.replace(".1","")
#df_WFW.columns=df_WFW.columns.str.replace("\n","")
#df_WFW.columns=df_WFW.columns.str.strip()
#df_WFW=df_WFW.drop([0],axis=0)
#df_WFW.index = range(1, len(df_WFW) + 1)

#df_T=pd.read_excel(URL, sheet_name="TANKS",usecols='B:N',skiprows=1)
#df_T.columns=df_T.columns.str.upper()
#df_T.columns=df_T.columns.str.replace(".1","")
#df_T.columns=df_T.columns.str.replace("\n","")
#df_T.columns=df_T.columns.str.strip()
#df_T.index = range(1, len(df_T) + 1)

#df_DFL=pd.read_excel(URL, sheet_name="DFL SHOTS",usecols='Z:AJ',skiprows=1)
#df_DFL.columns=df_DFL.columns.str.upper()
#df_DFL.columns=df_DFL.columns.str.replace(".1","")
#df_DFL.columns=df_DFL.columns.str.replace("\n","")
#df_DFL.columns=df_DFL.columns.str.strip()
#df_DFL.index = range(1, len(df_DFL) + 1)


#df_Data_Input=pd.read_excel(URL, sheet_name="DATA INPUT",usecols='B:I',skiprows=1)
#df_Data_Input.columns=df_Data_Input.columns.str.upper()
#df_Data_Input.columns=df_Data_Input.columns.str.replace(".1","")
#df_Data_Input.columns=df_Data_Input.columns.str.replace("\n","")
#df_Data_Input.columns=df_Data_Input.columns.str.strip()
#df_Data_Input.index = range(1, len(df_Data_Input) + 1)

#df_List=pd.read_excel(URL, sheet_name="WELL LIST",usecols='A:B',skiprows=1)
#df_List.columns=df_List.columns.str.upper()
#df_List.columns=df_List.columns.str.replace(".1","")
#df_List.columns=df_List.columns.str.replace("\n","")
#df_List.columns=df_List.columns.str.strip()
#df_List.index = range(1, len(df_List) + 1)

TANKS = pd.read_excel(URL, sheet_name="TANKS", header = None)
TANKSS =pd.read_excel(url, sheet_name="TANKS", header = None)


Well_Data= pd.read_excel(URL, sheet_name="WELL DATA",usecols='B:AP',skiprows=1, nrows=145)
Well_Data.columns=Well_Data.columns.str.upper()
Well_Data.columns=Well_Data.columns.str.replace(".1","")
Well_Data.columns=Well_Data.columns.str.replace("\n","")
Well_Data.columns=Well_Data.columns.str.strip()
Well_Data=Well_Data.drop([0,1], axis=0)
Well_Data.index = range(1, len(Well_Data) + 1)

Well_Data_SRP = Well_Data[Well_Data['STATUS']=="SRP"]
Well_Data_SRP=Well_Data_SRP.drop(columns=['AMPS', 'HZ', 'MOTOR RATING AMP','WELL TYPE', 'MOTOR LOADING PERCENTAGE','EST PRODN','UNNAMED: 19','UNNAMED: 37','DOWNTIME CAUSEINPUT WILL BE BLUE IF CAUSE REQUIRED'])
Well_Data_SRP_Null=int(Well_Data_SRP.isnull().sum().sum())
Well_Data['HRS ONLINE']=Well_Data['HRS ONLINE'].astype(float)
Well_Data_24=Well_Data[(Well_Data['HRS ONLINE']<24) & (Well_Data['STATUS']!='SI') ]
Well_Data_24_Null=int(Well_Data_24['DOWNTIME CAUSEINPUT WILL BE BLUE IF CAUSE REQUIRED'].isnull().sum())

Well_Data_ESP = Well_Data[Well_Data['STATUS']=="ESP"]
Well_Data_ESP=Well_Data_ESP.drop(columns=['UNNAMED: 19','UNNAMED: 37','DOWNTIME CAUSEINPUT WILL BE BLUE IF CAUSE REQUIRED','FLUID LEVELS', 'UNNAMED: 16', 'UNNAMED: 17', 'UNNAMED: 18',
       'UNNAMED: 19', 'PUMP DEPTH', 'PUMP SIZE "', 'DAILY SAM DATA',
       'UNNAMED: 23', 'UNNAMED: 24', 'UNNAMED: 25', 'SUCKER ROD PUMP DATA',
       'UNNAMED: 27', 'UNNAMED: 28', 'UNNAMED: 29'])
Well_Data_ESP_Null=int(Well_Data_ESP.isnull().sum().sum())
Well_Data_Other = Well_Data[(Well_Data['STATUS']!="ESP") & (Well_Data['STATUS']!="SRP")&(Well_Data['STATUS']!="SI") ]
Well_Data_Other=Well_Data_Other.drop(columns=['UNNAMED: 19','UNNAMED: 37','DOWNTIME CAUSEINPUT WILL BE BLUE IF CAUSE REQUIRED','FLUID LEVELS', 'UNNAMED: 16', 'UNNAMED: 17', 'UNNAMED: 18',
       'UNNAMED: 19', 'PUMP DEPTH', 'PUMP SIZE "', 'DAILY SAM DATA',
       'UNNAMED: 23', 'UNNAMED: 24', 'UNNAMED: 25', 'SUCKER ROD PUMP DATA',
       'UNNAMED: 27', 'UNNAMED: 28', 'UNNAMED: 29','AMPS', 'HZ', 'MOTOR RATING AMP','WELL TYPE', 'MOTOR LOADING PERCENTAGE','EST PRODN'])
Well_Data_Other_Null=int(Well_Data_Other.isnull().sum().sum())








Well_Data_24_Null_Per=Well_Data_24_Null/Well_Data_24.shape[0]
Well_Data_Other_Null_Per= Well_Data_Other_Null / (Well_Data_Other.shape[0] * Well_Data_Other.shape[1])
Well_Data_SRP_Null_Per = Well_Data_SRP_Null / (Well_Data_SRP.shape[0] * Well_Data_SRP.shape[1])
Well_Data_ESP_Null_Per = Well_Data_ESP_Null / (Well_Data_ESP.shape[0] * Well_Data_ESP.shape[1])
Well_Data_Comp=100-(Well_Data_ESP_Null_Per+Well_Data_24_Null_Per+Well_Data_SRP_Null_Per+Well_Data_Other_Null_Per)*100



    
WF_Data= pd.read_excel(URL, sheet_name="WATER FLOOD WELLS",usecols='B:S',skiprows=1, nrows=74)
WF_Data['HRS ONLINE']=WF_Data['HRS ONLINE'].astype(float)
WF_Data=WF_Data.drop([0])

WF_Data_INJ = WF_Data[WF_Data['STATUS']=='INJ']
WF_Data_INJ = WF_Data_INJ.drop(columns=['DOWNTIME CAUSE'])

WF_Data_INJ.head(2)
WF_Data_INJ_Null=int(WF_Data_INJ.isnull().sum().sum())


WF_Data_24 = WF_Data[(WF_Data['STATUS']!='SI')&(WF_Data['HRS ONLINE']<24)]
WF_Data_24_Null=int(WF_Data_24['DOWNTIME CAUSE'].isnull().sum())

WF_Data_24_Null_Per = WF_Data_24_Null / WF_Data_24.shape[0]
WF_Data_INJ_Null_Per = WF_Data_INJ_Null/ (WF_Data_INJ.shape[0]*WF_Data_INJ.shape[1])
WF_Comp = 100-(WF_Data_INJ_Null_Per+WF_Data_24_Null_Per)*100


st.markdown(" <center>  <h1> RAZZAZK OIL REPORT ANALYSIS </h1> </font> </center> </h1> ",
            unsafe_allow_html=True)


st.markdown(" <center>  <h1> Oil Variance Validation </h1> </font> </center> </h1> ",  unsafe_allow_html=True)


#st.markdown('<p style="color:blue; font-size:24px;">This is blue and 24 pixels big!</p>',    unsafe_allow_html=True)

# Oil Variance Alerts
Oil_Prod=pd.read_excel(URL, sheet_name="REPORT", header = None)
Oil_Var= int(Oil_Prod.iloc[17,7]-Oil_Prod.iloc[17,6])/int(Oil_Prod.iloc[17,7])
Oil_Var_Per= abs(int(Oil_Prod.iloc[17,7]-Oil_Prod.iloc[17,6])/int(Oil_Prod.iloc[17,7]))*100

if Oil_Var_Per==0:
    df = pd.DataFrame({"Excellent Oil Variance": ['The Oil Variance %',Oil_Var_Per]})
    styled_df = df.style.set_properties(**{"background-color": "#008000", "color": "white"})
    st.dataframe(styled_df)
    Oil_Var_Alert = 0

elif Oil_Var_Per<= 5:
    df = pd.DataFrame({"High Variance Warning": ['The Oil Variance %',Oil_Var_Per]})
    styled_df = df.style.set_properties(**{"background-color": "#E1AD01", "color": "white"})
    st.dataframe(styled_df)
    Oil_Var_Alert = 5
else:
    df = pd.DataFrame({"Critical Alerts": ['The Oil Variance %',Oil_Var_Per]})
    styled_df = df.style.set_properties(**{"background-color": "#B30E08", "color": "white"})
    st.dataframe(styled_df)
    Oil_Var_Alert = 10


Deduction.append(Oil_Var_Alert)


st.markdown(" <center>  <h1> Completeness Validation </h1> </font> </center> </h1> ",  unsafe_allow_html=True)

 
if Well_Data_Comp >=80:
    df = pd.DataFrame({"Excellent Data Comleteness": ['Well Data Completeness',Well_Data_Comp]})
    styled_df = df.style.set_properties(**{"background-color": "#008000", "color": "white"})
    st.dataframe(styled_df)
    Well_Comp_Alert = 0

elif Well_Data_Comp <80:
    df = pd.DataFrame({"Quality Issues Alerts": ['Well Data Completeness',Well_Data_Comp]})
    styled_df = df.style.set_properties(**{"background-color": "#8FD9FB", "color": "white"})
    st.dataframe(styled_df)
    Well_Comp_Alert = 2


if WF_Comp >=80:
    df = pd.DataFrame({"Excellent Data Comleteness": ['Water Flood Completeness',WF_Comp]})
    styled_df = df.style.set_properties(**{"background-color": "#008000", "color": "white"})
    st.dataframe(styled_df)
    WF_Comp_Alert = 0

elif WF_Comp <80:
    df = pd.DataFrame({"Quality Issues Alerts": ['Water Flood Completeness',WF_Comp]})
    styled_df = df.style.set_properties(**{"background-color": "#8FD9FB", "color": "white"})
    st.dataframe(styled_df)
    WF_Comp_Alert = 2

if WF_Comp_Alert==2 or Well_Comp_Alert==2:
    Deduction.append(2)
elif WF_Comp_Alert==0 and Well_Comp_Alert==0:
    Deduction.append(0)





st.markdown(" <center>  <h1> Tank Validation </h1> </font> </center> </h1> ",  unsafe_allow_html=True)

T_201_Today = pd.read_excel(URL, sheet_name="TANKS")
T_201_YT =pd.read_excel(url, sheet_name="TANKS")

T_201_Today_Val = float(T_201_Today.iloc[1,3].replace(',','.'))
T_201_YT_Val = float(T_201_YT.iloc[2,3].replace(',','.'))
T_201_Var = (abs(T_201_Today_Val-T_201_YT_Val)/T_201_Today_Val)*100


if T_201_Var == 0:
    df = pd.DataFrame({"Excellent Tank-201 Variance": ['Tank-201 Variance',T_201_Var]})
    styled_df = df.style.set_properties(**{"background-color": "#008000", "color": "white"})
    st.dataframe(styled_df)
    T_201_Alert = 0
elif T_201_Var<=.1:
    df = pd.DataFrame({"Excellent Tank-201 Variance": ['Tank-201 Variance',T_201_Var]})
    styled_df = df.style.set_properties(**{"background-color": "#305CDE", "color": "white"})
    st.dataframe(styled_df)
    T_201_Alert = 0
else:
    df = pd.DataFrame({"Quality Issue Tank-201 Variance": ['Tank-201',T_201_Var]})
    styled_df = df.style.set_properties(**{"background-color": "#305CDE", "color": "white"})
    st.dataframe(styled_df)
    T_201_Alert = 2
    
T_202_Today = pd.read_excel(URL, sheet_name="TANKS")
T_202_YT =pd.read_excel(url, sheet_name="TANKS")

T_202_Today_Val = float(T_202_Today.iloc[3,3].replace(',','.'))
T_202_YT_Val = float(T_202_YT.iloc[4,3].replace(',','.'))
T_202_Var = (abs(T_202_Today_Val-T_202_YT_Val)/T_202_Today_Val)*100

if T_202_Var == 0:
    df = pd.DataFrame({"Excellent Tank-202 Variance": ['Tank-202 Variance',T_202_Var]})
    styled_df = df.style.set_properties(**{"background-color": "#008000", "color": "white"})
    st.dataframe(styled_df)
    T_202_Alert = 0
elif T_202_Var<=.1:
    df = pd.DataFrame({"Excellent Tank-202 Variance": ['Tank-202 Variance',T_202_Var]})
    styled_df = df.style.set_properties(**{"background-color": "#008000", "color": "white"})
    st.dataframe(styled_df)
    T_202_Alert = 0
else:
    df = pd.DataFrame({"Quality Issue Tank-202 Variance": ['Tank-202',T_202_Var]})
    styled_df = df.style.set_properties(**{"background-color": "#305CDE", "color": "white"})
    st.dataframe(styled_df)
    T_202_Alert = 2
if T_201_Alert ==2 or T_202_Alert ==2:
    Deduction.append(2)
elif T_201_Alert ==0 or T_202_Alert ==0:
    Deduction.append(0)



st.markdown(" <center>  <h1> Test Coverage Validation </h1> </font> </center> </h1> ",  unsafe_allow_html=True)

Well_Data['LAST WELL TEST']=pd.to_datetime(Well_Data['LAST WELL TEST'])
Well_Data['LAST WELL TEST']=Well_Data['LAST WELL TEST'].dt.strftime('%d-%m-%Y')

Today = date.today()
Today = Today.strftime('%d-%m-%Y')

#Tested_Today= int(Well_Data[Well_Data['LAST WELL TEST']==Today]['LAST WELL TEST'].count())
Tested_Today= int(Well_Data[Well_Data['LAST WELL TEST']=='22-09-2026']['LAST WELL TEST'].count())

Tested_Today_Per = float(Tested_Today/(Well_Data['STATUS']!='SI').sum())*100





if Tested_Today_Per >= 25:
    df = pd.DataFrame({"Excellent Test Coverage": ['Covered Tests Today',Tested_Today_Per]})
    styled_df = df.style.set_properties(**{"background-color": "#008000", "color": "white"})
    st.dataframe(styled_df)
    Deduction.append(0)
elif Tested_Today_Per<25:
    df = pd.DataFrame({"Quality Issue": ['Covered Tests Today',Tested_Today_Per]})
    styled_df = df.style.set_properties(**{"background-color": "#305CDE", "color": "white"})
    st.dataframe(styled_df)
    Deduction.append(0)

st.markdown(" <center>  <h1> Water Flood Validation </h1> </font> </center> </h1> ",  unsafe_allow_html=True)

WF= pd.read_excel(URL, sheet_name="WATER FLOOD WELLS")

MRZK_INJ_REQ=WF.iloc[97,10]
WRZK_INJ_REQ=WF.iloc[98,10]
ERZK_INJ_REQ=WF.iloc[99,10]
NRQ_INJ_REQ=WF.iloc[100,10]

MRZK_REQ=WF.iloc[97,9]
WRZK_REQ=WF.iloc[98,9]
ERZK_REQ=WF.iloc[99,9]
NRQ_REQ=WF.iloc[100,9]

MRZK_WF_Var = abs((MRZK_INJ_REQ)/MRZK_REQ)*100
WRZK_WF_Var = abs((WRZK_INJ_REQ)/WRZK_REQ)*100
ERZK_WF_Var = abs((ERZK_INJ_REQ)/ERZK_REQ)*100
NRQ_WF_Var = abs((NRQ_INJ_REQ)/NRQ_REQ)*100
if MRZK_WF_Var <= 5: 
    df = pd.DataFrame({"Excellent INJ Vs. REQ.": ['MRZK INJ Validation',MRZK_WF_Var]})
    styled_df = df.style.set_properties(**{"background-color": "#008000", "color": "white"})
    st.dataframe(styled_df)
    MRZK_WF_Alert=0
elif MRZK_WF_Var > 5:
    df = pd.DataFrame({"Quality Issue": ['MRZK INJ Validation',MRZK_WF_Var]})
    styled_df = df.style.set_properties(**{"background-color": "#305CDE", "color": "white"})
    st.dataframe(styled_df)
    MRZK_WF_Alert=2
if ERZK_WF_Var <= 5: 
    df = pd.DataFrame({"Excellent INJ Vs. REQ.": ['ERZK INJ Validation',ERZK_WF_Var]})
    styled_df = df.style.set_properties(**{"background-color": "#008000", "color": "white"})
    st.dataframe(styled_df)
    ERZK_WF_Alert=0
elif ERZK_WF_Var > 5:
    df = pd.DataFrame({"Quality Issue": ['ERZK INJ Validation',ERZK_WF_Var]})
    styled_df = df.style.set_properties(**{"background-color": "#305CDE", "color": "white"})
    st.dataframe(styled_df)
    ERZK_WF_Alert=2

if WRZK_WF_Var <= 5: 
    df = pd.DataFrame({"Excellent INJ Vs. REQ.": ['WRZK INJ Validation',WRZK_WF_Var]})
    styled_df = df.style.set_properties(**{"background-color": "#008000", "color": "white"})
    st.dataframe(styled_df)
    WRZK_WF_Alert=0
elif WRZK_WF_Var > 5:
    df = pd.DataFrame({"Quality Issue": ['WRZK INJ Validation',WRZK_WF_Var]})
    styled_df = df.style.set_properties(**{"background-color": "#305CDE", "color": "white"})
    st.dataframe(styled_df)
    WRZK_WF_Alert=2

if NRQ_WF_Var <= 5: 
    df = pd.DataFrame({"Excellent INJ Vs. REQ.": ['NRQ INJ Validation',NRQ_WF_Var]})
    styled_df = df.style.set_properties(**{"background-color": "#008000", "color": "white"})
    st.dataframe(styled_df)
    NRQ_WF_Alert=0
elif NRQ_WF_Var > 5:
    df = pd.DataFrame({"Quality Issue": ['NRQ INJ Validation',NRQ_WF_Var]})
    styled_df = df.style.set_properties(**{"background-color": "#305CDE", "color": "white"})
    st.dataframe(styled_df)
    NRQ_WF_Alert=2

WF_Alert = NRQ_WF_Alert+ MRZK_WF_Alert+ ERZK_WF_Alert+ WRZK_WF_Alert
if WF_Alert>0:
    Deduction.append(0)
elif WF_Alert==0:
    Deduction.append(2)

RAZZAK_REPORT_Evalulation =100-sum(Deduction)


if RAZZAK_REPORT_Evalulation>=95:
    colors=['#636B2F','#000000' ]

elif RAZZAK_REPORT_Evalulation>=85:
    colors=['#FFCE1B','#000000' ]
elif RAZZAK_REPORT_Evalulation>=75:
    colors=['#8FD9FB','#000000' ]
else:
    colors=['#880808','#000000' ]


plt.figure(figsize=(6, 6)) # Sets a square figure size
plt.title('RAZZAK REPORT EVALUATION TODAY')

RAZZAZK_REPORT_PI_CHART = [RAZZAK_REPORT_Evalulation,sum(Deduction)]


WSW_Data= pd.read_excel(URL, sheet_name="WATER FLOOD WELLS",usecols='B:S',skiprows=81, nrows=14)
WSW_Data



fig, ax = plt.subplots()
ax.pie(RAZZAZK_REPORT_PI_CHART,labels = ['RAZZAZK Val.','Deduction'],colors=colors, autopct='%1.1f%%', shadow=True,             startangle=140       )
ax.axis('equal')
st.pyplot(fig)
#plt.show()


st.session_state["RRE"]=RAZZAK_REPORT_Evalulation
st.session_state["OnLine"]=Well_Data[(Well_Data['STATUS']!='SI')].shape[0]
st.session_state["SI"]=Well_Data[(Well_Data['STATUS']=='SI')].shape[0]
st.session_state["DRR"]=URL
st.session_state["YRR"]=url
st.session_state["OPT"]=int(Oil_Prod.iloc[17,7])
st.session_state["OPY"]=int(Oil_Prod.iloc[17,6])








