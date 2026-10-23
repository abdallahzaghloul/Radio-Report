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
OPT= st.session_state["OPT"]
OPY= st.session_state["OPY"]
WSW=st.session_state["WSW"]
WD = st.session_state['WD']

OPDiff = OPT-OPY



st.write(f"{RRE} %")
st.write(f"[{URL}   &   {url} ] Selected Report Date")

st.write(f" Today's SI Well's Count =[{SI_Wells_No}] ")

st.write(f" Today's Online Producing Well's Count =[{On_Line_Wells_No}] ")
st.write(f" Today [{URL}] Oil Production = [{OPT}] BO ")

st.write(f" Yesterday {[url]} Oil Production = [{OPY}] BO ")

if OPDiff <0:
            st.write(f" Differ's Oil Production = [{OPDiff}] BO ")
elif OPDiff >0:
            st.write(f" Gain's Oil Production = [{OPDiff}] BO ")
elif OPDiff ==0:
            st.write(f" No Differ in Today Vs. YesterdayProd. ")




WD= WD[(WD['STATUS']!="SI") & WD['WELL NAME'] ]
WD.index = range(1, len(WD) + 1)

OWW = WD['HRS ONLINE'] *  WD['UNNAMED: 36']*  WD['UNNAMED: 38']*.01*(1/24)

Total_WC_Production=OWW.sum()
OWW


st.write(f"Today's Total Water Produced F/ Oil Producing Wells {[Total_WC_Production]} BW ")


WSW= WSW[WSW['STATUS']!="SI" ]
WSW.index = range(1, len(WSW) + 1)

#WSW['BWPD']
WSW['FLUSH HRS'] = WSW['FLUSH HRS'].apply(lambda x: 0 if isnull(x)==True)
WSW['FLUSH HRS']
#Total_WS_Production = WSW['BWPD']*(1-(WSW['FLUSH HRS']/24))
#st.write(f"Today's Total Water Produced F/ WSW {[Total_WS_Production]} BW ")









