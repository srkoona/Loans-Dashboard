
import pandas as pd
import plotly.express as px
import streamlit as st
import requests
from io import BytesIO

st.set_page_config(page_title="Debt Comps", 
                   page_icon=":chart_increasing:", 
                   layout="wide")

url = 'https://franklintempleton-my.sharepoint.com/:x:/r/personal/s_rkoonadi_benefitstreetpartners_com/Documents/Desktop/Python/Dashboard/Debt%20Comps%20Dash%20HC.xlsx?d=w28708552ee434b359d6690957ae5f186&csf=1&web=1&e=Otl65q'

# Fetch the raw file content using requests
response = requests.get(url)
response.raise_for_status()  # Ensure the request was successful

# Read the content of the file into a pandas DataFrame
excel_file = BytesIO(response.content)  # Convert the content into a file-like object

df = pd.read_excel(
  io = excel_file,
  engine = "openpyxl",
  sheet_name= "Loans",
  skiprows=3,
  usecols="G:X",
  nrows=261
  )

df = getdata_excel()

#===========SIDEBAR SECTION===========

#YTM upper and lower limits
qupper_YTM = df["YTM"].quantile(0.95)
qlower_YTM = df["YTM"].quantile(0.1)

qupper_P = df["Ask"].quantile(0.99)
qlower_P = df["Ask"].quantile(0.135)

qupper_DM =  df["DM"].quantile(0.95)
qlower_DM = df["DM"].quantile(0.1)

df = df[(df["DM"] > qlower_DM) & (df["DM"] < qupper_DM)]

st.sidebar.header("Filter Options:")

sector = st.sidebar.multiselect(
    "Subsector:",
    options=df["Industry"].unique(),
    default=df["Industry"].unique()
)

subsector = st.sidebar.multiselect(
    "Subsector:",
    options=df["Segment"].unique(),
    default=df["Segment"].unique()
)

rating = st.sidebar.multiselect(
    "Rating:",
    options=df["Moodys"].unique(),
    default=df["Moodys"].unique()
)

yieldn, yieldx = st.sidebar.slider(
    "Yield:",
    min_value = df["YTM"].min(),
    max_value = df["YTM"].max(),
    value=(qupper_YTM, qlower_YTM),
    step = 0.01
)

pricen, pricex = st.sidebar.slider(
    "Ask Price:",
    min_value = df["Ask"].min(),
    max_value = df["Ask"].max(),
    value=(qupper_P, qlower_P),
    step = 0.1
)


DMn, DMx = st.sidebar.slider(
    "Discount Margin:",
    min_value = df["DM"].min(),
    max_value = df["DM"].max(),
    value=(qupper_DM, qlower_DM),
    step = 0.01
)

df_selection = df.query(
    "Industry == @sector & Segment == @subsector & Moodys == @rating & YTM >= @yieldn & YTM <= @yieldx & Ask >= @pricen & Ask <= @pricex & DM >= @DMn & DM <= @DMx"
)
st.dataframe(df_selection)

#===============MAINPAGE=====================

st.title("Debt Comparables")
st.markdown("##")

#===============TOP KPI's====================
Max_Y = round(df_selection["YTM"].max(), 1)
Maxid_Y = df_selection["YTM"].idxmax()
MaxName_Y = df.loc[Maxid_Y,"Issuer"]
Max_DM = round(df_selection["DM"].max(), 1)
Maxid_DM = df_selection["DM"].idxmax()
MaxName_DM = df.loc[Maxid_DM,"Issuer"]
AvgDM = round(df_selection["DM"].mean(), 1)

left_column, middle_column, right_column = st.columns(3)
with left_column:
    st.subheader("Max YTM")
    st.subheader(f"{Max_Y}")
    st.subheader(f"{Maxid_Y}")
with middle_column:
    st.subheader("Max DM")
    st.subheader(f"{Max_DM}")
    st.subheader(f"{MaxName_DM}")
with right_column:
    st.subheader("Avg DM")
    st.subheader(f"{AvgDM}")
st.markdown("""---""")


#============Filtering for outliers=======

qupper_DM =  df["DM"].quantile(0.9)
df_filtered = df_selection[(df_selection["DM"] > qlower_DM) & (df_selection["DM"] < qupper_DM)]


#=============DM by Segment===============


DM_Segment = df_filtered.groupby(by=["Segment"])["DM"].mean().to_frame()

chart_segment = px.bar(
    DM_Segment,
    x=DM_Segment.index,
    y="DM",
    title="<b>DM by Segment</b>",
    color_discrete_sequence=["#0083B8"] * len(DM_Segment),
    template="plotly_white",
)

#=============DM by Rating================
#for ordering by rating
gorder = {'Aaa': 1, 'Aa1': 2, 'Aa2': 3, 'Aa3': 4, 'A1': 5, 'A2': 6, 'A3': 7, 'Baa1': 8, 'Baa2': 9,'Baa3': 10, 'Ba1': 11, 'Ba2': 12, 'Ba3': 13, 'B1': 14, 'B2': 15, 'B3': 16, 'Caa1': 17, 'Caa2': 18, 'Caa3': 19, 'Ca': 20, 'NR': 21}

DM_Rating = df_filtered.groupby(by=["Moodys"])["DM"].mean().to_frame()
DM_Rating = DM_Rating.sort_values("Moodys", key=lambda x: x.map(gorder))

chart_rating = px.bar(
    DM_Rating,
    x=DM_Rating.index,
    y="DM",
    title="<b>DM by Rating</b>",
    color_discrete_sequence=["#0083B8"] * len(DM_Rating),
    template="plotly_white",
)

#display both the charts

left_column, right_column = st.columns(2)
left_column.plotly_chart(chart_segment, use_container_width=True)
right_column.plotly_chart(chart_rating, use_container_width=True)

#===========Price x DM Chart==========
mid_DM = df_filtered["DM"].mean()
mid_Ask = df_filtered["Ask"].mean()

DM_Price = df_filtered[["Issuer","Moodys","Ask","DM"]]

chart_op = px.scatter(
    DM_Price,
    x="DM",
    y="Ask",
    hover_name = "Issuer",
    color="Moodys",
    title="<b>Price x DM</b>"
)

#Adding reference lines to the charts

#vertical line
chart_op.add_shape(
    x0=mid_DM, x1=mid_DM, y0="80", y1=df['Ask'].max(),
    line=dict(color="Red", width=2, dash="dash")
)

#horizontal line
chart_op.add_shape(
    x0=qlower_DM, x1=qupper_DM, y0=mid_Ask, y1=mid_Ask,
    line=dict(color="Red", width=2, dash="dash")
)

#===============Leverage vs DM======================

#mid_L= df_filtered["LEVERAGE"].mean()

#DM_Lev = df_filtered[["Issuer","Leverage","Moodys","DM"]]
#chart_leverage = px.scatter(
#    DM_Lev,
#    x="DM",
#    y="Leverage",
#    hover_name = "Issuer",
#    color="Moodys",
#    title="<b>Leverage v DM</b>"
#)


#chart_lev.add_shape(
#      x0=mid_DM, x1=mid_DM, y0=df_filtered["Leverage"].min(), y1=df_filtered["Leverage"].max(),
#     line=dict(color="Red", width=2, dash="dash")
#)

#chart_lev.add_shape(
#    x0=qlower_DM, x1=qupper_DM, y0=mid_L, y1=mid_L,
#    line=dict(color="Red", width=2, dash="dash")
#)

#===================YTM v Price=======================
df_filtered_YTM = df_selection[(df_selection["YTM"] > qlower_YTM) & (df_selection["YTM"] < qupper_YTM)]
mid_YTM = df_filtered_YTM["YTM"].mean()
mid_Ask = df_filtered_YTM["Ask"].mean()

Yield_Price = df_filtered_YTM[["Issuer","YTM","Ask","Segment"]]

chart_yp = px.scatter(
    Yield_Price,
    x="YTM",
    y="Ask",
    hover_name="Issuer",
    color="Segment",
    title="<b>Price x YTM</b>"
)

#vertical line
chart_yp.add_shape(
    x0=mid_YTM, x1=mid_YTM, y0="80", y1=df['Ask'].max(),
    line=dict(color="Red", width=2, dash="dash")
)

qupper_YTM = df["YTM"].quantile(0.925)

#horizontal line
chart_yp.add_shape(
    x0=qlower_YTM, x1=qupper_YTM, y0=mid_Ask, y1=mid_Ask,
    line=dict(color="Red", width=2, dash="dash")
)

#=============DM ranked by Ratings====================

gorder = {'Aaa': 1, 'Aa1': 2, 'Aa2': 3, 'Aa3': 4, 'A1': 5, 'A2': 6, 'A3': 7, 'Baa1': 8, 'Baa2': 9,'Baa3': 10, 'Ba1': 11, 'Ba2': 12, 'Ba3': 13, 'B1': 14, 'B2': 15, 'B3': 16, 'Caa1': 17, 'Caa2': 18, 'Caa3': 19, 'Ca': 20, 'NR': 21}
Relrating = df_selection[["Issuer","Moodys","DM"]]
Relrating = Relrating.sort_values("Moodys", key=lambda x: x.map(gorder))

chart_Relrating = px.box(
    Relrating,
    x="Moodys",
    y="DM",
    points = "all",
    hover_name = "Issuer",
    title="<b>OAS Outliers by Rating</b>"
)


#=============New Project========================


st.plotly_chart(chart_op, use_container_width=True)
st.plotly_chart(chart_yp, use_container_width=True)
st.plotly_chart(chart_Relrating, use_container_width=True)
