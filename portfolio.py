import gspread
from oauth2client.service_account import ServiceAccountCredentials
from scipy import stats
import seaborn as sns
import pandas as pd
import yfinance as yf
import matplotlib.pyplot as plt
import numpy as np
import streamlit as st
import os.path
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import statsmodels.api as sm
from statsmodels import regression

# If modifying these scopes, delete the file token.json.
import xlwings as xw
wb = xw.Book('PAG.xlsx')
sht = xw.sheets('Data')
sht2 = xw.sheets('Stats')

COMM_Holdings = sht.range('B31:B40').value
COMM_Weights = sht2.range('G7:G11').value

CD_Holdings = sht.range('B41:B50').value
CD_Weights = sht2.range('G13:G17').value

CS_Holdings = sht.range('B51:B60').value
CS_Weights = sht2.range('G19:G24').value

E_Holdings =sht.range('B61:B70').value
E_Weights = sht2.range('G26:G30').value

FIN_Holdings = sht.range('B71:B80').value
FIN_Weights = sht2.range('G32:G38').value

H_Holdings = sht.range('B81:B90').value
H_Weights = sht2.range('G40:G48').value

IND_Holdings = sht.range('B91:B100').value
IND_Weights = sht2.range('G50:G56').value

IT_Holdings = sht.range('B101:B110').value
IT_Weights = sht2.range('G58:G66').value

MAT_Holdings =sht.range('B111:B120').value
MAT_Weights = sht2.range('G68:G71').value

RE_Holdings = sht.range('B121:B130').value
RE_Weights = sht2.range('G73:G76').value

U_Holdings = sht.range('B131:B140').value
U_Weights = sht2.range('G78:G81').value


option = st.selectbox(
    'Select your sector',
    ('COMM', 'CD', 'CS', 'E', 'FIN', 'H', 'IND', 'IT', 'MAT', 'RE', 'U'))

if option == 'COMM':
    holdings = COMM_Holdings
    WTS = COMM_Weights
if option == 'CD':
    holdings = CD_Holdings
    WTS = CD_Weights
if option == 'CS':
    holdings = CS_Holdings
    WTS = CS_Weights
if option == 'E':
    holdings = E_Holdings
    WTS = E_Weights
if option == 'FIN':
    holdings = FIN_Holdings
    WTS = FIN_Weights

if option == 'H':
    holdings = H_Holdings
    WTS = H_Weights

if option == 'IND':
    holdings = IND_Holdings
    WTS = IND_Weights
if option == 'IT':
    holdings = IT_Holdings
    WTS = IT_Weights

if option == 'MAT':
    holdings = MAT_Holdings
    WTS = MAT_Weights
if option == 'RE':
    holdings = RE_Holdings
    WTS = RE_Weights
if option == 'U':
    holdings = U_Holdings
    WTS = U_Weights


sum = sum(WTS)
for x in range(len(WTS)):
    WTS[x] = (WTS[x]/sum)
holdings = list(filter(lambda item: item is not None, holdings))
assets = holdings


assets = holdings
weights = WTS
st.title("Nittany Lion Fund")
start = st.date_input("Pick a start date for portfolio history", value = pd.to_datetime('2023-01-01'))
data = yf.download(assets, start=start)['Adj Close']
ret_df = data.pct_change()[1:]
port_ret = (ret_df * WTS).sum(axis = 1)
cumul_ret = (port_ret + 1).cumprod()-1
benchmark = yf.download('^SPX', start = start)['Adj Close']
bench_ret = benchmark.pct_change()[1:]
bench_dev = (bench_ret + 1).cumprod() - 1 
W = (np.ones(len(ret_df.cov()))/len(ret_df.cov()))
pf_std = (W.dot(ret_df.cov()).dot(W)) ** 1/2
st.subheader('Portfolio vs. Index')
x = bench_dev
y = cumul_ret
x.name = 'Benchmark Performance'
y.name = ('Portfolio Performance')
tog = pd.concat([bench_dev, cumul_ret], axis=1)

st.line_chart(data = tog)

st.subheader('Portfolio Beta:')
(beta, alpha) = stats.linregress(bench_dev.values,
                cumul_ret.values)[0:2]
beta = round(beta, 2)
st.metric(label="Beta", value=beta)


st.subheader('Weekly Performance')
if len(assets) == 1:
    col1 = st.columns(1, gap = 'small')
if len(assets) == 2:
    col1, col2= st.columns(2, gap = 'small')
if len(assets) == 3:
    col1, col2, col3= st.columns(3, gap = 'small')
if len(assets) == 4:
    col1, col2, col3, col4= st.columns(4, gap = 'small')
if len(assets) == 5:
    col1, col2, col3, col4, col5 = st.columns(5, gap = 'small')
if len(assets) == 6:
    col1, col2, col3, col4, col5, col6= st.columns(6, gap = 'small')
if len(assets) == 7:
    col1, col2, col3, col4, col5, col6, col7 = st.columns(7, gap = 'small')
if len(assets) == 8:
    col1, col2, col3, col4, col5, col6, col7, col8 = st.columns(8, gap = 'small')
if len(assets) == 9:
    col1, col2, col3, col4, col5, col6, col7, col8, col9 = st.columns(9, gap = 'small')
if len(assets) > 0:
    col1.metric(assets[0], round(data.iloc[0][assets[0]], 2), round(data.iloc[1][assets[0]] - data.iloc[0][assets[0]],2)
)
if len(assets) > 1:
    col2.metric(assets[1],round(data.iloc[0][assets[1]],2), round(data.iloc[1][assets[1]] - data.iloc[0][assets[1]],2)
)
if len(assets) > 2:
    col3.metric(assets[2], round(data.iloc[0][assets[2]],2), round(data.iloc[1][assets[2]] - data.iloc[0][assets[2]],2)
)
if len(assets) > 3:
    col4.metric(assets[3], round(data.iloc[0][assets[3]],2), round(data.iloc[1][assets[3]] - data.iloc[0][assets[3]],2)
)
if len(assets) > 4:
    col5.metric(assets[4], round(data.iloc[0][assets[4]],2), round(data.iloc[1][assets[4]] - data.iloc[0][assets[4]],2)
)
if len(assets) > 5:
    col6.metric(assets[5], round(data.iloc[0][assets[5]],2), round(data.iloc[1][assets[5]] - data.iloc[0][assets[5]],2)
)
if len(assets) > 6:
    col7.metric(assets[6], round(data.iloc[0][assets[6]],2), round(data.iloc[1][assets[6]] - data.iloc[0][assets[6]],2)
)
if len(assets) > 7:
    col8.metric(assets[7], round(data.iloc[0][assets[7]],2), round(data.iloc[1][assets[7]] - data.iloc[0][assets[7]],2)
)
if len(assets) > 8:
    col8.metric(assets[8], round(data.iloc[0][assets[8]],2), round(data.iloc[1][assets[8]] - data.iloc[0][assets[8]],2)
)


st.subheader('Portfolio Composition')
fig, ax = plt.subplots(facecolor = '#FFFFFF')
ax.pie(WTS, labels = assets, autopct='%1.1f%%', textprops={'color':'black'})
st.pyplot(fig)
#python -m streamlit run "c:/Users/ompat/Documents/Portfolio Visualizer/portfolio.py
