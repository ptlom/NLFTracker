

from scipy import stats
import pandas as pd
import yfinance as yf
import matplotlib.pyplot as plt
import numpy as np
import streamlit as st
from statsmodels import regression
from openpyxl.workbook import Workbook
from openpyxl import load_workbook
# If modifying these scopes, delete the file token.json.
import xlwings as xw
wb = load_workbook('C:\\Users\\ompat\\Documents\\Portfolio Visualizer\\PAG.xlsx')
sht = wb['Data']
sht2 = wb['Stats']

COMM_Holdings = sht['B31:B40']
COMM_Weights = sht2['G7:G11']

CD_Holdings = sht['B41:B50']
CD_Weights = sht2['G13:G17']

CS_Holdings = sht['B51:B60']
CS_Weights = sht2['G19:G24']

E_Holdings =sht['B61:B70']
E_Weights = sht2['G26:G30']

FIN_Holdings = sht['B71:B80']
FIN_Weights = sht2['G32:G38']

H_Holdings = sht['B81:B90']
H_Weights = sht2['G40:G48']

IND_Holdings = sht['B91:B100']
IND_Weights = sht2['G50:G56']

IT_Holdings = sht['B101:B110']
IT_Weights = sht2['G58:G66']

MAT_Holdings =sht['B111:B120']
MAT_Weights = sht2['G68:G71']

RE_Holdings = sht['B121:B130']
RE_Weights = sht2['G73:G76']

U_Holdings = sht['B131:B140']
U_Weights = sht2['G78:G81']


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
assets = []
for cell in holdings:
    for x in cell:
        assets.append(x.value)
weights = []
for cell in WTS:
    for x in cell:
        weights.append( x.value)

print(assets)
print(weights)



sum = sum(weights)


for x in range(len(weights)):
    weights[x] = (weights[x]/sum)
assets = list(filter(lambda item: item is not None, assets))

st.title("Nittany Lion Fund")
start = st.date_input("Pick a start date for portfolio history", value = pd.to_datetime('2023-01-01'))
data = yf.download(assets, start=start)['Adj Close']
ret_df = data.pct_change()[1:]
port_ret = (ret_df * weights).sum(axis = 1)
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
ax.pie(weights, labels = assets, autopct='%1.1f%%', textprops={'color':'black'})
st.pyplot(fig)
#python -m streamlit run "c:/Users/ompat/Documents/Portfolio Visualizer/portfolio.py
