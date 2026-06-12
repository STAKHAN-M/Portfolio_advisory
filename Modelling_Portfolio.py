import xlwings as xw
import yfinance as yf
import numpy as np 
import pandas as pd 
import plotly.express as px
from pypfopt import expected_returns, risk_models, EfficientFrontier
import warnings
import Portfolio_functions as pf
warnings.simplefilter(action='ignore', category= FutureWarning)

'''
MODELLING PORTFOLIO IS A AUTOMATE PROGRAM WHICH ALLOW YOU TO OPTIMIZE YOUR PORTFOLIO MANAGEMENT & ANALYSIS ! 

BY EXTRACTING STOCKS HOLD AND THEIR WEIGHTS IN YOUR PORTFOLIO, WE WILL EXTRACT PRICE TO ANALYZE YOUR PERFORMANCE AGAINST A BENCHMARK  
AND TRY TO FIGURE OUT AN OPTIMAL ALLOCATION REGARDING YOUR EXPECTATIONS (RETURNS, RISKS)

THIS PROGRAM WORKS WITH A COUPLE OF FILES: 
- Portfolio_functions.py
- Portfolio_Template.xlsm

BE AWARE TO THE PATH OF "Portftolio_Template.xlsm" & 
BE SURE TO RUN THIS PROGRAM WITH "Portfolio_functions.py" IN THE SAME FOLDER

'''

# ----------------------------------------------------------------------------
### ASSUMPTIONS

rf = 0.033
exp_return = 0.08

###  EXCEL DATA EXTRACTION 
## Implementation of Excel Files containing Portfolio
# You must use Portfolio model 

app = xw.App(visible=False)
wb = xw.Book(r"C:\Users\rmore\Desktop\Portfolio_Le_M.xlsm")

# 1. ACTUAL PORTFOLIO
# §§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§

sheet_PF = wb.sheets["Portfolio"]
value_pf = sheet_PF["C10:C400"].value
w_pf = sheet_PF["S10:S400"].value

# ---------------------------------------------------------------------------

## Looping to build portfolio tickers' list and weight
# Create a list of tickers 
liste_pf = []

for x in range(len(value_pf)):
    add = value_pf[x]
    if add is None:
        continue
    else:
        liste_pf.append(add)

# Create a list with weight of actual portfolio
weight_pf = []

for x in range(len(w_pf)):
    add = w_pf[x]
    if add is None:
        continue
    else:
        weight_pf.append(add)

w_actual = np.array(weight_pf)

# !!!! REMEMBER TO SWITCH IN THE CODE => TICKER BY liste_pf

benchmark = ["^FCHI"]

# ------------------------------------------------------------------------

### EXTRACTING DATA AND COMPUTE RETURNS, VOL, COVARIANCE, CORR, VAR

data_day = yf.download(liste_pf,start="2023-01-01", interval="1d")["Close"]

## Compute the stocks returns 
data_stock = yf.download(liste_pf, start = "2018-12-01", interval= "1mo")["Close"]
basis_return_stock = ((data_stock/data_stock.shift(1)) - 1)
basis_return_stock = basis_return_stock.fillna(0)
basis_return_stock = basis_return_stock[1:]

## Compute the benchmark returns
data_bench = yf.download(benchmark, start = "2018-12-01", interval= "1mo")["Close"]
basis_return_bench = ((data_bench/data_bench.shift(1)) - 1)
basis_return_bench = basis_return_bench.fillna(0)
basis_return_bench = basis_return_bench[1:]


## DISPLAY TABLE (DATAFRAME) WITH RETURNS 
table_stock_return = pf.return_ratio(basis_return_stock, ind = liste_pf)
table_bench_return = pf.return_ratio(basis_return_bench, ind = benchmark)

table_return = pd.concat([table_stock_return, table_bench_return], axis = 0)
print(table_return)


## Compute matrix variance covariance 
cov_stock = basis_return_stock.cov()

## Compute the correlation of the portfolio
matrix_returns = pd.concat([basis_return_stock, basis_return_bench], axis=1)
matrix_returns.dropna(inplace=True)
corr_matrix = matrix_returns.corr()
corr_asset = corr_matrix.iloc[0,1]

print("Portfolio's Corr w/ benchmark : ", corr_asset.round(4) )

# Function to compute the portfolio performance by stocks
return_index = (100*(1+basis_return_stock).cumprod()).round(2)

## Compute the Sharpe Ratio for stocks and benchmark
sh_ratio_stock = pf.sharpe_ratio(basis_return_stock, rf, 12)
sh_ratio_bench = pf.sharpe_ratio(basis_return_bench, rf, 12)

# ------------------------------------------------------------------------

## DISPLAY TABLE (DATAFRAME) WITH DRAWDOWN / VARIATION FROM BEGINNING / HIGEST PEAK

table_drawdown_bench = pf.drawdown(basis_return_bench[benchmark[0]], ind=basis_return_bench.index)
table_drawdown_bench['Ticker'] = 'Benchmark'

portfolio_returns = basis_return_stock.dot(w_actual)
# Calcul du drawdown du portefeuille
table_drawdown_portfolio = pf.drawdown(portfolio_returns, ind=basis_return_stock.index)

# Ajout du nom pour identification
table_drawdown_portfolio['Ticker'] = 'Portfolio' 

## VaR & CVaR

df_Var_stock = pd.DataFrame({
    "VaR historic" : pf.var_historic(basis_return_stock),
    "VaR Gaussian" : pf.var_gaussian(basis_return_stock),
    "CVaR" : pf.cvar_historic(basis_return_stock)}, index = liste_pf)

df_VaR_bench = pd.DataFrame({
    "VaR historic" : pf.var_historic(basis_return_bench),
    "VaR Gaussian" : pf.var_gaussian(basis_return_bench),
    "CVaR" : pf.cvar_historic(basis_return_bench)}, index = benchmark)

table_VaR = pd.concat([df_Var_stock, df_VaR_bench], axis = 0)

# ------------------------------------------------------------------------

### PORTFOLIO OPTIMIZATION
## COMPUTE ANNUALIZE RETURNS & VOL 

set_ret_stock = pf.annualize_rets(basis_return_stock, 12) #Put 12 bc it's monthly return
set_vol_stock = pf.annualize_vol(basis_return_stock, 12)

set_ret_bench = pf.annualize_rets(basis_return_bench, 12)
set_vol_bench = pf.annualize_vol(basis_return_bench, 12)
sh_bench = (set_ret_bench - rf)/set_vol_bench

# ------------------------------------------------------------------------

## BLOC TO REVIEW -> TRY TO MODELLING DIFFERENT PORTFOLIO

# §§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§
# Initialization w/ mu = expected return and s = cov matrix 
mu = expected_returns.mean_historical_return(basis_return_stock, returns_data=True, frequency=12)
s = risk_models.sample_cov(basis_return_stock, returns_data=True)

w_opti = pf.minimize_vol(exp_return, set_ret_stock, cov_stock) #input opti return and find me a pf with the optimal allocation

# §§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§

actual = pf.actual_pf(w_actual, set_ret_stock, cov_stock, rf)

equal = pf.equal_w_pf(liste_pf, set_ret_stock, cov_stock, rf)

er_min_vol = pf.portfolio_ret(w_opti, set_ret_stock)
ev_min_vol = pf.portfolio_vol(w_opti, cov_stock)
sh_min_vol = (er_min_vol - rf)/ev_min_vol

msr = pf.msr_w_pf(mu, s, set_ret_stock, cov_stock, rf)

gmv = pf.gmv_w_pf(mu, s, set_ret_stock, cov_stock, rf)

utility = pf.utility_w_pf(mu, s, set_ret_stock, cov_stock, rf)

sortino = pf.sortino_w_pf(mu, s, set_ret_stock, cov_stock, rf)

'''
# !!!! ACTUAL PORTFOLIO
w_actual = np.array(weight_pf)
return_actual_pf = pf.portfolio_ret(w_actual, set_ret_stock)
vol_actual_pf = pf.portfolio_vol(w_actual, cov_stock)
sh_actual = (return_actual_pf - rf)/vol_actual_pf

# §§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§

# EQUAL PORFTOLIO
w = np.repeat(1/(len(liste_pf)), len(liste_pf))
return_pf = pf.portfolio_ret(w, set_ret_stock)
vol_pf = pf.portfolio_vol(w, cov_stock)
sh_equal = (return_pf - rf)/vol_pf

# §§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§


## A revoir car output weird 

print("This is the optimal weight for each stock : ", w_opti.round(3)) 

# §§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§

# MINIMUM VOLATILITY

er_min_vol = pf.portfolio_ret(w_opti, set_ret_stock)
ev_min_vol = pf.portfolio_vol(w_opti, cov_stock)
sh_min_vol = (er_min_vol - rf)/ev_min_vol

# §§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§

## MSR 
ef_msr = EfficientFrontier(mu,s)
ef_msr.max_sharpe()
w_msr = ef_msr.clean_weights()
w_msr = list(w_msr.values())
# Weight
w_msr = np.array(w_msr)
# Return & Vol
er_msr = pf.portfolio_ret(w_msr, set_ret_stock)
ev_msr = pf.portfolio_vol(w_msr, cov_stock)
sh_msr = (er_msr - rf)/ev_msr

# §§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§

## GLOBAL MINIMUM VARIANCE
ef_mv = EfficientFrontier(mu,s)
ef_mv.min_volatility()
w_gmv = ef_mv.clean_weights()
w_gmv = list(w_gmv.values())
# Weight
w_gmv = np.array(w_gmv)
# Return & Vol 
er_gmv = pf.portfolio_ret(w_gmv, set_ret_stock)
ev_gmv = pf.portfolio_vol(w_gmv, cov_stock)
sh_gmv = (er_gmv - rf)/ev_gmv

# §§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§

## MAX UTILITY 
ef_utility = EfficientFrontier(mu,s)
ef_utility.max_quadratic_utility()
w_utility = ef_utility.clean_weights()
w_utility = list(w_utility.values())
# Weight
w_utility = np.array(w_utility)
# Return & Vol 
er_utility = pf.portfolio_ret(w_utility, set_ret_stock)
ev_utility = pf.portfolio_vol(w_utility, cov_stock)
sh_utility = (er_utility - rf)/ev_utility

# §§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§

## MAX SORTINO
ef_sortino = EfficientFrontier(mu,s)
ef_sortino.max_quadratic_utility(np.sqrt(12),rf)
w_sortino = ef_sortino.clean_weights()
w_sortino = list(w_sortino.values())
# Weight
w_sortino = np.array(w_sortino)
# Return & Vol 
er_sortino = pf.portfolio_ret(w_sortino, set_ret_stock)
ev_sortino = pf.portfolio_vol(w_sortino, cov_stock)
sh_sortino = (er_sortino - rf)/ev_sortino
'''

# ------------------------------------------------------------------------

### CHART SECTION

fig_returns = px.line(return_index, x=return_index.index, y=liste_pf)
fig_ret_vol = px.histogram(table_stock_return, x=liste_pf, y=table_bench_return.columns,
             barmode="group", height = 450)


# ------------------------------------------------------------------------

###PORTFOLIO ANALYSIS RATIO (MARKET, OPTIMIZATION PF, RETURNS & VOL)
## MARKET RATIO (PER, P/B, DIV YIELD, BETA) 
#info = pf.information_portfolio(liste_pf)
#print(pf.information_portfolio(liste_pf))

table_w_stock = pd.DataFrame({
        "Annualized Return" : set_ret_stock.round(4),
        "Volatility" : set_vol_stock.round(4),
        "Actual Portfolio" : w_actual.round(4),
        "Weight" : equal[0].round(4),
        "Sharpe Ratio" : sh_ratio_stock.round(4),
        f'Min Variance Weight for {exp_return} Exp Return' : w_opti.round(4),
        "Max Sharpe Ratio Weight" : msr[0].round(4),
        "Global Min Variance Weight" : gmv[0].round(4),
        "Max Utility" : utility[0].round(4),
        "Max Sortino" : sortino[0].round(4)
    })

table_w_stock = table_w_stock.transpose()

print(table_w_stock)

## RATIO TABLE INCLUDING ANNUALIZED RET & VOL, SHARPE

table_ratio_pf = pd.DataFrame({
        "Name" : ["Benchmark","Actual Portfolio","Equal Portfolio","Minimum Variance", "Maximum Sharpe Ratio", "Global Minimum Variance", "Max Utility", "Max Sortino"],
        "Annualized Return" : [float(set_ret_bench.round(4)),float(actual[1].round(4)),float(equal[1].round(4)),float(er_min_vol.round(4)), float(msr[1].round(4)), float(gmv[1].round(4)), float(utility[1].round(4)), float(sortino[1].round(4))],
        "Volatility" : [float(set_vol_bench.round(4)),float(actual[2].round(4)),float(equal[2].round(4)),float(ev_min_vol.round(4)), float(msr[2].round(4)), float(gmv[2].round(4)), float(utility[2].round(4)), float(sortino[2].round(4))],
        "Sharpe Ratio" : [float(sh_bench.round(4)),float(actual[3].round(4)),float(equal[3].round(4)), float(sh_min_vol.round(4)), float(msr[3].round(4)), float(gmv[3].round(4)), float(utility[3].round(4)), float(sortino[3].round(4))]

})


table_ratio_pf = table_ratio_pf.transpose()
print(table_ratio_pf)

## Performance table
stock_perf = (100*(1+basis_return_stock).cumprod()).round(2)
bench_perf = (100*(1+basis_return_bench).cumprod()).round(2)


total_perf_actual_pf = stock_perf.multiply(w_actual, axis=1)
total_perf_actual_pf = total_perf_actual_pf.sum(axis=1)

total_perf_equal_pf = stock_perf.multiply(equal[0], axis=1)
total_perf_equal_pf = total_perf_equal_pf.sum(axis=1)

total_perf_min_vol_pf = stock_perf.multiply(w_opti, axis=1)
total_perf_min_vol_pf = total_perf_min_vol_pf.sum(axis=1)

total_perf_msr_pf = stock_perf.multiply(msr[0], axis=1)
total_perf_msr_pf = total_perf_msr_pf.sum(axis=1)

total_perf_gmv_pf = stock_perf.multiply(gmv[0], axis=1)
total_perf_gmv_pf = total_perf_gmv_pf.sum(axis=1)


table_perf = pd.DataFrame({
    "Benchmark" : bench_perf.squeeze(),
    "Actual PF " : total_perf_actual_pf,
    "Excess Return Actual" : (total_perf_actual_pf - bench_perf.squeeze())/100,
    "Equal PF" : total_perf_equal_pf,
    "Excess Return Equal": (total_perf_equal_pf - bench_perf.squeeze())/100,
    "Minimum Variance PF": total_perf_min_vol_pf,
    "Excess Return MV": (total_perf_min_vol_pf - bench_perf.squeeze())/100,
    "Max Sharpe Ratio PF": total_perf_msr_pf,
    "Excess Return MSR": (total_perf_msr_pf - bench_perf.squeeze())/100,
    "Global Minimum Variance PF": total_perf_gmv_pf,
    "Excess Return GMV": (total_perf_gmv_pf - bench_perf.squeeze())/100
})




# 2. PORTFOLIO SIMULATION ONE 
# §§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§

sheet_PF_S_1 = wb.sheets["Portfolio_Simulation_One"]
value_pf_S_1 = sheet_PF_S_1["C10:C400"].value
w_pf_S_1 = sheet_PF_S_1["S10:S400"].value

# ---------------------------------------------------------------------------

## Looping to build portfolio tickers' list and weight
liste_pf_S_1 = []

for x in range(len(value_pf_S_1)):
    add = value_pf_S_1[x]
    if add is None:
        continue
    else:
        liste_pf_S_1.append(add)

# Create a list with weight of actual portfolio
weight_pf_S_1 = []

for x in range(len(w_pf_S_1)):
    add = w_pf_S_1[x]
    if add is None:
        continue
    else:
        weight_pf_S_1.append(add)

w_pf_S_1 = np.array(weight_pf_S_1)

# ------------------------------------------------------------------------

### EXTRACTING DATA AND COMPUTE RETURNS, VOL, COVARIANCE, CORR, VAR

data_day_s1 = yf.download(liste_pf_S_1,start="2023-01-01", interval="1d")["Close"]

## Compute the stocks returns 
data_stock_S_1 = yf.download(liste_pf_S_1, start = "2018-12-01", interval= "1mo")["Close"]
basis_return_stock_S_1 = ((data_stock_S_1/data_stock_S_1.shift(1)) - 1)
basis_return_stock_S_1 = basis_return_stock_S_1.fillna(0)
basis_return_stock_S_1 = basis_return_stock_S_1[1:]


## DISPLAY TABLE (DATAFRAME) WITH RETURNS 
table_stock_return_S_1 = pf.return_ratio(basis_return_stock_S_1, ind = liste_pf_S_1)

table_stock_return_S_1 = pd.concat([table_stock_return_S_1, table_bench_return], axis = 0)
print(table_stock_return_S_1)

## Compute matrix variance covariance 
cov_stock_S_1 = basis_return_stock_S_1.cov()

## Compute the correlation of the portfolio
matrix_returns_S_1 = pd.concat([basis_return_stock_S_1, basis_return_bench], axis=1)
matrix_returns_S_1.dropna(inplace=True)
corr_matrix_S_1 = matrix_returns_S_1.corr()
corr_asset_S_1 = corr_matrix_S_1.iloc[0,1]

print("Portfolio's Corr w/ benchmark : ", corr_asset_S_1.round(4) )

## Compute the Sharpe Ratio for stocks and benchmark
sh_ratio_stock_S_1 = pf.sharpe_ratio(basis_return_stock_S_1, rf, 12)

## VaR & CVaR

df_Var_stock_S_1 = pd.DataFrame({
    "VaR historic" : pf.var_historic(basis_return_stock_S_1),
    "VaR Gaussian" : pf.var_gaussian(basis_return_stock_S_1),
    "CVaR" : pf.cvar_historic(basis_return_stock_S_1)}, index = liste_pf)


table_VaR_S_1 = pd.concat([df_Var_stock_S_1, df_VaR_bench], axis = 0)

# ------------------------------------------------------------------------

### PORTFOLIO OPTIMIZATION
## COMPUTE ANNUALIZE RETURNS & VOL 

set_ret_stock_S_1 = pf.annualize_rets(basis_return_stock_S_1, 12) #Put 12 bc it's monthly return
set_vol_stock_S_1 = pf.annualize_vol(basis_return_stock_S_1, 12)

# ------------------------------------------------------------------------

# Initialization w/ mu = expected return and s = cov matrix 
mu_S_1 = expected_returns.mean_historical_return(basis_return_stock_S_1, returns_data=True, frequency=12)
s_S_1 = risk_models.sample_cov(basis_return_stock_S_1, returns_data=True)

w_opti_S_1 = pf.minimize_vol(exp_return, set_ret_stock_S_1, cov_stock_S_1) #input opti return and find me a pf with the optimal allocation

# ------------------------------------------------------------------------

actual_S_1 = pf.actual_pf(w_pf_S_1, set_ret_stock_S_1, cov_stock_S_1, rf)

equal_S_1 = pf.equal_w_pf(liste_pf_S_1, set_ret_stock_S_1, cov_stock_S_1, rf)

er_min_vol_S_1 = pf.portfolio_ret(w_opti_S_1, set_ret_stock_S_1)
ev_min_vol_S_1 = pf.portfolio_vol(w_opti_S_1, cov_stock_S_1)
sh_min_vol_S_1 = (er_min_vol_S_1 - rf)/ev_min_vol_S_1

msr_S_1 = pf.msr_w_pf(mu_S_1, s_S_1, set_ret_stock_S_1, cov_stock_S_1, rf)

gmv_S_1 = pf.gmv_w_pf(mu_S_1, s_S_1, set_ret_stock_S_1, cov_stock_S_1, rf)

utility_S_1 = pf.utility_w_pf(mu_S_1, s_S_1, set_ret_stock_S_1, cov_stock_S_1, rf)

sortino_S_1 = pf.sortino_w_pf(mu_S_1, s_S_1, set_ret_stock_S_1, cov_stock_S_1, rf)

# ------------------------------------------------------------------------

table_w_stock_S_1 = pd.DataFrame({
        "Annualized Return" : set_ret_stock_S_1.round(4),
        "Volatility" : set_vol_stock_S_1.round(4),
        "Actual Portfolio" : w_pf_S_1.round(4),
        "Weight" : equal_S_1[0].round(4),
        "Sharpe Ratio" : sh_ratio_stock_S_1.round(4),
        f'Min Variance Weight for {exp_return} Exp Return' : w_opti_S_1.round(4),
        "Max Sharpe Ratio Weight" : msr_S_1[0].round(4),
        "Global Min Variance Weight" : gmv_S_1[0].round(4),
        "Max Utility" : utility_S_1[0].round(4),
        "Max Sortino" : sortino_S_1[0].round(4)
    })

table_w_stock_S_1 = table_w_stock_S_1.transpose()

print(table_w_stock_S_1)

## RATIO TABLE INCLUDING ANNUALIZED RET & VOL, SHARPE

table_ratio_pf_S_1 = pd.DataFrame({
        "Name" : ["Benchmark","Actual Portfolio","Equal Portfolio","Minimum Variance", "Maximum Sharpe Ratio", "Global Minimum Variance", "Max Utility", "Max Sortino"],
        "Annualized Return" : [float(set_ret_bench.round(4)),float(actual_S_1[1].round(4)),float(equal_S_1[1].round(4)),float(er_min_vol_S_1.round(4)), float(msr_S_1[1].round(4)), float(gmv_S_1[1].round(4)), float(utility_S_1[1].round(4)), float(sortino_S_1[1].round(4))],
        "Volatility" : [float(set_vol_bench.round(4)),float(actual_S_1[2].round(4)),float(equal_S_1[2].round(4)),float(ev_min_vol_S_1.round(4)), float(msr_S_1[2].round(4)), float(gmv_S_1[2].round(4)), float(utility_S_1[2].round(4)), float(sortino_S_1[2].round(4))],
        "Sharpe Ratio" : [float(sh_bench.round(4)),float(actual_S_1[3].round(4)),float(equal_S_1[3].round(4)), float(sh_min_vol_S_1.round(4)), float(msr_S_1[3].round(4)), float(gmv_S_1[3].round(4)), float(utility_S_1[3].round(4)), float(sortino_S_1[3].round(4))]

})

table_ratio_pf_S_1 = table_ratio_pf_S_1.transpose()
print(table_ratio_pf_S_1)

# ------------------------------------------------------------------------

## Performance table
stock_perf_S_1 = (100*(1+basis_return_stock_S_1).cumprod()).round(2)

total_perf_pf_S_1 = stock_perf_S_1.multiply(w_pf_S_1, axis=1)
total_perf_pf_S_1 = total_perf_pf_S_1.sum(axis=1)



# 3. PORTFOLIO SIMULATION TWO 
# §§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§

sheet_PF_S_2 = wb.sheets["Portfolio_Simulation_Two"]
value_pf_S_2 = sheet_PF_S_2["C10:C400"].value
w_pf_S_2 = sheet_PF_S_2["S10:S400"].value

# ---------------------------------------------------------------------------

## Looping to build portfolio tickers' list and weight
liste_pf_S_2 = []

for x in range(len(value_pf_S_2)):
    add = value_pf_S_2[x]
    if add is None:
        continue
    else:
        liste_pf_S_2.append(add)

# Create a list with weight of actual portfolio
weight_pf_S_2 = []

for x in range(len(w_pf_S_2)):
    add = w_pf_S_2[x]
    if add is None:
        continue
    else:
        weight_pf_S_2.append(add)

w_pf_S_2 = np.array(weight_pf_S_2)

# ------------------------------------------------------------------------

### EXTRACTING DATA AND COMPUTE RETURNS, VOL, COVARIANCE, CORR, VAR

## Compute the stocks returns 
data_stock_S_2 = yf.download(liste_pf_S_2, start = "2018-12-01", interval= "1mo")["Close"]
basis_return_stock_S_2 = ((data_stock_S_2/data_stock_S_2.shift(1)) - 1)
basis_return_stock_S_2 = basis_return_stock_S_2.fillna(0)
basis_return_stock_S_2 = basis_return_stock_S_2[1:]


## DISPLAY TABLE (DATAFRAME) WITH RETURNS 
table_stock_return_S_2 = pf.return_ratio(basis_return_stock_S_2, ind = liste_pf_S_1)

table_stock_return_S_2 = pd.concat([table_stock_return_S_2, table_bench_return], axis = 0)
print(table_stock_return_S_2)

## Compute matrix variance covariance 
cov_stock_S_2 = basis_return_stock_S_2.cov()

## Compute the correlation of the portfolio
matrix_returns_S_2 = pd.concat([basis_return_stock_S_2, basis_return_bench], axis=1)
matrix_returns_S_2.dropna(inplace=True)
corr_matrix_S_2 = matrix_returns_S_2.corr()
corr_asset_S_2 = corr_matrix_S_2.iloc[0,1]

print("Portfolio's Corr w/ benchmark : ", corr_asset_S_2.round(4) )

## Compute the Sharpe Ratio for stocks and benchmark
sh_ratio_stock_S_2 = pf.sharpe_ratio(basis_return_stock_S_2, rf, 12)

## VaR & CVaR

df_Var_stock_S_2 = pd.DataFrame({
    "VaR historic" : pf.var_historic(basis_return_stock_S_2),
    "VaR Gaussian" : pf.var_gaussian(basis_return_stock_S_2),
    "CVaR" : pf.cvar_historic(basis_return_stock_S_2)}, index = liste_pf)

table_VaR_S_2 = pd.concat([df_Var_stock_S_2, df_VaR_bench], axis = 0)

# ------------------------------------------------------------------------

### PORTFOLIO OPTIMIZATION
## COMPUTE ANNUALIZE RETURNS & VOL 

set_ret_stock_S_2 = pf.annualize_rets(basis_return_stock_S_2, 12) #Put 12 bc it's monthly return
set_vol_stock_S_2 = pf.annualize_vol(basis_return_stock_S_2, 12)
# ------------------------------------------------------------------------

# Initialization w/ mu = expected return and s = cov matrix 
mu_S_2 = expected_returns.mean_historical_return(basis_return_stock_S_2, returns_data=True, frequency=12)
s_S_2 = risk_models.sample_cov(basis_return_stock_S_2, returns_data=True)

w_opti_S_2 = pf.minimize_vol(exp_return, set_ret_stock_S_2, cov_stock_S_2) #input opti return and find me a pf with the optimal allocation

# ------------------------------------------------------------------------

actual_S_2 = pf.actual_pf(w_pf_S_2, set_ret_stock_S_2, cov_stock_S_2, rf)

equal_S_2 = pf.equal_w_pf(liste_pf_S_2, set_ret_stock_S_2, cov_stock_S_2, rf)

er_min_vol_S_2 = pf.portfolio_ret(w_opti_S_2, set_ret_stock_S_2)
ev_min_vol_S_2 = pf.portfolio_vol(w_opti_S_2, cov_stock_S_2)
sh_min_vol_S_2 = (er_min_vol_S_2 - rf)/ev_min_vol_S_2

msr_S_2 = pf.msr_w_pf(mu_S_2, s_S_2, set_ret_stock_S_2, cov_stock_S_2, rf)

gmv_S_2 = pf.gmv_w_pf(mu_S_2, s_S_2, set_ret_stock_S_2, cov_stock_S_2, rf)

utility_S_2 = pf.utility_w_pf(mu_S_2, s_S_2, set_ret_stock_S_2, cov_stock_S_2, rf)

sortino_S_2 = pf.sortino_w_pf(mu_S_2, s_S_2, set_ret_stock_S_2, cov_stock_S_2, rf)

# ------------------------------------------------------------------------

table_w_stock_S_2 = pd.DataFrame({
        "Annualized Return" : set_ret_stock_S_2.round(4),
        "Volatility" : set_vol_stock_S_2.round(4),
        "Actual Portfolio" : w_pf_S_2.round(4),
        "Weight" : equal_S_2[0].round(4),
        "Sharpe Ratio" : sh_ratio_stock_S_2.round(4),
        f'Min Variance Weight for {exp_return} Exp Return' : w_opti_S_2.round(4),
        "Max Sharpe Ratio Weight" : msr_S_2[0].round(4),
        "Global Min Variance Weight" : gmv_S_2[0].round(4),
        "Max Utility" : utility_S_2[0].round(4),
        "Max Sortino" : sortino_S_2[0].round(4)
    })

table_w_stock_S_2 = table_w_stock_S_2.transpose()

print(table_w_stock_S_2)

## RATIO TABLE INCLUDING ANNUALIZED RET & VOL, SHARPE

table_ratio_pf_S_2 = pd.DataFrame({
        "Name" : ["Benchmark","Actual Portfolio","Equal Portfolio","Minimum Variance", "Maximum Sharpe Ratio", "Global Minimum Variance", "Max Utility", "Max Sortino"],
        "Annualized Return" : [float(set_ret_bench.round(4)),float(actual_S_2[1].round(4)),float(equal_S_2[1].round(4)),float(er_min_vol_S_2.round(4)), float(msr_S_2[1].round(4)), float(gmv_S_2[1].round(4)), float(utility_S_2[1].round(4)), float(sortino_S_2[1].round(4))],
        "Volatility" : [float(set_vol_bench.round(4)),float(actual_S_2[2].round(4)),float(equal_S_2[2].round(4)),float(ev_min_vol_S_2.round(4)), float(msr_S_2[2].round(4)), float(gmv_S_2[2].round(4)), float(utility_S_2[2].round(4)), float(sortino_S_2[2].round(4))],
        "Sharpe Ratio" : [float(sh_bench.round(4)),float(actual_S_2[3].round(4)),float(equal_S_2[3].round(4)), float(sh_min_vol_S_2.round(4)), float(msr_S_2[3].round(4)), float(gmv_S_2[3].round(4)), float(utility_S_2[3].round(4)), float(sortino_S_2[3].round(4))]

})

table_ratio_pf_S_2 = table_ratio_pf_S_2.transpose()
print(table_ratio_pf_S_2)

# ------------------------------------------------------------------------

## Performance table
stock_perf_S_2 = (100*(1+basis_return_stock_S_2).cumprod()).round(2)

total_perf_pf_S_2 = stock_perf_S_2.multiply(w_pf_S_2, axis=1)
total_perf_pf_S_2 = total_perf_pf_S_2.sum(axis=1)

# ------------------------------------------------------------------------

table_perf_multi_pf = pd.DataFrame({
    "Benchmark": bench_perf.squeeze(),
    "Actual PF": total_perf_actual_pf,
    "Portfolio Simulated 1": total_perf_pf_S_1,
    "Portfolio Simulated 2": total_perf_pf_S_2
})


# ------------------------------------------------------------------------

### EXTRACT PYTHON DATAS TO EXCEL

# 1. Actual Portfolio
# §§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§

ws_day_price = wb.sheets["DAY_PRICE"]
ws_day_price.clear_contents()
ws_day_price.range("A1").options(index = True).value = data_day

ws_day_price_s1 = wb.sheets["DAY_PRICE_S1"]
ws_day_price_s1.clear_contents()
ws_day_price_s1.range("A1").options(index = True).value = data_day_s1

ws_price_pf = wb.sheets["PRICE_PF"]
ws_price_pf.clear_contents()
ws_price_pf.range("A1").options(index = True).value = data_stock

ws_price_bench = wb.sheets["PRICE_BENCH"]
ws_price_bench.clear_contents()
ws_price_bench.range("A1").options(index = True).value = data_bench
ws_price_bench.range("D1").options(index = True).value = basis_return_bench

ws_return_pf = wb.sheets["RETURNS_PF"]
ws_return_pf.clear_contents()
ws_return_pf.range("A1").options(index = True).value = basis_return_stock

ws_covariance_matrix = wb.sheets["COV"]
ws_covariance_matrix.clear_contents()
ws_covariance_matrix.range("A1").options(index = True).value = cov_stock

ws_info = wb.sheets["INFO"]
ws_info.clear_contents()
#ws_info.range("A1").options(index = True).value = info
ws_info.range("O1").options(index = True).value = table_return


#ws_chart = wb.sheets["CHART"]
#ws_chart.clear_contents()
#ws_chart.pictures.add(fig_returns, name = "Performance of stocks", update = True)
#ws_chart.pictures.add(fig_ret_vol, name = "Returns and Volatility", update = True)

ws_var = wb.sheets["VAR"]
ws_var.clear_contents()
ws_var.range("A1").options(index = True).value = table_VaR

ws_drawdown = wb.sheets["DRAWDOWN"]
ws_drawdown.clear_contents()
ws_drawdown.range("A1").options(index = True).value = table_drawdown_bench
ws_drawdown.range("G1").options(index = True).value = table_drawdown_portfolio


ws_ratio = wb.sheets["RATIO"]
ws_ratio.clear_contents()
ws_ratio.range("A2").options(index = True).value = table_w_stock
ws_ratio.range("A17").options(index=True).value = table_ratio_pf
ws_ratio.range("A30").options(index = True).value = corr_asset

ws_perf = wb.sheets["PERFORMANCE"]
ws_perf.clear_contents()
ws_perf.range("B1").options(index = True).value = table_perf
ws_perf.range("O1").options(index = True).value = table_perf_multi_pf


# 2. Portfolio Simulation One
# §§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§

ws_var_S_1 = wb.sheets["VAR_S_1"]
ws_var_S_1.clear_contents()
ws_var_S_1.range("A1").options(index = True).value = table_VaR_S_1


ws_ratio_S_1 = wb.sheets["RATIO_S_1"]
ws_ratio_S_1.clear_contents()
ws_ratio_S_1.range("A2").options(index = True).value = table_w_stock_S_1
ws_ratio.range("A17").value = table_ratio_pf_S_1
ws_ratio_S_1.range("A30").options(index = True).value = corr_asset_S_1


ws_info_S_1 = wb.sheets["INFO_S_1"]
ws_info_S_1.clear_contents()
ws_info_S_1.range("A1").options(index = True).value = table_stock_return_S_1


# 3. Portfolio Simulation Two
# §§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§

ws_var_S_2 = wb.sheets["VAR_S_2"]
ws_var_S_2.clear_contents()
ws_var_S_2.range("A1").options(index = True).value = table_VaR_S_2


ws_ratio_S_2 = wb.sheets["RATIO_S_2"]
ws_ratio_S_2.clear_contents()
ws_ratio_S_2.range("A2").options(index = True).value = table_w_stock_S_2
ws_ratio_S_2.range("A17").options(index = True).value = table_ratio_pf_S_2
ws_ratio_S_2.range("A30").options(index = True).value = corr_asset_S_2


ws_info_S_2 = wb.sheets["INFO_S_2"]
ws_info_S_2.clear_contents()
ws_info_S_2.range("A1").options(index = True).value = table_stock_return_S_2



wb.save()
wb.close()
app.quit()

# ------------------------------------------------------------------------

print("***************************End of program***************************")

# fin du programme


