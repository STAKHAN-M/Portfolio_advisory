import os
import xlwings as xw
import yfinance as yf
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import requests as rq
from scipy.stats import norm
from scipy.optimize import minimize
from pypfopt import expected_returns, risk_models, EfficientFrontier, objective_functions
import warnings
warnings.simplefilter(action='ignore', category= FutureWarning)
warnings.simplefilter(action='ignore', category= RuntimeWarning)


def information_portfolio(tickers):
    API_key = os.environ.get("FMP_API_KEY", "")
    dfs = []

    for ticker in tickers:
        IS = rq.get(f'https://financialmodelingprep.com/api/v3/income-statement/{ticker}?limit=2&apikey={API_key}').json()
        BS = rq.get(f'https://financialmodelingprep.com/api/v3/balance-sheet-statement/{ticker}?limit=2&apikey={API_key}').json()
        RATIO = rq.get(f"https://financialmodelingprep.com/api/v3/ratios/{ticker}?limit=2&apikey={API_key}").json()
        PROFILE = rq.get(f'https://financialmodelingprep.com/api/v3/profile/{ticker}?apikey={API_key}').json()
        SHARES = rq.get(f'https://financialmodelingprep.com/api/v4/shares_float?symbol={ticker}&apikey={API_key}').json()

        try:
            name = PROFILE[0]["companyName"]
            price = yf.download(ticker)["Adj Close"].iloc[-1]
            exchange = PROFILE[0]["exchangeShortName"]
            sector = PROFILE[0]["sector"]
            beta = PROFILE[0]["beta"]
            market_cap = price * SHARES[0]["outstandingShares"]
            ent_value = market_cap + BS[0]["netDebt"]
            per = market_cap / IS[0]["netIncome"]
            pb_ratio = market_cap / BS[0]["totalEquity"]
            ev_rev = ent_value / IS[0]["revenue"]
            ev_ebitda = ent_value / IS[0]["ebitda"]
            div_yield = RATIO[0]["dividendYield"]

            df = pd.DataFrame({
                "Ticker": [ticker],
                "Name": [name],
                "Exchange": [exchange],
                "Sector": [sector],
                "Price": [price.round(3)],
                "Beta": [beta],
                "Market Capitalization": [market_cap],
                "PER": [per.round(3)],
                "P/B Ratio": [pb_ratio.round(3)],
                "EV/REVENUE": [ev_rev.round(3)],
                "EV/EBITDA": [ev_ebitda.round(3)],
                "Dividend Yield": [div_yield]
            })
            dfs.append(df)
        except:
            print(f"Data not found for ticker: {ticker}")

    final_df = pd.concat(dfs, ignore_index=True)
    return final_df


def return_ratio(r, ind):
    '''
    Computes ratios below in a DataFrame from a table of returns
    BE CAREFUL FOR INTERVAL DATE 
    mean_d -> mean of daily return 
    mean_y -> annualized return
    vol -> std daily 
    vol_y -> std annual 
    r_r_ration -> compute the ratio btw return to risk 
    '''
    
    mean_d = (r.mean()).round(4)
    mean_y = (((1 + r.mean()) ** 12) -1).round(4)
    vol = (r.std()).round(4)
    vol_y = (r.std()*(12**0.5)).round(4)

    r_r_ratio = (mean_y / vol_y).round(4)

    table = pd.DataFrame({"Monthly Return (%)" : mean_d, "Yearly Return (%)" : mean_y, "Monthly Volatilty (%)" : vol, "Yearly Volatility (%)" : vol_y, " Return / Volatility" : r_r_ratio}, index = ind)
    return table


def drawdown(r, ind): 
    """
    Takes a times series of asset return
    Computes and returns a DataFrame that contains: 
    the wealth index
    the previous peak
    percent drawdowns  
    """

    #formula to compute performance of the portfolio
    wealth_index = (100*(1+r).cumprod()).round(2)
    #garde le plus haut historique
    previous_peak = wealth_index.cummax()
    drawdowns = ((wealth_index - previous_peak) / previous_peak).round(2)
    
    return pd.DataFrame({
        "Wealth": wealth_index,
        "Peaks": previous_peak,
        "Drawdown": drawdowns
    }, index=ind)

def var_historic(r, level = 5):
    """
    VaR Historic
    """
    if isinstance(r, pd.DataFrame):
        return r.aggregate(var_historic, level = level)
    elif isinstance(r, pd.Series):
        return -np.percentile(r, level)
    else: 
        raise TypeError("Expected r to be Series or DataFrame")

#Don't forget to import scipy.stats
def var_gaussian(r, level = 5): 
    """
    Returns the parametric Gaussian VaR of a Series or DataFrame
    """
    z = norm.ppf(level/100)
    
    return -(r.mean() + z * r.std(ddof = 0))


def cvar_historic(r, level = 5):
    """
    Computes the conditional VaR of Series or DataFrame
    """

    if isinstance(r, pd.Series):
        is_beyond = r <= -var_historic(r, level= level)
        return -r[is_beyond].mean()
    elif isinstance(r, pd.DataFrame):
        return r.aggregate(cvar_historic, level = level)
    else:
        raise TypeError("Expected r to be a Series or DataFrame")


def annualize_rets(r, periods_p_year):
    """
    Annualizes a set of returns (don't forget the period belonging to your df_return)
    Compute Compounded_growth 
    """
    comp_gwth = (1+r).prod()
    n_periods = r.shape[0]
    return comp_gwth ** (periods_p_year/n_periods)-1

def annualize_vol(r, periods_p_year):
    """
    Annualizes vol of a set of returns (be careful to period)
    """
    return r.std()*(periods_p_year**0.5)



def sharpe_ratio(r, rf, periods_p_year):
    """
    Computes sharpe ratio = (r.pf - rf) / vol.pf
    """

    rf_p_period = (1+rf)**(1/periods_p_year)-1
    excess_return = r - rf_p_period
    y_excess_return = annualize_rets(excess_return, periods_p_year)
    y_vol = annualize_vol(r,periods_p_year)
    return y_excess_return / y_vol


def portfolio_ret(w, annualized_returns):
    """
    Weights -> returns
    """
    return w.T @ annualized_returns

def portfolio_vol(w, covmat):
    """ 
    Weights -> Vol
    """
    return (w.T @ covmat @ w)**0.5

# --------------------------------------------------------------------------------------------

### Weight, Returns, Vol & Sharpe ratio of PF

def actual_pf(weight, returns, cov, rf): 
    w_actual = np.array(weight)
    return_actual_pf = portfolio_ret(w_actual, returns)
    vol_actual_pf = portfolio_vol(w_actual, cov)
    sh_actual = (return_actual_pf - rf)/vol_actual_pf
    return w_actual, return_actual_pf, vol_actual_pf, sh_actual


def equal_w_pf(liste, returns, cov, rf):
    w_equal = np.repeat(1/(len(liste)), len(liste))
    return_equal_pf = portfolio_ret(w_equal, returns)
    vol_equal_pf = portfolio_vol(w_equal, cov)
    sh_equal = (return_equal_pf - rf)/vol_equal_pf
    return w_equal, return_equal_pf, vol_equal_pf, sh_equal


def msr_w_pf(mu,s, returns, cov, rf):
    ef_msr = EfficientFrontier(mu,s)
    ef_msr.max_sharpe()
    w_msr = ef_msr.clean_weights()
    w_msr = list(w_msr.values())
    # Weight
    w_msr = np.array(w_msr)
    # Return & Vol
    er_msr = portfolio_ret(w_msr, returns)
    ev_msr = portfolio_vol(w_msr, cov)
    sh_msr = (er_msr - rf)/ev_msr
    return w_msr, er_msr, ev_msr, sh_msr

def gmv_w_pf(mu, s, returns, cov, rf):
    ## GLOBAL MINIMUM VARIANCE
    ef_mv = EfficientFrontier(mu,s)
    ef_mv.min_volatility()
    w_gmv = ef_mv.clean_weights()
    w_gmv = list(w_gmv.values())
    # Weight
    w_gmv = np.array(w_gmv)
    # Return & Vol 
    er_gmv = portfolio_ret(w_gmv, returns)
    ev_gmv = portfolio_vol(w_gmv, cov)
    sh_gmv = (er_gmv - rf)/ev_gmv
    return w_gmv, er_gmv, ev_gmv, sh_gmv

def utility_w_pf(mu, s, returns, cov, rf):
    ## MAX UTILITY 
    ef_utility = EfficientFrontier(mu,s)
    ef_utility.max_quadratic_utility()
    w_utility = ef_utility.clean_weights()
    w_utility = list(w_utility.values())
    # Weight
    w_utility = np.array(w_utility)
    # Return & Vol 
    er_utility = portfolio_ret(w_utility, returns)
    ev_utility = portfolio_vol(w_utility, cov)
    sh_utility = (er_utility - rf)/ev_utility
    return w_utility, er_utility, ev_utility, sh_utility


def sortino_w_pf(mu, s, returns, cov, rf): 
    ## MAX SORTINO
    ef_sortino = EfficientFrontier(mu,s)
    ef_sortino.max_quadratic_utility(np.sqrt(12),rf)
    w_sortino = ef_sortino.clean_weights()
    w_sortino = list(w_sortino.values())
    # Weight
    w_sortino = np.array(w_sortino)
    # Return & Vol 
    er_sortino = portfolio_ret(w_sortino, returns)
    ev_sortino = portfolio_vol(w_sortino, cov)
    sh_sortino = (er_sortino - rf)/ev_sortino
    return w_sortino, er_sortino, ev_sortino, sh_sortino


# --------------------------------------------------------------------------------------------

def minimize_vol(target_return, annualized_return, cov):
    """ 
    target return to W 
    """

    n = annualized_return.shape[0]
    init_guess = np.repeat(1/n, n)
    bounds = ((0.0,1.0),) * n
    return_is_target = {
        "type" : 'eq',
        "args" : (annualized_return,),
        "fun" : lambda w, annualized_return : target_return - portfolio_ret(w, annualized_return)
    }
    w_sum_to_1 = {
        "type": 'eq',
        'fun': lambda w : np.sum(w) - 1}
    results = minimize(portfolio_vol, init_guess, args =(cov,),method = "SLSQP", options = {"disp" : False}, constraints=(return_is_target, w_sum_to_1), bounds=bounds)
    return results.x

def optimal_w(n_points, annualized_return, cov):
    """ 
    output an optimal list of weights to run the optimizer -> minimize the vol
    """
    target_rs = np.linspace(annualized_return.min(),annualized_return.max(), n_points)
    w = [minimize_vol(target_return, annualized_return, cov) for target_return in target_rs]
    return w

def msr(rf, er, cov):
    """ 
    Risk free rate + E(R) + COV -> W => We're looking for the max Sharpe Ratio
    """

    n = er.shape[0]
    init_guess = np.repeat(1/n, n)
    bounds = ((0.0,1.0),) * n

    w_sum_to_1 = {
        "type": 'eq',
        'fun': lambda w : np.sum(w) - 1}

    def neg_sharpe_ratio(w, rf, er, cov):
        """   
        Returns the negative of the sharpe ratio, given weights
        """
        r = portfolio_ret(w, er)
        vol = portfolio_vol(w,cov)
        return -(r - rf)/vol

    results = minimize(neg_sharpe_ratio, init_guess, args =(rf, er, cov,),method = "SLSQP", options = {"disp" : False}, constraints=(w_sum_to_1), bounds=bounds)
    return results.x


def gmv(cov):
    """
    We compute GMV of the portfolio -> returns weight of GMV PF 
    """

    n = cov.shape[0]
    return msr(0, np.repeat(1,n), cov)


def plot_ef(n_points, annualized_return, cov, rf, show_cml = False, style = ".-", show_ew=False, show_gmv=False):
    """  
    Plot multi-asset efficient frontier
    """
    weights = optimal_w(n_points, annualized_return, cov)
    rets = [portfolio_ret(w,annualized_return) for w in weights]
    vols = [portfolio_vol(w,cov) for w in weights]
    ef = pd.DataFrame({
        "Returns" : rets,
        "Volatility" : vols
    })
    ax = ef.plot.line(x = "Volatility", y = "Returns", style = style)
    if show_ew:
        n = annualized_return.shape[0]
        w_ew = np.repeat(1/n,n)
        r_ew = portfolio_ret(w_ew, annualized_return)
        vol_ew = portfolio_vol(w_ew, cov)
        # display EW
        ax.plot([vol_ew], [r_ew], color="goldenrod", marker='o', markersize = 12)
    if show_gmv:
        w_gmv = gmv(cov)
        r_gmv = portfolio_ret(w_gmv,annualized_return)
        vol_gmv = portfolio_vol(w_gmv, cov)
        ax.plot([vol_gmv], [r_gmv], color = "midnightblue", marker= "o", markersize = 10)
    if show_cml:
        #ax.set_xlim(left = 0)
        #weights of sharpe ratio
        w_msr = msr(rf, annualized_return, cov)
        r_msr = portfolio_ret(w_msr, annualized_return)
        vol_msr = portfolio_vol(w_msr, cov)
        #Add Capital Market Line
        cml_x = [0, vol_msr]
        cml_y = [rf, r_msr]
        ax.plot(cml_x, cml_y, color = "red", linestyle = "dashed")

    return ax




    ### POSSIBLE GRAPH 

    """ 
    ## Chart pour le PF et le benchmark 

    #wealth_index.plot().line()
    #previous_peak.plot().line()
    #return_bench.plot.bar()

    ##STRUCURE POUR METTRE UN GRAPHIQUE DE PERF PF ET BENCHMARK 



    chart = data.plot(title =  f"{benchmark} Stock Price", ylabel = "Closing Price", figsize =[10,6])
    grid = plt.show()

    print(grid)
    """



