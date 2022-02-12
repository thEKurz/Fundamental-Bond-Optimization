from pulp import *
def optimalportfolio(Portfolio=None, Arated=.20, duration=5.25, Subs=3, bonds=25, oneyearmature=0, LadderStart=3, LadderLength=7, MaxPerYear=5, maxladderlength=11, excel=1):
    import pandas as pd
    from dateutil.relativedelta import relativedelta
    from datetime import datetime
    import math
    import numpy as np
    now = datetime.now().strftime('%Y-%m-%d %H %M %S')
    prob = LpProblem("Corporate_Bond_Portfolio", LpMaximize)
    df = pd.read_excel("S:\AHB\Research\Fixed-Income\BondList.xlsx")
    ddf = pd.read_excel("S:\AHB\Research\Fixed-Income\BondList.xlsx")
    if Portfolio is not None:
        df = pd.concat([Portfolio, df]).drop_duplicates(subset='CUSIP',keep='last')
        pcusips = list(Portfolio['CUSIP'])
        ebonds=len(pcusips)
        df['In Portfolio'] = df['CUSIP'].isin(pcusips)
        df['In Portfolio'] = df['In Portfolio'].astype(int)
    df['Seniority Level'] = (df['Seniority Level'] == 'Subordinate')
    df['Seniority Level'] = df['Seniority Level'].astype(int)
    df['Rating'] = (df['Rating'].str.contains("A") == True)
    df['Rating'] = df['Rating'].astype(int)
    dfDummies = pd.get_dummies(df['Industry'], prefix='Sector')
    df = pd.concat([df, dfDummies], axis=1)
    df['Invested'] = 1
    Cusips = list(df['CUSIP'])
    yields = dict(zip(Cusips, df['Ask Yield To Worst']))
    Rating = dict(zip(Cusips, df['Rating']))
    Maturity = dict(zip(Cusips, df['Maturity']))
    Duration = dict(zip(Cusips, df['Duration']))
    Seniority = dict(zip(Cusips, df['Seniority Level']))
    Invested = dict(zip(Cusips, df['Invested']))
    Ticker = dict(zip(Cusips, df['Ticker']))
    CS=None
    if 'Sector_Communication Services' in df.columns:
        CS = dict(zip(Cusips, df['Sector_Communication Services']))
    CD=None
    if 'Sector_Consumer Discretionary' in df.columns:
        CD = dict(zip(Cusips, df['Sector_Consumer Discretionary']))
    CST=None
    if 'Sector_Consumer_Staples' in df.columns:
        CST = dict(zip(Cusips, df['Sector_Consumer Staples']))
    E = None
    if 'Sector_Energy' in df.columns:
        E = dict(zip(Cusips, df['Sector_Energy']))
    F=None
    if 'Sector_Financials' in df.columns:
        F = dict(zip(Cusips, df['Sector_Financials']))
    HC=None
    if 'Sector_Health Care' in df.columns:
        HC = dict(zip(Cusips, df['Sector_Health Care']))
    I=None
    if 'Sector_Industrials' in df.columns:
        I = dict(zip(Cusips, df['Sector_Industrials']))
    IT=None
    if 'Sector_Information Technology' in df.columns:
        IT = dict(zip(Cusips, df['Sector_Information Technology'])) 
    M=None
    if 'Sector_Materials' in df.columns:
        M = dict(zip(Cusips, df['Sector_Materials']))
    U=None
    if 'Sector_Utilities' in df.columns:
        U = dict(zip(Cusips, df['Sector_Utilities']))
    Symbols = list(df['Ticker'])
    LadderMax = df['Maturity'].max()
    LadderMin = df['Maturity'].min()
    Arated = math.floor(Arated*bonds)
    LadderYears = relativedelta(LadderMax, LadderMin).years
    Cusip_var = LpVariable.dicts("Cusips", Cusips, upBound=1.0, lowBound=0, cat='Integer')
    prob += lpSum([yields[i] * Cusip_var[i] for i in Cusip_var])
    if Portfolio is not None:
        EP = dict(zip(Cusips,df['In Portfolio']))
        prob += lpSum([EP[i] * Cusip_var[i] for i in Cusip_var]) == ebonds
    prob += lpSum([Invested[i] * Cusip_var[i] for i in Cusip_var]) == bonds
    prob += lpSum([Rating[i] * Cusip_var[i] for i in Cusip_var]) >= Arated
    prob += lpSum([Duration[i] * Cusip_var[i] for i in Cusip_var]) <= (duration * bonds)
    prob += lpSum([Seniority[i] * Cusip_var[i] for i in Cusip_var]) <= Subs
    if CS is not None:
        prob += lpSum([CS[i] * Cusip_var[i] for i in Cusip_var]) <= math.ceil(.2*bonds)
    if CD is not None:
        prob += lpSum([CD[i] * Cusip_var[i] for i in Cusip_var]) <= math.ceil(.2*bonds)
    if CST is not None:
        prob += lpSum([CST[i] * Cusip_var[i] for i in Cusip_var]) <= math.ceil(.2*bonds)
    if E is not None:
        prob += lpSum([E[i] * Cusip_var[i] for i in Cusip_var]) <= 1
    if M is not None:
        prob += lpSum([M[i] * Cusip_var[i] for i in Cusip_var]) <= 1
    if F is not None:
        prob += lpSum([F[i] * Cusip_var[i] for i in Cusip_var]) <= math.ceil(9/25*bonds)
    if HC is not None:
        prob += lpSum([HC[i] * Cusip_var[i] for i in Cusip_var]) <= math.ceil(.2*bonds)
    if I is not None:
        prob += lpSum([I[i] * Cusip_var[i] for i in Cusip_var]) <= math.ceil(.2*bonds)
    if IT is not None: 
       prob += lpSum([IT[i] * Cusip_var[i] for i in Cusip_var]) <= math.ceil(.2*bonds)
    if U is not None:
       prob += lpSum([U[i] * Cusip_var[i] for i in Cusip_var]) <= math.ceil(.15*bonds)
    g2 = datetime.now()
    ed = datetime.now().replace(month=12, day=31)
    if oneyearmature > 0:
        prob += oneyearmature == lpSum([Cusip_var[i] for i in Cusip_var if g2 <= Maturity[i] <= g2 + relativedelta(years=1)])
    for s in Symbols:
        prob += 1 >= lpSum([Cusip_var[i] for i in Cusip_var if Ticker[i] == s])
    prob += 1 >= lpSum([Cusip_var[i] for i in Cusip_var if Ticker[i] == 'CVS' or Ticker[i] =='WBA'])
    if Portfolio is None:
        for a in range(LadderStart, LadderLength):
            g = g2 + relativedelta(years=a)
            g1 = g2 + relativedelta(years=a + 1)
            rb = ed + relativedelta(years=a)
            ry = ed + relativedelta(years=a + 1)
            prob += 2 <= lpSum([Cusip_var[i] for i in Cusip_var if g <= Maturity[i] <= g1])
            prob += MaxPerYear >= lpSum([Cusip_var[i] for i in Cusip_var if g <= Maturity[i] <= g1])
            prob += 2 <= lpSum([Cusip_var[i] for i in Cusip_var if rb <= Maturity[i] <= ry])
    for a in range(LadderStart, LadderYears + 1):
        rb = ed + relativedelta(years=a)
        ry = ed + relativedelta(years=a + 1)
        prob += MaxPerYear >= lpSum([Cusip_var[i] for i in Cusip_var if rb <= Maturity[i] <= ry])
    prob += 0 == lpSum([Cusip_var[i] for i in Cusip_var if Maturity[i] >= g2 + relativedelta(years=maxladderlength)])
    prob.solve()
    portfolio = []
    for v in prob.variables():
        if v.varValue > 0:
            portfolio.append(v.name[-9:])
    OptimalPortfolio = ddf[ddf['CUSIP'].isin(portfolio)]
    if excel == 1:
        OptimalPortfolio.to_excel(
            "S:\AHB\Research\Fixed-Income\Weekly List\Optimal Portfolio\\" + now + " Bonds_" + str(
                bonds) + " Duration_" + str(duration) + " Subs_" + str(Subs) + ".xlsx")
    print(np.average(OptimalPortfolio['Ask Yield To Worst'])*100)
    return OptimalPortfolio