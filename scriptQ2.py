import pandas as pd

xls = pd.ExcelFile('Technical Interview Questions 4-5-24.xlsx')

# printing last day of every month listed 
q2_UP = xls.parse(2)
q2_NOSHFF = xls.parse(3)
q2_NOSH = xls.parse(4)

# UP = Unadjusted price
q2_UP['Dates'] = pd.to_datetime(q2_UP['Dates'])
# skips december 2019
q2_UP = q2_UP[q2_UP['Dates'] >= '2020-01-01']
q2_UP.set_index('Dates', inplace=True)
up_lastDOM = q2_UP.resample('M').last()
up_lastDOM.reset_index(inplace=True)

# NOSHFF = Float factor
q2_NOSHFF['Dates'] = pd.to_datetime(q2_NOSHFF['Dates'])
# skips december 2019
q2_NOSHFF = q2_NOSHFF[q2_NOSHFF['Dates'] >= '2020-01-01']
q2_NOSHFF.set_index('Dates', inplace=True)
noshff_lastDOM = q2_NOSHFF.resample('M').last()
noshff_lastDOM.reset_index(inplace=True)

# NOSH = Number of shares
q2_NOSH['Dates'] = pd.to_datetime(q2_NOSH['Dates'])
q2_NOSH.set_index('Dates', inplace=True)
nosh_lastDOM = q2_NOSH.resample('M').last()
nosh_lastDOM.reset_index(inplace=True)

# identifies last day of each month present in all sheets
common_dates = set(up_lastDOM['Dates']).intersection(noshff_lastDOM['Dates']).intersection(nosh_lastDOM['Dates'])

# filters data based on common dates
up_lastDOM = up_lastDOM[up_lastDOM['Dates'].isin(common_dates)]
noshff_lastDOM = noshff_lastDOM[noshff_lastDOM['Dates'].isin(common_dates)]
nosh_lastDOM = nosh_lastDOM[nosh_lastDOM['Dates'].isin(common_dates)]

# melts data to have long format for easier merge
up_lastDOM_long = up_lastDOM.melt(id_vars='Dates', var_name='Stock', value_name='UP')
noshff_lastDOM_long = noshff_lastDOM.melt(id_vars='Dates', var_name='Stock', value_name='NOSHFF')
nosh_lastDOM_long = nosh_lastDOM.melt(id_vars='Dates', var_name='Stock', value_name='NOSH')

# merge frames on Dates and stock
marketcap = pd.merge(up_lastDOM_long, noshff_lastDOM_long, on=['Dates', 'Stock'])
marketcap = pd.merge(marketcap, nosh_lastDOM_long, on=['Dates', 'Stock'])

# Market Cap Formula
marketcap['Market Cap'] = marketcap['UP'] * marketcap['NOSHFF'] * marketcap['NOSH']

# save to new sheet
q2solution = 'Q2_marketcap.xlsx'
marketcap.to_excel(q2solution, index=False)

print(marketcap)
