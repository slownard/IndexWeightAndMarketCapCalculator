import pandas as pd

# reads excel file
xls = pd.ExcelFile('Technical Interview Questions 4-5-24.xlsx')

# shows Q1 Marketcap sheet 
q1_Data = xls.parse(1)

# looks for specific column name
q1_date = 'date'
q1_ticker = 'ticker'
q1_marketcap = 'marketcap'
q1_weight = 'weight'

# gets rid of duplicates
unique_dates = q1_Data[q1_date].drop_duplicates()

# counts how many different dates there are
num_unique_dates = len(unique_dates)

# counts sum of unique tickes per date group
sum_of_unique_tickes = q1_Data.groupby(q1_date)[q1_ticker].nunique() 

# gets total marketcap for every date group
sum_marketcap_perdate = q1_Data.groupby(q1_date)[q1_marketcap].sum()

date_groups = q1_Data.groupby(q1_date)
for date, group in date_groups:
    total_market_cap = group[q1_marketcap].sum()
    q1_Data.loc[group.index, q1_weight] = (group[q1_marketcap] / total_market_cap)

    # THIS prints weight for every constituent for every rebalance period
    # print(f"For date {date}:")
    # print(group[[q1_ticker, 'weight']])

# updated sheet created
calculatedweight = 'Q1_calculatedweight.xlsx'

# this adds the weight to new sheet
q1_Data.to_excel(calculatedweight, index = False)

print("Updated DataFrame with 'index_weight' column:")
print(q1_Data)


# TESTS // uncomment to test data 
# print("Number of unique dates in the column:", num_unique_dates)
# print("Sum of unique tickers with the same date:", sum_of_unique_tickes)
# print("Sum of market caps within each group by date:",sum_marketcap_perdate)