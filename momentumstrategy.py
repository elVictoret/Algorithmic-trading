#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Fri Feb  4 22:14:51 2022

@author: victorgutierrezgutierrez
"""
import numpy as np
import pandas as pd
import requests 
import xlsxwriter 
import math 
from scipy.stats import percentileofscore as score 
from statistics import mean

stocks = pd.read_csv('sp_500_stocks_fixed.csv')
from secrets2 import IEX_CLOUD_API_TOKEN

def chunks(lst, n):
    """Yield successive n-sized chunks from lst."""
    for i in range(0, len(lst), n):
        yield lst[i:i + n]   
        
symbol_groups = list(chunks(stocks['Ticker'], 100))
symbol_strings = []
for i in range(0, len(symbol_groups)):
    symbol_strings.append(','.join(symbol_groups[i]))


hqm_columns = [
                'Ticker', 
                'Company name',
                'Price', 
                'Number of Shares to Buy', 
                'ROI',
                'One-Year Price Return', 
                'One-Year Return Percentile',
                'Six-Month Price Return',
                'Six-Month Return Percentile',
                'Three-Month Price Return',
                'Three-Month Return Percentile',
                'One-Month Price Return',
                'One-Month Return Percentile',
                'HQM Score'
                ]

hqm_dataframe = pd.DataFrame(columns = hqm_columns)

for symbol_string in symbol_strings:
    batch_api_call_url = f'https://sandbox.iexapis.com/stable/stock/market/batch/?types=stats,quote&symbols={symbol_string}&token={IEX_CLOUD_API_TOKEN}'
    data = requests.get(batch_api_call_url).json()
    for symbol in symbol_string.split(','):
        hqm_dataframe = hqm_dataframe.append(
                                        pd.Series([symbol, 
                                                   data[symbol]['stats']['companyName'],
                                                   data[symbol]['quote']['latestPrice'],
                                                   'N/A',
                                                   data[symbol]['stats']['peRatio'],
                                                   data[symbol]['stats']['year1ChangePercent'],
                                                   'N/A',
                                                   data[symbol]['stats']['month6ChangePercent'],
                                                   'N/A',
                                                   data[symbol]['stats']['month3ChangePercent'],
                                                   'N/A',
                                                   data[symbol]['stats']['month1ChangePercent'],
                                                   'N/A',
                                                   'N/A'
                                                   ], 
                                                  index = hqm_columns), 
                                        ignore_index = True)

# def portfolio_input():
#     global portfolio_size
#     portfolio_size = input ('Enter value of your portfolio:')
    
#     try:
#         val = float(portfolio_size)
#     except ValueError:
#         print("That's not a number! \n Try again:")
#         portfolio_size = input("Enter the value of your portfolio:")
    
# portfolio_input()
# float(portfolio_size)

portfolio_size = 1000000
position_size = float(portfolio_size) / hqm_dataframe.shape[0]

for i in range(hqm_dataframe.shape[0]):
    hqm_dataframe.loc[i,'Number of Shares to Buy'] = math.floor(position_size) / hqm_dataframe['Price'][i]
    
for column in hqm_dataframe.columns:
    hqm_dataframe.dropna(inplace=True)

# Deleting some missing data which caused problems

hqm_dataframe.sort_values('One-Year Price Return', ascending = False, inplace = True)
hqm_dataframe = hqm_dataframe[:50]
hqm_dataframe.reset_index(drop = True, inplace = True)

time_periods = [
                'One-Year',
                'Six-Month',
                'Three-Month',
                'One-Month'
                ]

for row in hqm_dataframe.index:
    for time_period in time_periods:
        hqm_dataframe.loc[row, f'{time_period} Return Percentile'] = score(hqm_dataframe[f'{time_period} Price Return'], hqm_dataframe.loc[row, f'{time_period} Price Return'])/100

for row in hqm_dataframe.index:
    momentum_percentiles=[]
    for time_period in time_periods:
        momentum_percentiles.append(hqm_dataframe.loc[row,f'{time_period} Return Percentile'])
    hqm_dataframe.loc[row, 'HQM Score'] = mean(momentum_percentiles)
    
for i in range(hqm_dataframe.shape[0]):
    hqm_dataframe.loc[i,'ROI']=100/hqm_dataframe['ROI'][i]

hqm_dataframe.sort_values('HQM Score', ascending = False, inplace = True)
hqm_dataframe.reset_index(drop = True, inplace = True)

writer = pd.ExcelWriter('recommended_trades_momentum.xlsx', engine='xlsxwriter')
hqm_dataframe.to_excel(writer, sheet_name='Recommended Trades', index = False)

background_color = '#0a0a23'
font_color = '#ffffff'

title_format = writer.book.add_format(
        {
            'font_color': font_color,
            'font_size': 10,
            'bg_color': background_color,
            'border': 1,
            'align': 'center'
        }
    )

string_format = writer.book.add_format(
        {
            'font_color': font_color,
            'bg_color': background_color,
            'border': 1,
            'align': 'center'
        }
    )

dollar_format = writer.book.add_format(
        {
            'num_format':'$0.00',
            'font_color': font_color,
            'bg_color': background_color,
            'border': 1,
            'align': 'center'
        }
    )


decimal_format = writer.book.add_format(
        {
            'num_format':'0.00',
            'font_color': font_color,
            'bg_color': background_color,
            'border': 1,
            'align': 'center'
        }
    )

integer_format = writer.book.add_format(
        {
            'num_format':'0',
            'font_color': font_color,
            'bg_color': background_color,
            'border': 1,
            'align': 'center'
        }
    )

percentage_format = writer.book.add_format(
        {
            'num_format':'0.00%',
            'font_color': font_color,
            'bg_color': background_color,
            'border': 1,
            'align': 'center'
        }
    )

column_formats = { 
                    'A': ['Ticker', string_format],
                    'B': ['Company Name', string_format],
                    'C': ['Price', dollar_format],
                    'D': ['Number of Shares to Buy', integer_format],
                    'E': ['ROI', percentage_format],
                    'F': ['One-Year Price Return',decimal_format],
                    'G': ['One-Year Return Percentile', decimal_format],
                    'H': ['Six-Month Price Return', decimal_format],
                    'I': ['Six-Month Return Percentile', decimal_format],
                    'J': ['Three-Month Price Return', decimal_format],
                    'K': ['Three-Month Return Percentile', decimal_format],
                    'L': ['One-Month Price Return', decimal_format],
                    'M': ['One-Month Return Percentile', decimal_format],
                    'N': ['HQM Score', string_format
                ]
                    }

for column in column_formats.keys():
    writer.sheets['Recommended Trades'].set_column(f'{column}:{column}', 30, column_formats[column][1])
    writer.sheets['Recommended Trades'].write(f'{column}1', column_formats[column][0], title_format)
    
writer.save()



