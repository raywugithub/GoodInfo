import pandas as pd
from datetime import date
import datetime

date_string = '2021'+'1207'

today = date.today()
yesterday = today - datetime.timedelta(days=1)
#today = '2021-12-06'
#yesterday = '2021-12-02'

openposition = 'TodayOpenPosition_' + date_string + '.xlsx'
reference = 'GoodInfo_StockList_' + date_string + '.xlsx'


def change_name(temp):
    return (temp['名稱'] + '(' + str(temp['代號']) + ')')


def change_type(temp):
    if temp['_merge'] == 'left_only':
        return 'LOST'
    if temp['_merge'] == 'both':
        return 'KEEP'
    if temp['_merge'] == 'right_only':
        return 'NEW'


def drop_from_high(temp):
    if temp['Type'] == 'KEEP':
        return str(round(temp['\t現價'] / temp['一年最高股價'], 4) * 100) + '%'


open_position_df = pd.read_excel(openposition)

reference_df = pd.read_excel(reference)
reference_df['Name'] = reference_df.apply(change_name, axis=1)
reference_df = reference_df[[
    'Name', 'Class', '一年最高股價', '3日分數', '5日分數', '連續分數']]
reference_df.rename(columns={'Name': '股票名稱'}, inplace=True)

analysis_df = open_position_df.merge(
    reference_df, how='outer', indicator=True, on='股票名稱')
analysis_df['Type'] = analysis_df.apply(change_type, axis=1)
analysis_df['距離最高點'] = analysis_df.apply(drop_from_high, axis=1)
analysis_df.drop(columns={'_merge'}, inplace=True)
analysis_df.sort_values(by='Type', inplace=True)
analysis_df.to_excel('持股分析_' + str(today) + '.xlsx')

temp_analysis_df = analysis_df[analysis_df['Type'] == 'LOST']
yesterday_analysis_df = pd.read_excel('持股分析_' + str(yesterday) + '.xlsx')
yesterday_analysis_df = yesterday_analysis_df[yesterday_analysis_df['Type'] == 'LOST']

two_days_lost_df = temp_analysis_df.merge(
    yesterday_analysis_df, how='outer', indicator=True, on='股票名稱').loc[lambda x: x['_merge'] == 'both']
two_days_lost_df.to_excel('持股分析_' + str(today) + '_LOST.xlsx')

open_position_profit_df = pd.read_excel('未實現分析.xlsx')
open_position_profit = float(
    analysis_df['\t未實現損益'].sum()) / float(analysis_df['\t持有成本'].sum())
open_position_profit = round(open_position_profit*100, 2)
try:
    if str(open_position_profit_df['日期'].to_list()[-1]) != str(today):
        open_position_profit_df = open_position_profit_df.append(
            {'日期': str(today), '未實現報酬率': open_position_profit}, ignore_index=True)
        open_position_profit_df.to_excel('未實現分析.xlsx', index=False)
except:
    open_position_profit_df = open_position_profit_df.append(
        {'日期': str(today), '未實現報酬率': open_position_profit}, ignore_index=True)
    open_position_profit_df.to_excel('未實現分析.xlsx', index=False)
