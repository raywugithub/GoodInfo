import pandas as pd
from datetime import date
import datetime

date_string = '2021'+'1201'

today = date.today()
yesterday = today - datetime.timedelta(days=1)

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
reference_df = reference_df[['Name', 'Class', '一年最高股價']]
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
