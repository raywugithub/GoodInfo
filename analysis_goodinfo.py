import pandas as pd
from datetime import date

today = date.today()

openposition = 'TodayOpenPosition_20211130.xlsx'
reference = 'GoodInfo_StockList_20211130.xlsx'


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
