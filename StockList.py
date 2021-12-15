from numpy import result_type
import pandas as pd

today = '2021-12-16'
yesterday = '2021-12-14'


def is_stair(stair):
    # if stair['半年最低股價'] > stair['一年最低股價']:
    if stair['半年最高股價'] == stair['一年最高股價']:
        if stair['三個月最高股價'] == stair['半年最高股價']:
            if float(stair['成交']) >= float(stair['一年最高股價']) * 0.85:
                return 'TRUE'


def copy_type_N_from_yesterday(type_N):
    if type_N['類型'+'_'+yesterday] == 'N':
        return 'N'


stock_list_df = pd.read_csv('StockList_' + today + '.csv')
stock_list_df['強勢股'] = stock_list_df.apply(is_stair, axis=1)
stock_list_df = stock_list_df[stock_list_df['強勢股'] == 'TRUE']
stock_list_df['類型'] = stock_list_df.apply(lambda x: None, axis=1)
stock_list_df['名稱'] = stock_list_df['名稱'].astype(str)
stock_list_df['代號'] = stock_list_df.apply(
    lambda x: str(x['代號']).replace('=', ''), axis=1)
stock_list_df['代號'] = stock_list_df.apply(
    lambda x: str(x['代號']).replace('"', ''), axis=1)
stock_list_df['股票名稱_國泰'] = stock_list_df.apply(
    lambda x: x['名稱'] + '(' + str(x['代號']) + ')', axis=1)
stock_list_df['股票名稱_群益'] = stock_list_df.apply(
    lambda x: str(x['代號']) + x['名稱'], axis=1)
# stock_list_df.to_excel('temp.xlsx', index=False)

stock_list_yesterday_df = pd.read_excel('StockList_' + yesterday + '.xlsx')
stock_list_yesterday_df = stock_list_yesterday_df[['名稱', '類型'+'_'+yesterday]]
stock_list_yesterday_df.rename(
    columns={'類型'+'_'+yesterday: '類型'}, inplace=True)
merge_df = stock_list_df.merge(stock_list_yesterday_df, how='outer', on=[
                               '名稱'], indicator=True, suffixes=('_' + today, '_' + yesterday))
merge_df.rename(columns={'_merge': '_merge_with_yesterday'}, inplace=True)
merge_df['類型'+'_' + today] = merge_df.apply(copy_type_N_from_yesterday, axis=1)
merge_df.to_excel('StockList_' + today + '.xlsx', index=False)
stock_list_df = pd.read_excel('StockList_' + today + '.xlsx')

# 未平倉
open_position_df = pd.read_excel('TodayOpenPosition_' + today + '.xlsx')
open_position_df.rename(
    columns={'股票名稱': '股票名稱_國泰', '\t庫存股數': '庫存股數_國泰', '\t持有成本': '持有成本_國泰', '\t未實現損益': '未實現損益_國泰'}, inplace=True)
open_position_df = open_position_df[[
    '股票名稱_國泰', '庫存股數_國泰', '持有成本_國泰', '未實現損益_國泰']]
merge_df = stock_list_df.merge(
    open_position_df, how='outer', on='股票名稱_國泰', indicator=True)
merge_df.rename(
    columns={'_merge': '_merge_with_OpenPosition_國泰'}, inplace=True)
merge_df.to_excel('StockList_' + today + '.xlsx', index=False)
stock_list_df = pd.read_excel('StockList_' + today + '.xlsx')

# 未平倉
open_position_df = pd.read_excel('TodayOpenPosition_' + today + '_.xlsx')
open_position_df.rename(
    columns={'股票名稱': '股票名稱_群益', '庫存股數': '庫存股數_群益', '付出成本': '付出成本_群益', '損益': '損益_群益'}, inplace=True)
open_position_df = open_position_df[['股票名稱_群益', '庫存股數_群益', '付出成本_群益', '損益_群益']]
merge_df = stock_list_df.merge(
    open_position_df, how='outer', on='股票名稱_群益', indicator=True)
merge_df.rename(
    columns={'_merge': '_merge_with_OpenPosition_群益'}, inplace=True)
merge_df.to_excel('StockList_' + today + '.xlsx', index=False)
stock_list_df = pd.read_excel('StockList_' + today + '.xlsx')

stock_list_df['本週創新高'] = stock_list_df.apply(
    lambda x: x['5日最高股價'] == x['一年最高股價'], axis=1)
stock_list_df['低於近5日低點'] = stock_list_df.apply(
    lambda x: x['成交'] <= x['5日最低股價'], axis=1)
stock_list_df['低於近10日低點'] = stock_list_df.apply(
    lambda x: x['成交'] <= x['10日最低股價'], axis=1)
stock_list_df.to_excel('StockList_' + today + '.xlsx', index=False)
stock_list_df = pd.read_excel('StockList_' + today + '.xlsx')

print_df = stock_list_df[stock_list_df['漲跌幅'] > 5]
print(print_df[['代號', '名稱', '漲跌幅']])
print('今日漲5%以上\n')

open_position_cost = float(
    stock_list_df['持有成本_國泰'].sum()) + float(stock_list_df['付出成本_群益'].sum())
print('今日未平倉成本：  $', open_position_cost)
open_position_profit = float(
    stock_list_df['未實現損益_國泰'].sum()) + float(stock_list_df['損益_群益'].sum())
open_position_profit_percent = round(
    open_position_profit * 100 / open_position_cost, 2)
print('今日未平倉績效：  $', open_position_profit,
      ' ... ', open_position_profit_percent, '%')
