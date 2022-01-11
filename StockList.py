from numpy import result_type
import pandas as pd

account_cty = 0

today = '2022-01-11'
yesterday = '2022-01-06'


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
if account_cty == 1:
    stock_list_df['股票名稱_國泰'] = stock_list_df.apply(
        lambda x: x['名稱'] + '(' + str(x['代號']) + ')', axis=1)
stock_list_df['股票名稱_群益'] = stock_list_df.apply(
    lambda x: str(x['代號']) + x['名稱'], axis=1)
stock_list_df['股票名稱_群益'] = stock_list_df['股票名稱_群益'].astype('str')
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

if account_cty == 1:
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
open_position_df['股票名稱_群益'] = open_position_df['股票名稱_群益'].astype('str')
merge_df = stock_list_df.merge(
    open_position_df, how='outer', on='股票名稱_群益', indicator=True)
merge_df.rename(
    columns={'_merge': '_merge_with_OpenPosition_群益'}, inplace=True)
merge_df.to_excel('StockList_' + today + '.xlsx', index=False)
stock_list_df = pd.read_excel('StockList_' + today + '.xlsx')

stock_list_df['5日創新高'] = stock_list_df.apply(
    lambda x: x['5日最高股價'] == x['一年最高股價'], axis=1)
stock_list_df['10日創新高'] = stock_list_df.apply(
    lambda x: x['10日最高股價'] == x['一年最高股價'], axis=1)
stock_list_df['低於近5日低點'] = stock_list_df.apply(
    lambda x: x['成交'] <= x['5日最低股價'], axis=1)
stock_list_df['低於近10日低點'] = stock_list_df.apply(
    lambda x: x['成交'] <= x['10日最低股價'], axis=1)
stock_list_df.to_excel('StockList_' + today + '.xlsx', index=False)
stock_list_df = pd.read_excel('StockList_' + today + '.xlsx')

print_df = stock_list_df[stock_list_df['漲跌幅'] > 5]
print(print_df[['代號', '名稱', '漲跌幅']])
print('今日漲5%以上\n')

if account_cty == 1:
    open_position_cost = float(
        stock_list_df['持有成本_國泰'].sum()) + float(stock_list_df['付出成本_群益'].sum())
else:
    open_position_cost = float(stock_list_df['付出成本_群益'].sum())
print('今日未平倉成本：  $', open_position_cost)
if account_cty == 1:
    open_position_profit = float(
        stock_list_df['未實現損益_國泰'].sum()) + float(stock_list_df['損益_群益'].sum())
else:
    open_position_profit = float(stock_list_df['損益_群益'].sum())
try:
    open_position_profit_percent = round(
        open_position_profit * 100 / open_position_cost, 2)
    print('今日未平倉績效：  $', open_position_profit,
          ' ... ', open_position_profit_percent, '%')
except:
    print('今日未平倉績效：  N/A')


# 已平倉
try:
    close_position_df = pd.read_excel('TodayClosePosition_' + today + '_.xlsx')
    close_position_profit = float(close_position_df['損益'].sum())
    print('今日已平倉績效：  $', close_position_profit)
except:
    close_position_profit = 0
    print('今日已平倉績效：  N/A')
profit_df = pd.read_excel('profit.xlsx')
profit_df = profit_df.append(
    {'日期': today, '未平倉績效': str(open_position_profit_percent) + '%', '未平倉成本': str(open_position_cost), '未實現損益': str(open_position_profit), '已實現損益': str(close_position_profit)}, ignore_index=True)
profit_df.to_excel('profit.xlsx', index=False)
