import pandas as pd


def is_stair(stair):
    # if stair['半年最低股價'] > stair['一年最低股價']:
    if stair['半年最高股價'] == stair['一年最高股價']:
        if stair['三個月最高股價'] == stair['半年最高股價']:
            if float(stair['成交']) >= float(stair['一年最高股價']) * 0.85:
                return 'TRUE'


stock_list_df = pd.read_csv('GoodInfo_StockList_20211209__.csv')
stock_list_df['強勢股'] = stock_list_df.apply(is_stair, axis=1)
stock_list_df = stock_list_df[stock_list_df['強勢股'] == 'TRUE']
stock_list_df['類型'] = stock_list_df.apply(lambda x: None, axis=1)
stock_list_df['代號'] = stock_list_df.apply(
    lambda x: str(x['代號']).replace('=', ''), axis=1)
stock_list_df['代號'] = stock_list_df.apply(
    lambda x: str(x['代號']).replace('"', ''), axis=1)
stock_list_df['股票名稱'] = stock_list_df.apply(
    lambda x: x['名稱'] + '(' + str(x['代號']) + ')', axis=1)
# stock_list_df.to_excel('temp.xlsx', index=False)

open_position_df = pd.read_excel('TodayOpenPosition_20211209.xlsx')

merge_df = stock_list_df.merge(
    open_position_df, how='outer', on='股票名稱', indicator=True)
merge_df['本週創新高'] = merge_df.apply(
    lambda x: x['5日最高股價'] == x['一年最高股價'], axis=1)
merge_df['低於近5日低點'] = merge_df.apply(
    lambda x: x['成交'] <= x['5日最低股價'], axis=1)
merge_df['低於近10日低點'] = merge_df.apply(
    lambda x: x['成交'] <= x['10日最低股價'], axis=1)
# merge_df.to_excel('temp.xlsx', index=False)


out_df = pd.read_excel('temp.xlsx')
out_df = out_df[out_df['類型'] == 'X']
print(out_df['股票名稱'])
