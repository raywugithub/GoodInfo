import pandas as pd


def change_name(temp):
    return (temp['名稱'] + '(' + str(temp['代號']) + ')')


def stock_id_transfer(transfer):
    result = str(transfer['代號']).replace('=', '')
    return result.replace('"', '')


def three_days_score_calculate(temp):
    score = 0
    if temp['3日累計漲跌(%)'] > 0:
        score = score + 2
    elif temp['3日累計漲跌(%)'] < 0:
        score = score - 2
    else:
        score = score
    if temp['三大法人3日累計買賣超佔成交(%)'] > 0:
        score = score + 2
    elif temp['三大法人3日累計買賣超佔成交(%)'] < 0:
        score = score - 2
    else:
        score = score
    if temp['券資比3日增減'] > 0:
        score = score + 1
    elif temp['券資比3日增減'] < 0:
        score = score - 1
    else:
        score = score
    return score


def five_days_score_calculate(temp):
    score = 0
    if temp['5日累計漲跌(%)'] > 0:
        score = score + 2
    elif temp['5日累計漲跌(%)'] < 0:
        score = score - 2
    else:
        score = score
    if temp['三大法人5日累計買賣超佔成交(%)'] > 0:
        score = score + 2
    elif temp['三大法人5日累計買賣超佔成交(%)'] < 0:
        score = score - 2
    else:
        score = score
    if temp['券資比5日增減'] > 0:
        score = score + 1
    elif temp['券資比5日增減'] < 0:
        score = score - 1
    else:
        score = score
    return score


def ten_percent(temp):
    return round(temp['一年最高股價_y'] * 0.9, 2)


lost_df = pd.read_csv('lost.csv')
lost_df['代號'] = lost_df.apply(stock_id_transfer, axis=1)
lost_df['股票名稱'] = lost_df.apply(change_name, axis=1)

open_lost_df = pd.read_excel('持股分析_2021-12-08.xlsx')
open_lost_df.drop(columns={'5日累計漲跌(%)'}, inplace=True)
merge_df = open_lost_df.merge(
    lost_df, how='outer', on='股票名稱', indicator=True).loc[lambda x: x['_merge'] == 'both']
merge_df['3日分數'] = merge_df.apply(three_days_score_calculate, axis=1)
merge_df['5日分數'] = merge_df.apply(five_days_score_calculate, axis=1)
merge_df['10%'] = merge_df.apply(ten_percent, axis=1)

merge_df.to_excel('lost.xlsx')
