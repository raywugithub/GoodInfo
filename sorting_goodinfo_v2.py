import pandas as pd

date_string = '2021'+'1203'

filter = 'Filter_20211202.xlsx'
input = 'GoodInfo_StockList_' + date_string + '.csv'
reference = 'GoodInfo_StockList_20211202.xlsx'
output = 'GoodInfo_StockList_' + date_string + '.xlsx'


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


df = pd.read_csv(input)

df['代號'] = df.apply(stock_id_transfer, axis=1)
filter_df = pd.read_excel(filter)
filter_df['代號'] = filter_df['代號'].astype('str')
merged_df = df.merge(filter_df, how='outer',
                     indicator=True, on=['代號', '名稱']).loc[lambda x: x['_merge'] == 'left_only']
#merged_df.drop(columns={'_merge'}, inplace=True)
merged_df = merged_df[['代號', '名稱', 'Class', '一年最高股價', '3日累計漲跌(%)', '5日累計漲跌(%)',
                       '三大法人3日累計買賣超佔成交(%)', '三大法人5日累計買賣超佔成交(%)', '券資比3日增減', '券資比5日增減']]
merged_df.to_excel(output)
df = pd.read_excel(output)
filter_df = pd.read_excel(reference)
filter_df['代號'] = filter_df['代號'].astype('object')
filter_df.drop(columns={'merge_status'}, inplace=True)
# filter_df.drop(columns={'一年最高股價'}, inplace=True)  # test
filter_df = filter_df[['名稱', 'Class', '3日分數', '5日分數']]
merged_df = df.merge(filter_df, how='outer', on=['名稱'],
                     indicator=True).loc[lambda x: x['_merge'] != 'right_only']
merged_df.rename(inplace=True, columns={'_merge': 'merge_status'})
#merged_df.drop(columns={'_merge'}, inplace=True)

merged_df['3日分數'] = merged_df.apply(three_days_score_calculate, axis=1)
merged_df['5日分數'] = merged_df.apply(five_days_score_calculate, axis=1)
merged_df.to_excel(output)
