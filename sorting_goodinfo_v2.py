import pandas as pd

date_string = '2021'+'1201'

filter = 'Filter_20211130.xlsx'
input = 'GoodInfo_StockList_' + date_string + '.csv'
reference = 'GoodInfo_StockList_20211130.xlsx'
output = 'GoodInfo_StockList_' + date_string + '.xlsx'


def stock_id_transfer(transfer):
    result = str(transfer['代號']).replace('=', '')
    return result.replace('"', '')


df = pd.read_csv(input)

df['代號'] = df.apply(stock_id_transfer, axis=1)
filter_df = pd.read_excel(filter)
filter_df['代號'] = filter_df['代號'].astype('str')
merged_df = df.merge(filter_df, how='outer',
                     indicator=True, on=['代號', '名稱']).loc[lambda x: x['_merge'] == 'left_only']
#merged_df.drop(columns={'_merge'}, inplace=True)
merged_df = merged_df[['代號', '名稱', 'Class', '一年最高股價']]
merged_df.to_excel(output)

df = pd.read_excel(output)
filter_df = pd.read_excel(reference)
filter_df['代號'] = filter_df['代號'].astype('object')
filter_df.drop(columns={'merge_status'}, inplace=True)
merged_df = df.merge(filter_df, how='outer', on=['名稱'],
                     indicator=True).loc[lambda x: x['_merge'] != 'right_only']
merged_df.rename(inplace=True, columns={'_merge': 'merge_status'})
#merged_df.drop(columns={'_merge'}, inplace=True)
merged_df.to_excel(output)
