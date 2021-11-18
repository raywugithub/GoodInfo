import pandas as pd


def stock_id_transfer(transfer):
    result = str(transfer['代號']).replace('=', '')
    return result.replace('"', '')


df = pd.read_csv('GoodInfo_StockList_20211118.csv')

df['代號'] = df.apply(stock_id_transfer, axis=1)
filter_df = pd.read_excel('Filter_20211117.xlsx')
merged_df = df.merge(filter_df, how='outer',
                     indicator=True, on=['代號', '名稱']).loc[lambda x: x['_merge'] == 'left_only']
merged_df.drop(columns={'_merge'}, inplace=True)
merged_df.to_excel('temp_GoodInfo_StockList_20211118.xlsx')


df = pd.read_excel('temp_GoodInfo_StockList_20211118.xlsx')
filter_df = pd.read_excel('GoodInfo_StockList_20211117_.xlsx')
merged_df = df.merge(filter_df, how='outer',
                     indicator=True, on=['名稱'])
merged_df.drop(columns={'_merge'}, inplace=True)
merged_df.to_excel('temp_GoodInfo_StockList_20211118.xlsx')


df = pd.read_excel('temp_GoodInfo_StockList_20211118.xlsx')
filter_df = pd.read_excel('Filter_20211118.xlsx')
merged_df = df.merge(filter_df, how='outer',
                     indicator=True, on=['名稱']).loc[lambda x: x['_merge'] == 'left_only']
merged_df.drop(columns={'_merge'}, inplace=True)
merged_df.to_excel('temp_GoodInfo_StockList_20211118.xlsx')
