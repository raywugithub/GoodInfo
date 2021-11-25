import pandas as pd

filter = 'Filter_20211124.xlsx'
input = 'GoodInfo_StockList_20211125.csv'
reference = 'GoodInfo_StockList_20211124.xlsx'
output = 'GoodInfo_StockList_20211125.xlsx'
drop = 'GoodInfo_StockList_20211125_drop.xlsx'
open_position = r'..\TradingModel_v2\TEMP_TradingModel_TotalOpenPosition.xlsx'


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
merged_df = merged_df[['代號', '名稱', 'Class']]
merged_df.to_excel(output)

df = pd.read_excel(output)
filter_df = pd.read_excel(reference)
filter_df['代號'] = filter_df['代號'].astype('object')
filter_df.drop(columns={'merge_status'}, inplace=True)
merged_df = df.merge(filter_df, how='outer', on=['代號', '名稱'],
                     indicator=True).loc[lambda x: x['_merge'] != 'right_only']
merged_df.rename(inplace=True, columns={'_merge': 'merge_status'})
#merged_df.drop(columns={'_merge'}, inplace=True)
merged_df.to_excel(output)


drop_df = df.merge(filter_df, how='outer', on=['代號', '名稱'],
                   indicator=True).loc[lambda x: x['_merge'] == 'right_only']
drop_df.to_excel(drop)

open_position_df = pd.read_excel(open_position)
open_position_df.rename(columns={'Stock_Id': '代號'}, inplace=True)
open_to_drop_df = drop_df.merge(open_position_df, how='inner', on='代號')
print('\n\n***OpenPosition to Drop List***\n', open_to_drop_df[['代號', '名稱']])
