import pandas as pd

reference = 'GoodInfo_StockList_20211125.xlsx'
open_position = r'..\TradingModel_v2\TEMP_TradingModel_TotalOpenPosition.xlsx'

reference_df = pd.read_excel(reference)

open_position_df = pd.read_excel(open_position)
open_position_df.rename(columns={'Stock_Id': '代號'}, inplace=True)
open_to_drop_df = reference_df.merge(
    open_position_df, how='outer', on='代號', indicator=True).loc[lambda x: x['_merge'] != 'left_only']
print(open_to_drop_df[['代號', '名稱', 'Class', 'PositionSize', '_merge']])
