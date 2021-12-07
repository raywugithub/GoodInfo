import pandas as pd


def stock_id_transfer(transfer):
    result = str(transfer['代號']).replace('=', '')
    return result.replace('"', '')


def change_name(temp):
    return (temp['名稱'] + '(' + str(temp['代號']) + ')')


#StockList_temp_df = pd.read_excel('GoodInfo_StockList_20211207.xlsx')
#StockList_temp_df['代號'] = StockList_temp_df.apply(stock_id_transfer, axis=1)
#StockList_temp_df['股票名稱'] = StockList_temp_df.apply(change_name, axis=1)
StockList_temp_df = pd.read_excel('持股分析_2021-12-07.xlsx')

temp_df = pd.read_excel('temp.xlsx')

merge_df = temp_df.merge(StockList_temp_df, how='outer',
                         on='股票名稱', indicator=True).loc[lambda x: x['_merge'] != 'left_only']
merge_df.to_excel('主力買賣超_20211207.xlsx')
