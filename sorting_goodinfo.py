import pandas as pd

# 4分以上 : 整理平台不明顯，直線上漲
# 3分以上 : 有整理平台，階梯上漲
# 2.5以上 : 突破整理平台
# 2分以上 : 有整理平台，在整理區間內
# 0分 : 不符合需求

# 0.8以上 : 週K歷史新高位階

sorted = False
input_csv = 'GoodInfo_StockList_20211117.csv'
output_xlsx = 'GoodInfo_StockList_20211117_.xlsx'
reference_xlsx = ''


def stair_condition(stair_df):
    if (stair_df['半年最低股價'] > stair_df['一年最低股價']):
        if (stair_df['三個月最低股價'] > stair_df['半年最低股價']):
            if(stair_df['一個月最低股價'] > stair_df['三個月最低股價']):
                return 'YES'


if not sorted:
    goodinfo_stocklist = pd.read_csv(input_csv)

    # 代號
    # 名稱
    # 成交
    # 漲跌幅
    # 一個月最低股價
    # 三個月最低股價
    # 半年最低股價
    # 一年最低股價
    # 5日累計漲跌(%)
    # 一個月累計漲跌(%)
    # 三個月累計漲跌(%)
    # 半年累計漲跌(%)
    # 一年累計漲跌(%)
    # 21W45外資買賣超張數
    # 21W45外資買賣超佔成交(%)
    # 一個月最高股價
    # 三個月最高股價
    # 半年最高股價
    # 一年最高股價
    sorted_df = pd.DataFrame()
    for n in range(len(goodinfo_stocklist['代號'])):
        goodinfo_stocklist['Condition'] = goodinfo_stocklist.apply(
            stair_condition, axis=1)
    goodinfo_stocklist = goodinfo_stocklist[goodinfo_stocklist['Condition'] == 'YES']
    goodinfo_stocklist = goodinfo_stocklist[[
        '代號', '名稱', '成交', '一年累計漲跌(%)']]
    goodinfo_stocklist.to_excel(output_xlsx)

    reference_df = pd.read_excel(reference_xlsx)
    #sorted_df = pd.read_excel(output_xlsx)
    reference_df = reference_df[['代號', '名稱', 'Class']]
    merging_df = pd.merge(goodinfo_stocklist, reference_df,
                          how='outer', on=['名稱'])
    merging_df['代號_x'] = merging_df['代號_y']
    merging_df.rename(columns={'代號_x': '代號'}, inplace=True)
    merging_df.sort_values(by='Class', inplace=True, ascending=False)
    merging_df = merging_df[[
        '代號', '名稱', '5日累計漲跌(%)', '三個月累計漲跌(%)', '半年累計漲跌(%)', '一年累計漲跌(%)', 'Class']]
    merging_df.to_excel(output_xlsx, index=False)
    print(merging_df)
else:
    print('merged')
    merged_df = pd.read_excel(reference_xlsx)
    print(merged_df.sort_values(by='Class'))
