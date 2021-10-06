import numpy as np
import workbook_manage as wm
from common_functions import average_calc, get_balance_table, get_profit_table

def get_sheet_name():
    return '盈利能力分析'

def profit_margin_calc(df):
    gross_margin_list = []
    op_margin_list = []
    net_margin_list = []
    for index, rows in df.iterrows():
        op_revenue = np.double(rows['营业收入'])
        op_cost = np.double(rows['营业成本'])
        gross_margin_list.append((op_revenue-op_cost)/op_revenue * 100)
        op_margin_list.append((op_revenue-op_cost-np.double(rows['销售费用'])-np.double(rows['管理费用'])-np.double(rows['财务费用'])-np.double(rows['研发费用']))/op_revenue * 100)
        net_margin_list.append(np.double(rows['五、净利润'])/op_revenue * 100)
    return gross_margin_list, op_margin_list, net_margin_list

def income_data(df):
    income_data_df = df[['报表日期', '营业收入', '营业成本', '销售费用', '管理费用', '财务费用', '研发费用', '五、净利润']]
    income_data_df['毛利率'], income_data_df['营业利润率'], income_data_df['净利率'] = profit_margin_calc(df)
    # 写入excel
    wbm = wm.Workbook_Manage(get_sheet_name())
    wbm.write_dataframe(income_data_df)
    # 画折线图
    wbm.write_line_chart('营业收入视角看盈利能力', [1,df.shape[0]+1,9,11],[2,df.shape[0]+1,1],f'N1')

#    净资产收益率 = 净利润 ÷ 平均净资产　　　　　　　　　　　　　　　超过20%佳       [^平均资产 = (期末值 + 期初值) ÷ 2]   
#    总资产收益率 = 净利润 ÷ 平均总资产　　　　　　　　　　　　　　　净资产收益率相同，总资产收益率更高，说明有更强的盈利能力，并承受更小的风险   

def return_rate_calc(df):
    roe_list = []
    roa_list = []
    for index, rows in df.iterrows():
        roe_list.append(np.double(rows['归属于母公司所有者的净利润'])/rows['平均净资产'] * 100)
        roa_list.append(np.double(rows['五、净利润'])/rows['平均总资产'] * 100)
    return roe_list, roa_list

def assets_data(df):
    stock_balance_table_df = get_balance_table()
    assets_data_df = df[['报表日期', '归属于母公司所有者的净利润', '五、净利润']]
    assets_data_df = assets_data_df.join(stock_balance_table_df[['归属于母公司股东权益合计', '资产总计']])
    assets_data_df['平均净资产'] = average_calc(assets_data_df, ['归属于母公司股东权益合计'])
    assets_data_df['平均总资产'] = average_calc(assets_data_df, ['资产总计'])
    assets_data_df['净资产收益率ROE'], assets_data_df['总资产收益率ROA'] = return_rate_calc(assets_data_df)
    # 写入excel
    wbm = wm.Workbook_Manage(get_sheet_name())
    wbm.write_dataframe(assets_data_df)
    # 画折线图
    start = df.shape[0] +3 + 1
    wbm.write_line_chart('资产视角看盈利能力', [start,df.shape[0]+start,8,9],[2,df.shape[0]+1,1],f'N{start}')

if __name__ == '__main__':
    stock_profit_table_df = get_profit_table()
    income_data(stock_profit_table_df)
    assets_data(stock_profit_table_df)