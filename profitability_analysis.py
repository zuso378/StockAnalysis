import numpy as np
import workbook_manage as wm
from common_functions import average_calc, get_balance_table, get_profit_table, ratio_calc

def get_sheet_name():
    return '盈利能力分析'

def profit_margin_calc(df):
    gross_margin_list = []
    op_margin_list = []
    for index, rows in df.iterrows():
        op_revenue = np.double(rows['营业收入'])
        op_cost = np.double(rows['营业成本'])
        gross_margin_list.append((op_revenue-op_cost)/op_revenue * 100)
        op_margin_list.append((op_revenue-op_cost-np.double(rows['销售费用'])-np.double(rows['管理费用'])-np.double(rows['财务费用'])-np.double(rows['研发费用']))/op_revenue * 100)
    return gross_margin_list, op_margin_list

def income_data(df):
    income_data_df = df[['报表日期', '营业收入', '营业成本', '销售费用', '管理费用', '财务费用', '研发费用', '五、净利润']]
    income_data_df['毛利率'], income_data_df['营业利润率'] = profit_margin_calc(df)
    income_data_df['净利率'] = ratio_calc(df, ['五、净利润'], ['营业收入'], 100)
    # 写入excel
    wbm = wm.Workbook_Manage(get_sheet_name())
    wbm.write_dataframe(income_data_df)
    # 画折线图
    wbm.write_line_chart('营业收入视角看盈利能力', [1,df.shape[0]+1,9,11],[2,df.shape[0]+1,1],f'N1')

def assets_data(df):
    stock_balance_table_df = get_balance_table()
    assets_data_df = df[['报表日期', '归属于母公司所有者的净利润', '五、净利润']]
    assets_data_df = assets_data_df.join(stock_balance_table_df[['归属于母公司股东权益合计', '资产总计']])
    assets_data_df['平均净资产'] = average_calc(assets_data_df, ['归属于母公司股东权益合计'])
    assets_data_df['平均总资产'] = average_calc(assets_data_df, ['资产总计'])
    assets_data_df['净资产收益率ROE'] = ratio_calc(assets_data_df, ['归属于母公司所有者的净利润'], ['平均净资产'], 100)
    assets_data_df['总资产收益率ROA'] = ratio_calc(assets_data_df, ['五、净利润'], ['平均总资产'], 100)
    # 写入excel
    wbm = wm.Workbook_Manage(get_sheet_name())
    wbm.write_dataframe(assets_data_df)
    # 画折线图
    start = df.shape[0] +3 + 1
    wbm.write_line_chart('资产视角看盈利能力', [start,df.shape[0]+start,8,9],[2,df.shape[0]+1,1],f'N{start}')

def profitablility_analysis_data():
    stock_profit_table_df = get_profit_table()
    income_data(stock_profit_table_df)
    assets_data(stock_profit_table_df)

if __name__ == '__main__':
    profitablility_analysis_data()