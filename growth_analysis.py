from common_functions import get_balance_table, get_profit_table, growth_rate_calc
import workbook_manage as wm

def get_sheet_name():
    return '成长性分析'

def income_growth_data():
    stock_profit_table_df = get_profit_table()
    income_growth_data_df = stock_profit_table_df[['报表日期', '营业收入', '三、营业利润']]
    income_growth_data_df['营业收入增长率'] = growth_rate_calc(income_growth_data_df['营业收入'])
    income_growth_data_df['营业利润增长率'] = growth_rate_calc(income_growth_data_df['三、营业利润'])
    # 写入excel
    wbm = wm.Workbook_Manage(get_sheet_name())
    wbm.write_dataframe(income_growth_data_df)
    # 画折线图
    wbm.write_line_chart('营业收入视角看成长能力', [1,income_growth_data_df.shape[0]+1,4,5],[2,income_growth_data_df.shape[0]+1,1],f'H1')

def assets_growth_data():
    stock_balance_table_df = get_balance_table()
    assets_growth_data_df = stock_balance_table_df[['报表日期', '资产总计', '所有者权益(或股东权益)合计']]
    assets_growth_data_df['资产增长率'] = growth_rate_calc(assets_growth_data_df['资产总计'])
    assets_growth_data_df['净资产增长率'] = growth_rate_calc(assets_growth_data_df['所有者权益(或股东权益)合计'])
    # 写入excel
    wbm = wm.Workbook_Manage(get_sheet_name())
    wbm.write_dataframe(assets_growth_data_df)
    # 画折线图
    start = assets_growth_data_df.shape[0] + 4
    wbm.write_line_chart('资产视角看成长能力', [start,assets_growth_data_df.shape[0]+start,4,5],[2,assets_growth_data_df.shape[0]+1,1],f'H{start}')

if __name__ == '__main__':
    income_growth_data()
    assets_growth_data()
