import workbook_manage as wm
from common_functions import get_balance_table, get_cashflow_table, get_profit_table

def get_sheet_name():
    return '现金流分析'

def profit_quality_data(p_df, c_df):
    df = c_df[['报表日期', '经营活动产生的现金流量净额']]
    df = df.join(p_df[['五、净利润']])
    # 写入excel
    wbm = wm.Workbook_Manage(get_sheet_name())
    wbm.write_dataframe(df)
    # 画折线图
    wbm.write_line_chart('利润质量分析', [1,df.shape[0]+1,2,3],[2,df.shape[0]+1,1],f'F1')

def income_growth_analysis_data(p_df, c_df):
    df = c_df[['报表日期', '销售商品、提供劳务收到的现金']]
    df = df.join(p_df[['营业收入']])
    # 写入excel
    wbm = wm.Workbook_Manage(get_sheet_name())
    wbm.write_dataframe(df)
    # 画折线图
    start = df.shape[0] + 3 + 1
    wbm.write_line_chart('营业收入增长分析', [start,df.shape[0]+start,2,3],[2,df.shape[0]+1,1],f'F{start}')

def cash_support_analysis_data(c_df, b_df):
    df = b_df[['报表日期', '短期借款', '交易性金融负债', '一年内到期的非流动负债', '长期借款', '应付债券']]
    df = df.join(c_df[['现金的期末余额', '投资活动现金流出小计', '分配股利、利润或偿付利息所支付的现金']])
    df['有息负债'] = df['短期借款'] + df['交易性金融负债'] + df['一年内到期的非流动负债'] + df['长期借款'] + df['应付债券']
    # 写入excel
    wbm = wm.Workbook_Manage(get_sheet_name())
    wbm.write_dataframe(df)
    # 画折线图
    start = (df.shape[0] + 3) * 2 + 1
    wbm.write_line_chart('现金是否足以支撑投资和筹资活动分析', [start,df.shape[0]+start,7,10],[2,df.shape[0]+1,1],f'L{start}')

def cash_flow_analysis_data():
    stock_balance_table_df = get_balance_table()
    stock_profit_table_df = get_profit_table()
    stock_cashflow_table_df = get_cashflow_table()
    profit_quality_data(stock_profit_table_df, stock_cashflow_table_df)
    income_growth_analysis_data(stock_profit_table_df, stock_cashflow_table_df)
    cash_support_analysis_data(stock_cashflow_table_df, stock_balance_table_df)

if __name__ == '__main__':
    cash_flow_analysis_data()