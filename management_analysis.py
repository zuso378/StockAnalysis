from common_functions import average_calc, get_balance_table, get_profit_table, ratio_calc
import workbook_manage as wm

def get_sheet_name():
    return '管理层能力分析'

def turnover_rate_data():
    stock_profit_table_df = get_profit_table()
    stock_balance_table_df = get_balance_table()
    turnover_rate_data_df = stock_profit_table_df[['报表日期', '营业收入', '营业成本']]
    turnover_rate_data_df = turnover_rate_data_df.join(stock_balance_table_df[['应收账款', '存货', '固定资产净额', '资产总计']])
    turnover_rate_data_df['平均应收账款'] = average_calc(turnover_rate_data_df, ['应收账款'])
    turnover_rate_data_df['存货平均余额'] = average_calc(turnover_rate_data_df, ['存货'])
    turnover_rate_data_df['平均总资产'] = average_calc(turnover_rate_data_df, ['资产总计'])
    turnover_rate_data_df['应收账款周转率'] = ratio_calc(turnover_rate_data_df, ['营业收入'], ['平均应收账款'])
    turnover_rate_data_df['存货周转率'] = ratio_calc(turnover_rate_data_df, ['营业成本'], ['存货平均余额'])
    turnover_rate_data_df['固定资产周转率'] = ratio_calc(turnover_rate_data_df, ['营业收入'], ['固定资产净额'])
    turnover_rate_data_df['总资产周转率'] = ratio_calc(turnover_rate_data_df, ['营业收入'], ['平均总资产'])
    # 写入excel
    wbm = wm.Workbook_Manage(get_sheet_name())
    wbm.write_dataframe(turnover_rate_data_df)
    # 画折线图
    start = turnover_rate_data_df.shape[0]+ 3
    wbm.write_line_chart('应收账款周转率', [1,turnover_rate_data_df.shape[0]+1,11,11],[2,turnover_rate_data_df.shape[0]+1,1],f'A{start}')
    wbm.write_line_chart('存货周转率', [1,turnover_rate_data_df.shape[0]+1,12,12],[2,turnover_rate_data_df.shape[0]+1,1],f'J{start}')
    wbm.write_line_chart('固定资产周转率', [1,turnover_rate_data_df.shape[0]+1,13,13],[2,turnover_rate_data_df.shape[0]+1,1],f'A{start+18}')
    wbm.write_line_chart('总资产周转率', [1,turnover_rate_data_df.shape[0]+1,14,14],[2,turnover_rate_data_df.shape[0]+1,1],f'J{start+18}')

def management_analysis_data():
    turnover_rate_data()

if __name__ == '__main__':
    management_analysis_data()