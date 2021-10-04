import akshare as ak
import control_variables as cv

def log_to_csv(df, name):
    csv_name = cv.common_fname + '_' + name + '.csv'
    csv_path = '.\log\\' + csv_name
    df.to_csv(csv_path, encoding='utf_8_sig')

def get_profit_table():
    stock_profit_table_df = ak.stock_financial_report_sina(stock=cv.stock_code, symbol="利润表")
    log_to_csv(stock_profit_table_df, '利润表')
    return stock_profit_table_df[stock_profit_table_df['报表日期'].str.contains('1231')]

def get_balance_table():
    stock_financial_report_sina_df = ak.stock_financial_report_sina(stock=cv.stock_code, symbol="资产负债表")
    log_to_csv(stock_financial_report_sina_df, '资产负债表')
    return stock_financial_report_sina_df[stock_financial_report_sina_df['报表日期'].str.contains('1231')]