import akshare as ak
import control_variables as cv

def log_to_csv(df, name):
    csv_name = cv.common_fname + '_' + name + '.csv'
    csv_path = '.\log\\' + csv_name
    df.to_csv(csv_path, encoding='utf_8_sig')

def get_profit_table():
    stock_profit_table_df = ak.stock_financial_report_sina(stock=cv.stock_code, symbol="利润表")
    log_to_csv(stock_profit_table_df, '利润表')
    year_df = stock_profit_table_df[stock_profit_table_df['报表日期'].str.contains('1231')].sort_values(by='报表日期')
    year_df['报表日期'] = year_df['报表日期'].str[0:4]
    log_to_csv(year_df, '利润表_年报')
    return year_df

def get_balance_table():
    stock_balance_table_df = ak.stock_financial_report_sina(stock=cv.stock_code, symbol="资产负债表")
    log_to_csv(stock_balance_table_df, '资产负债表')
    year_df = stock_balance_table_df[stock_balance_table_df['报表日期'].str.contains('1231')].sort_values(by='报表日期')
    year_df['报表日期'] = year_df['报表日期'].str[0:4]
    log_to_csv(year_df, '资产负债表_年报')
    return year_df

def get_cashflow_table():
    stock_cashflow_table_df = ak.stock_financial_report_sina(stock=cv.stock_code, symbol="现金流量表")
    log_to_csv(stock_cashflow_table_df, '现金流量表')
    year_df = stock_cashflow_table_df[stock_cashflow_table_df['报表日期'].str.contains('1231')].sort_values(by='报表日期')
    year_df['报表日期'] = year_df['报表日期'].str[0:4]
    log_to_csv(year_df, '现金流量表_年报')
    return year_df