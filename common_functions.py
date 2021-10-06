import akshare as ak
import numpy as np
import control_variables as cv

def growth_rate_calc(sr):
    rate_list = []
    pv = np.double(sr[0:1])
    for value in sr.values:
        v = np.double(value)
        if 0==pv:
            rate = 0
        else:
            rate = (v - pv) / pv
        pv = v
        rate_list.append(rate*100)
    return rate_list

def average_calc(df, ls):
    avr_list = []
    p_v = 0
    for item in ls:
        p_v += np.double(df.iloc[0][item])
    for index, rows in df.iterrows():
        v = 0
        for item in ls:
            v += np.double(rows[item])
        avr_list.append((p_v+v)/2)
        p_v = v
    return avr_list

def ratio_calc(df, n_ls, d_ls, prct=1):
    ratio_list = []
    for index, rows in df.iterrows():
        n = 0
        d = 0
        for item in n_ls:
            n += np.double(rows[item])
        for item in d_ls:
            d += np.double(rows[item])
        ratio_list.append(n/d * prct)
    return ratio_list
        


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