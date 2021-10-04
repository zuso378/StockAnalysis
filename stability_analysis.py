import akshare as ak
from openpyxl.worksheet import worksheet
import pandas as pd
import numpy as np
from common_functions import get_profit_table, get_balance_table
import control_variables as cv
import workbook_manage as wm

def get_sheet_name():
    return '安全性分析'

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


def balance_assets_data(df, sr):
    balance_assets_data_df = df[['报表日期', '应收票据', '应收账款', '其他应收款', '长期待摊费用']].sort_values(by='报表日期')
    # 报表日期只需要保留年份数据
    balance_assets_data_df['报表日期'] = balance_assets_data_df['报表日期'].str[0:4]
    # 营业收入
    balance_assets_data_df['营业收入'] = sr
    # 计算 应收票据 增长率
    balance_assets_data_df['应收票据增长率'] = growth_rate_calc(balance_assets_data_df['应收票据'])
    # 计算 应收账款 增长率
    balance_assets_data_df['应收账款增长率'] = growth_rate_calc(balance_assets_data_df['应收账款'])
    # 计算 其他应收款 增长率
    balance_assets_data_df['其他应收款增长率'] = growth_rate_calc(balance_assets_data_df['其他应收款'])
    # 计算 长期待摊费用 增长率
    balance_assets_data_df['长期待摊费用增长率'] = growth_rate_calc(balance_assets_data_df['长期待摊费用'])
    # 计算 营业收入 增长率
    balance_assets_data_df['营业收入增长率'] = growth_rate_calc(balance_assets_data_df['营业收入'])
    # 写入excel
    wbm = wm.Workbook_Manage(get_sheet_name())
    wbm.write_dataframe(balance_assets_data_df)
    # 画折线图
    wbm.write_line_chart('资产端数据增长率', [1,df.shape[0]+1,7,11],[2,df.shape[0]+1,1],'N1')

def balance_liability_data(df, sr):
    balance_assets_data_df = df[['报表日期', '短期借款', '应付票据', '应付账款', '其他应付款']].sort_values(by='报表日期')
    # 报表日期只需要保留年份数据
    balance_assets_data_df['报表日期'] = balance_assets_data_df['报表日期'].str[0:4]
    # 营业收入
    balance_assets_data_df['营业收入'] = sr
    # 计算 应收票据 增长率
    balance_assets_data_df['短期借款增长率'] = growth_rate_calc(balance_assets_data_df['短期借款'])
    # 计算 应收账款 增长率
    balance_assets_data_df['应付票据增长率'] = growth_rate_calc(balance_assets_data_df['应付票据'])
    # 计算 其他应收款 增长率
    balance_assets_data_df['应付账款增长率'] = growth_rate_calc(balance_assets_data_df['应付账款'])
    # 计算 长期待摊费用 增长率
    balance_assets_data_df['其他应付款增长率'] = growth_rate_calc(balance_assets_data_df['其他应付款'])
    # 计算 营业收入 增长率
    balance_assets_data_df['营业收入增长率'] = growth_rate_calc(balance_assets_data_df['营业收入'])
    # 写入excel
    wbm = wm.Workbook_Manage(get_sheet_name())
    wbm.write_dataframe(balance_assets_data_df)
    # 画折线图
    start = df.shape[0] + 3 + 1
    wbm.write_line_chart('负债端数据增长率', [start,df.shape[0]+start,7,11],[2,df.shape[0]+1,1],f'N{start}')

def profit_compare_calc(df):
    oper_rate_list = []
    net_ratio_list = []
    three_ratio_list = []
    for index, rows in df.iterrows():
        profit_from_operation = np.double(rows['三、营业利润'])
        oper_rate_list.append(profit_from_operation/np.double(rows['营业收入']) * 100)
        net_ratio_list.append(np.double(rows['五、净利润'])/profit_from_operation * 100)
        three_ratio_list.append((np.double(rows['资产减值损失'])+np.double(rows['公允价值变动收益'])+np.double(rows['投资收益']))/profit_from_operation * 100)
    return oper_rate_list, net_ratio_list, three_ratio_list

def profits_data(df):
    profits_data_df = df[['报表日期', '五、净利润', '三、营业利润', '营业收入', '资产减值损失', '公允价值变动收益', '投资收益']].sort_values(by='报表日期')
    # 报表日期只需要保留年份数据
    profits_data_df['报表日期'] = profits_data_df['报表日期'].str[0:4]
    # 营业利润率   ，  净利润/营业利润  ，  (资产减值损失+公允价值变动收益+投资收益)/营业利润
    profits_data_df['营业利润率'], profits_data_df['净利占比'] , profits_data_df['三项占比']= profit_compare_calc(profits_data_df)
    # 写入excel
    wbm = wm.Workbook_Manage(get_sheet_name())
    wbm.write_dataframe(profits_data_df)
    # 画折线图
    start = (df.shape[0] + 3) * 2 + 1
    wbm.write_line_chart('利润端数据稳定性', [start,df.shape[0]+start,8,10],[2,df.shape[0]+1,1],f'N{start}')
    wbm.write_line_chart('营业利润率稳定性', [start,df.shape[0]+start,8,8],[2,df.shape[0]+1,1],f'W{start}')
    wbm.write_line_chart('净利占比稳定性', [start,df.shape[0]+start,9,9],[2,df.shape[0]+1,1],f'AE{start}')
    wbm.write_line_chart('三项占比稳定性', [start,df.shape[0]+start,10,10],[2,df.shape[0]+1,1],f'AN{start}')


if __name__ == '__main__':
    stock_profit_table_df = get_profit_table()
    stock_balance_table_df = get_balance_table()
    balance_assets_data(stock_balance_table_df, stock_profit_table_df['营业收入'])
    balance_liability_data(stock_balance_table_df, stock_profit_table_df['营业收入'])
    profits_data(stock_profit_table_df)


