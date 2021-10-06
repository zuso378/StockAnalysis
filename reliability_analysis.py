import numpy as np
import workbook_manage as wm
from common_functions import average_calc, get_balance_table, get_cashflow_table, log_to_csv, ratio_calc

def get_sheet_name():
    return '安全性分析'

def average_cash_equivalents_calc(df):
    average_list = []
    for index, rows in df.iterrows():
        average_list.append((np.double(rows['加:期初现金及现金等价物余额'])+np.double(rows['六、期末现金及现金等价物余额']))/2)
    return average_list

def cash_debt_ratio_data():
    cf_df = get_cashflow_table()
    b_df = get_balance_table()
    cash_debt_ratio_df = cf_df[['报表日期', '加:期初现金及现金等价物余额', '六、期末现金及现金等价物余额']]
    cash_debt_ratio_df = cash_debt_ratio_df.join(b_df[['短期借款', '交易性金融负债', '一年内到期的非流动负债', '长期借款', '应付债券', '交易性金融资产', '应收票据']])
    cash_debt_ratio_df['现金及现金等价物'] = average_cash_equivalents_calc(cash_debt_ratio_df)
    cash_debt_ratio_df['有息负债'] = average_calc(cash_debt_ratio_df, ['短期借款', '交易性金融负债', '一年内到期的非流动负债'])
    cash_debt_ratio_df['一年内到期的有息负债'] = average_calc(cash_debt_ratio_df, ['短期借款', '交易性金融负债', '一年内到期的非流动负债', '长期借款', '应付债券'])
    cash_debt_ratio_df['可迅速变现的金融资产净值'] = average_calc(cash_debt_ratio_df, ['交易性金融资产', '应收票据'])
    cash_debt_ratio_df['现金债务比'] = ratio_calc(cash_debt_ratio_df, ['现金及现金等价物'], ['有息负债'])
    cash_debt_ratio_df['现金债务比1'] = ratio_calc(cash_debt_ratio_df, ['现金及现金等价物','可迅速变现的金融资产净值'], ['有息负债'])
    cash_debt_ratio_df['现金债务比2'] = ratio_calc(cash_debt_ratio_df, ['现金及现金等价物','可迅速变现的金融资产净值'], ['一年内到期的有息负债'])
    # 写入excel
    wbm = wm.Workbook_Manage(get_sheet_name())
    wbm.write_dataframe(cash_debt_ratio_df)
    # 画折线图
    row_count = cash_debt_ratio_df.shape[0]
    wbm.write_line_chart('现金债务比', [1,row_count+1,15,17],[2,row_count+1,1],f'A{row_count+3}')

if __name__ == '__main__':
    cash_debt_ratio_data()