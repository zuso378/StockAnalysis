import numpy as np
import workbook_manage as wm
from common_functions import get_balance_table, get_cashflow_table, log_to_csv

def get_sheet_name():
    return '安全性分析'

def average_cash_equivalents_calc(df):
    average_list = []
    for index, rows in df.iterrows():
        average_list.append((np.double(rows['加:期初现金及现金等价物余额'])+np.double(rows['六、期末现金及现金等价物余额']))/2)
    return average_list

def debt_with_interest_calc(df):
    pv_s = np.double(df.iloc[0]['短期借款']) + np.double(df.iloc[0]['交易性金融负债']) + np.double(df.iloc[0]['一年内到期的非流动负债'])
    pv_l = pv_s + np.double(df.iloc[0]['长期借款']) + np.double(df.iloc[0]['应付债券'])
    average_s_list = []
    average_l_list = []
    for index, rows in df.iterrows():
        v_s = np.double(rows['短期借款']) + np.double(rows['交易性金融负债']) + np.double(rows['一年内到期的非流动负债'])
        v_l = v_s + np.double(rows['长期借款']) + np.double(rows['应付债券'])
        average_s_list.append((pv_s+v_s)/2)
        average_l_list.append((pv_l+v_l)/2)
        pv_s = v_s
        pv_l = v_l
    return average_l_list, average_s_list

def assets_to_cash_calc(df):
    pv = np.double(df.iloc[0]['交易性金融资产']) + np.double(df.iloc[0]['应收票据'])
    average_list = []
    for index, rows in df.iterrows():
        v = np.double(rows['交易性金融资产']) + np.double(rows['应收票据'])
        average_list.append((pv+v)/2)
        pv = v
    return average_list

def cash_debt_ratio_calc(df):
    ratio_list = []
    for index, rows in df.iterrows():
        ratio_list.append(rows['现金及现金等价物']/rows['有息负债'])
    return ratio_list

def cash_debt_ratio_calc1(df):
    ratio_list = []
    for index, rows in df.iterrows():
        ratio_list.append((rows['现金及现金等价物']+rows['可迅速变现的金融资产净值'])/rows['有息负债'])
    return ratio_list

def cash_debt_ratio_calc2(df):
    ratio_list = []
    for index, rows in df.iterrows():
        ratio_list.append((rows['现金及现金等价物']+rows['可迅速变现的金融资产净值'])/rows['一年内到期的有息负债'])
    return ratio_list

def cash_debt_ratio_data():
    cf_df = get_cashflow_table()
    b_df = get_balance_table()
    cash_debt_ratio_df = cf_df[['报表日期', '加:期初现金及现金等价物余额', '六、期末现金及现金等价物余额']]
    cash_debt_ratio_df = cash_debt_ratio_df.join(b_df[['短期借款', '交易性金融负债', '一年内到期的非流动负债', '长期借款', '应付债券', '交易性金融资产', '应收票据']])
    cash_debt_ratio_df['现金及现金等价物'] = average_cash_equivalents_calc(cash_debt_ratio_df)
    cash_debt_ratio_df['有息负债'], cash_debt_ratio_df['一年内到期的有息负债'] = debt_with_interest_calc(cash_debt_ratio_df)
    cash_debt_ratio_df['可迅速变现的金融资产净值'] = assets_to_cash_calc(cash_debt_ratio_df)
    cash_debt_ratio_df['现金债务比'] = cash_debt_ratio_calc(cash_debt_ratio_df)
    cash_debt_ratio_df['现金债务比1'] = cash_debt_ratio_calc1(cash_debt_ratio_df)
    cash_debt_ratio_df['现金债务比2'] = cash_debt_ratio_calc2(cash_debt_ratio_df)
    # 写入excel
    wbm = wm.Workbook_Manage(get_sheet_name())
    wbm.write_dataframe(cash_debt_ratio_df)
    # 画折线图
    row_count = cash_debt_ratio_df.shape[0]
    wbm.write_line_chart('现金债务比', [1,row_count+1,15,17],[2,row_count+1,1],f'A{row_count+3}')

if __name__ == '__main__':
    cash_debt_ratio_data()