import control_variables as cv

def log_to_csv(df, name):
    csv_name = cv.common_fname + '_' + name + '.csv'
    csv_path = '.\log\\' + csv_name
    df.to_csv(csv_path, encoding='utf_8_sig')