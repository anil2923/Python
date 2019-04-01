import pandas as pd
import numpy as np

def each_insert(each_row):
    return 'INSERT INTO ATRate VALUES ({0}, {1}, \'{2}\', \'{3}\', \'{4}\', \'{5}\', \'{6}\')'.format(each_row[0], each_row[1], each_row[2], each_row[3], each_row[4], each_row[5], each_row[6])

def each_sheet(each_record):
    data_frame_each = pd.read_excel('ATRates.xlsx', sheet_name=each_record[0], converters={'Id': np.int32, 'GroupId': np.int32, 'RateType': np.float64, 'NACEGroup': np.float64, 'NACEClass': np.float64, 'Region': np.float64, 'Salary': np.float64, 'RiskCategory': np.float64})
    dml_start = '------------------->> INSET SCRIPT STARTED FOR : {0}<<-------------------\n'.format(each_record[0])
    dml_middle = '\n'.join(map(each_insert, [(data_frame_each['GroupId'][i], each_record[1], data_frame_each['NACEGroup'][i], data_frame_each['NACEClass'][i], data_frame_each['Region'][i], data_frame_each['Salary'][i], data_frame_each['RiskCategory'][i]) for i in data_frame_each.index]))
    dml_end = '\n------------------->>INSERT SCRIPT ENDED FOR : {0}<<-------------------\n\n'.format(each_record[0])
    return dml_start + dml_middle + dml_end


with open("AT_RATE_UPD2.sql", "w+") as file:
    data_frame = pd.read_excel('ATRates.xlsx', sheet_name='ATRateTypes', converters={'RateType': np.int32, 'RateName': str})
    final_output = '\n'.join(map(each_sheet, [(data_frame['RateName'][i], data_frame['RateType'][i]) for i in data_frame.index]))
    file.write(final_output)