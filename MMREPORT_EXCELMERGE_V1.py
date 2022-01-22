from collections import defaultdict
import time
from openpyxl.workbook.workbook import Workbook
from pandas.core.base import DataError
from pandas.core.frame import DataFrame
import pygetwindow as gw
import pyautogui
import datetime
from pywinauto import Application
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
from itertools import islice
import pandas as pd
import numpy as np
import xlrd
import win32com.client as win32
import glob
import os
import sys
import pywinauto

# sys.stdout = open('Terminal_FTA_Merge.txt', 'w',encoding='UTF-8')





# #XLS -> XLSX로 변환
# for i in glob.glob("MMreport_ROH\*.xls"):

#     fname = i
#     excel = win32.gencache.EnsureDispatch('Excel.Application')
#     wb = excel.Workbooks.Open(fname)

#     wb.SaveAs(fname+"x", FileFormat = 51) # fileformat 51이 xlsx뜻함
#     wb.Close()
# excel.Application.Quit()

# #우선 모든 자료들을 FTA_ROH에 저장
# all_data = pd.DataFrame()
# all_data1 = pd.DataFrame()
# for f in glob.glob("MMreport_ROH\*.xlsx"): #해당 디렉토리에 있는 xlsx파일들 불러옴
#     df=pd.read_excel(f, sheet_name=None, engine='openpyxl')
#     all_data1 = pd.concat(df, ignore_index=True)
#     all_data=all_data.append(all_data1,ignore_index=True)
#     print(all_data)
# all_data.to_excel('MMreport_merged_ROH_v5.xlsx',encoding='utf-8-sig') #저장
# #엑셀로 저장하는게 오래 걸릴 때에는, csv로 저장할 수도 있습니다.
# #all_data.to_csv('합본.csv',encoding='utf-8-sig') 

# df_FERT = pd.read_excel('MMreport_merged_ROH_v5.xlsx', engine='openpyxl')
# print(df_FERT)
# condition = df_FERT['Unnamed: 1'] =='ROH'
# print(df_FERT[condition])
# df_FERT_ONLYdata = df_FERT[condition]
# df_FERT_ONLYdata.to_excel('MMreport_ROH_onlydata_v5.xlsx',encoding='utf-8-sig')

# df_ROH_ONLYdata = pd.read_excel('MMreport_ROH_onlydata_v5.xlsx', engine='openpyxl')
# df_ROH_SUMIF = df_ROH_ONLYdata.groupby('Unnamed: 2')[['Unnamed: 5','Unnamed: 6','Unnamed: 7','Unnamed: 8','Unnamed: 9','Unnamed: 10','Unnamed: 11','Unnamed: 12','Unnamed: 13','Unnamed: 14','Unnamed: 15','Unnamed: 16','Unnamed: 17', 'Unnamed: 18','Unnamed: 19','Unnamed: 20','Unnamed: 21','Unnamed: 22','Unnamed: 23','Unnamed: 24','Unnamed: 25','Unnamed: 26','Unnamed: 27','Unnamed: 28']].sum()
# df_ROH_SUMIF.to_excel('MMreport_ROH_SUMIF_v5.xlsx',encoding='utf-8-sig')

# XLS -> XLSX로 변환
# # for i in glob.glob("MMreport_HAWA\*.xls"):

# #     fname = i
# #     excel = win32.gencache.EnsureDispatch('Excel.Application')
# #     wb = excel.Workbooks.Open(fname)

# #     wb.SaveAs(fname+"x", FileFormat = 51) # fileformat 51이 xlsx뜻함
# #     wb.Close()
# # excel.Application.Quit()

# # #우선 모든 자료들을 FTA_ROH에 저장
# all_data = pd.DataFrame()
# all_data1 = pd.DataFrame()
# for f in glob.glob("MMreport_HAWA\*.xlsx"): #해당 디렉토리에 있는 xlsx파일들 불러옴
#     df=pd.read_excel(f, sheet_name=None, engine='openpyxl')
#     all_data1 = pd.concat(df, ignore_index=True)
#     all_data=all_data.append(all_data1,ignore_index=True)
#     print(all_data)
# all_data.to_excel('MMreport_merged_HAWA_V4.xlsx',encoding='utf-8-sig') #저장
# #엑셀로 저장하는게 오래 걸릴 때에는, csv로 저장할 수도 있습니다.
# #all_data.to_csv('합본.csv',encoding='utf-8-sig') 

# df_HAWA = pd.read_excel('MMreport_merged_HAWA_V4.xlsx', engine='openpyxl')
# print(df_HAWA)
# condition = df_HAWA['Unnamed: 1'] =='HAWA'
# print(df_HAWA[condition])
# df_HAWA_ONLYdata = df_HAWA[condition]
# df_HAWA_ONLYdata.to_excel('MMreport_HAWA_onlydata_v4.xlsx',encoding='utf-8-sig')

# df_HAWA_ONLYdata = pd.read_excel('MMreport_HAWA_onlydata_v4.xlsx', engine='openpyxl')
# df_HAWA_SUMIF = df_HAWA_ONLYdata.groupby('Unnamed: 2')[['Unnamed: 5','Unnamed: 6','Unnamed: 7','Unnamed: 8','Unnamed: 9','Unnamed: 10','Unnamed: 11','Unnamed: 12','Unnamed: 13','Unnamed: 14','Unnamed: 15','Unnamed: 16','Unnamed: 17', 'Unnamed: 18','Unnamed: 19','Unnamed: 20','Unnamed: 21','Unnamed: 22','Unnamed: 23','Unnamed: 24','Unnamed: 25','Unnamed: 26','Unnamed: 27','Unnamed: 28']].sum()
# df_HAWA_SUMIF.to_excel('MMreport_HAWA_SUMIF_v4.xlsx',encoding='utf-8-sig')





















# # XLS -> XLSX로 변환
# for i in glob.glob("MMreport_FERT\*.xls"):

#     fname = i
#     excel = win32.gencache.EnsureDispatch('Excel.Application')
#     wb = excel.Workbooks.Open(fname)

#     wb.SaveAs(fname+"x", FileFormat = 51) # fileformat 51이 xlsx뜻함
#     wb.Close()
# excel.Application.Quit()

# # 우선 모든 자료들을 FTA_ROH에 저장
# all_data = pd.DataFrame()
# all_data1 = pd.DataFrame()
# for f in glob.glob("MMreport_FERT\*.xlsx"): #해당 디렉토리에 있는 xlsx파일들 불러옴
#     df=pd.read_excel(f, sheet_name=None, engine='openpyxl')
#     all_data1 = pd.concat(df, ignore_index=True)
#     all_data=all_data.append(all_data1,ignore_index=True)
#     print(all_data)
# all_data.to_excel('MMreport_merged_FERT_V5.xlsx',encoding='utf-8-sig') #저장
# #엑셀로 저장하는게 오래 걸릴 때에는, csv로 저장할 수도 있습니다.
# #all_data.to_csv('합본.csv',encoding='utf-8-sig') 

# df_FERT = pd.read_excel('MMreport_merged_FERT_V5.xlsx', engine='openpyxl')
# print(df_FERT)
# condition = df_FERT['Unnamed: 1'] =='FERT'
# print(df_FERT[condition])
# df_FERT_ONLYdata = df_FERT[condition]
# df_FERT_ONLYdata.to_excel('MMreport_FERT_onlydata_V5.xlsx',encoding='utf-8-sig')

# df_FERT_ONLYdata = pd.read_excel('MMreport_FERT_onlydata_V5.xlsx', engine='openpyxl')
# df_FERT_SUMIF = df_FERT_ONLYdata.groupby('Unnamed: 2')[['Unnamed: 5','Unnamed: 6','Unnamed: 7','Unnamed: 8','Unnamed: 9','Unnamed: 10','Unnamed: 11','Unnamed: 12','Unnamed: 13','Unnamed: 14','Unnamed: 15','Unnamed: 16','Unnamed: 17', 'Unnamed: 18','Unnamed: 19','Unnamed: 20','Unnamed: 21','Unnamed: 22','Unnamed: 23','Unnamed: 24','Unnamed: 25','Unnamed: 26','Unnamed: 27','Unnamed: 28']].sum()
# df_FERT_SUMIF.to_excel('MMreport_FERT_SUMIF_V5.xlsx',encoding='utf-8-sig')












# # XLS -> XLSX로 변환
# for i in glob.glob("MMreport_HALB\*.xls"):

#     fname = i
#     excel = win32.gencache.EnsureDispatch('Excel.Application')
#     wb = excel.Workbooks.Open(fname)

#     wb.SaveAs(fname+"x", FileFormat = 51) # fileformat 51이 xlsx뜻함
#     wb.Close()
# excel.Application.Quit()

# # 우선 모든 자료들을 FTA_ROH에 저장
# all_data = pd.DataFrame()
# all_data1 = pd.DataFrame()
# for f in glob.glob("MMreport_HALB\*.xlsx"): #해당 디렉토리에 있는 xlsx파일들 불러옴
#     df=pd.read_excel(f, sheet_name=None, engine='openpyxl')
#     all_data1 = pd.concat(df, ignore_index=True)
#     all_data=all_data.append(all_data1,ignore_index=True)
#     print(all_data)
# all_data.to_excel('MMreport_merged_HALB_V5.xlsx',encoding='utf-8-sig') #저장
# #엑셀로 저장하는게 오래 걸릴 때에는, csv로 저장할 수도 있습니다.
# #all_data.to_csv('합본.csv',encoding='utf-8-sig') 

# df_FERT = pd.read_excel('MMreport_merged_HALB_V5.xlsx', engine='openpyxl')
# print(df_FERT)
# condition = df_FERT['Unnamed: 1'] =='HALB'
# print(df_FERT[condition])
# df_FERT_ONLYdata = df_FERT[condition]
# df_FERT_ONLYdata.to_excel('MMreport_HALB_onlydata_V5.xlsx',encoding='utf-8-sig')

# df_FERT_ONLYdata = pd.read_excel('MMreport_HALB_onlydata_V5.xlsx', engine='openpyxl')
# df_FERT_SUMIF = df_FERT_ONLYdata.groupby('Unnamed: 2')[['Unnamed: 5','Unnamed: 6','Unnamed: 7','Unnamed: 8','Unnamed: 9','Unnamed: 10','Unnamed: 11','Unnamed: 12','Unnamed: 13','Unnamed: 14','Unnamed: 15','Unnamed: 16','Unnamed: 17', 'Unnamed: 18','Unnamed: 19','Unnamed: 20','Unnamed: 21','Unnamed: 22','Unnamed: 23','Unnamed: 24','Unnamed: 25','Unnamed: 26','Unnamed: 27','Unnamed: 28']].sum()
# df_FERT_SUMIF.to_excel('MMreport_HALB_SUMIF_V5.xlsx',encoding='utf-8-sig')



# #XLS -> XLSX로 변환
# for i in glob.glob("ROH_2ND\*.xls"):

#     fname = i
#     excel = win32.gencache.EnsureDispatch('Excel.Application')
#     wb = excel.Workbooks.Open(fname)

#     wb.SaveAs(fname+"x", FileFormat = 51) # fileformat 51이 xlsx뜻함
#     wb.Close()
# excel.Application.Quit()

# #우선 모든 자료들을 FTA_ROH에 저장
# all_data = pd.DataFrame()
# all_data1 = pd.DataFrame()
# for f in glob.glob("ROH_2ND\*.xlsx"): #해당 디렉토리에 있는 xlsx파일들 불러옴
#     df=pd.read_excel(f, sheet_name=None, engine='openpyxl')
#     all_data1 = pd.concat(df, ignore_index=True)
#     all_data=all_data.append(all_data1,ignore_index=True)
#     print(all_data)
# all_data.to_excel('MMreport_merged_ROH_2ND.xlsx',encoding='utf-8-sig') #저장
# #엑셀로 저장하는게 오래 걸릴 때에는, csv로 저장할 수도 있습니다.
# #all_data.to_csv('합본.csv',encoding='utf-8-sig') 

# df_FERT = pd.read_excel('MMreport_merged_ROH_2ND.xlsx', engine='openpyxl')
# print(df_FERT)
# condition = df_FERT['Unnamed: 1'] =='ROH'
# print(df_FERT[condition])
# df_FERT_ONLYdata = df_FERT[condition]
# df_FERT_ONLYdata.to_excel('MMreport_ROH_onlydata_2ND.xlsx',encoding='utf-8-sig')

# df_ROH_ONLYdata = pd.read_excel('MMreport_ROH_onlydata_2ND.xlsx', engine='openpyxl')
# df_ROH_SUMIF = df_ROH_ONLYdata.groupby('Unnamed: 2')[['Unnamed: 5','Unnamed: 6','Unnamed: 7','Unnamed: 8','Unnamed: 9','Unnamed: 10','Unnamed: 11','Unnamed: 12','Unnamed: 13','Unnamed: 14','Unnamed: 15','Unnamed: 16','Unnamed: 17', 'Unnamed: 18','Unnamed: 19','Unnamed: 20','Unnamed: 21','Unnamed: 22','Unnamed: 23','Unnamed: 24','Unnamed: 25','Unnamed: 26','Unnamed: 27','Unnamed: 28']].sum()
# df_ROH_SUMIF.to_excel('MMreport_ROH_SUMIF_2ND.xlsx',encoding='utf-8-sig')



##XLS -> XLSX로 변환
# for i in glob.glob("HAWA_2ND\*.xls"):

#     fname = i
#     excel = win32.gencache.EnsureDispatch('Excel.Application')
#     wb = excel.Workbooks.Open(fname)

#     wb.SaveAs(fname+"x", FileFormat = 51) # fileformat 51이 xlsx뜻함
#     wb.Close()
# excel.Application.Quit()

# #우선 모든 자료들을 FTA_ROH에 저장
all_data = pd.DataFrame()
all_data1 = pd.DataFrame()
for f in glob.glob("HAWA_2ND\*.xlsx"): #해당 디렉토리에 있는 xlsx파일들 불러옴
    df=pd.read_excel(f, sheet_name=None, engine='openpyxl')
    all_data1 = pd.concat(df, ignore_index=True)
    all_data=all_data.append(all_data1,ignore_index=True)
    print(all_data)
all_data.to_excel('MMreport_merged_HAWA_2ND.xlsx',encoding='utf-8-sig') #저장
#엑셀로 저장하는게 오래 걸릴 때에는, csv로 저장할 수도 있습니다.
#all_data.to_csv('합본.csv',encoding='utf-8-sig') 

df_HAWA = pd.read_excel('MMreport_merged_HAWA_2ND.xlsx', engine='openpyxl')
print(df_HAWA)
condition = df_HAWA['Unnamed: 1'] =='HAWA'
print(df_HAWA[condition])
df_HAWA_ONLYdata = df_HAWA[condition]
df_HAWA_ONLYdata.to_excel('MMreport_HAWA_onlydata_2ND.xlsx',encoding='utf-8-sig')

df_HAWA_ONLYdata = pd.read_excel('MMreport_HAWA_onlydata_2ND.xlsx', engine='openpyxl')
df_HAWA_SUMIF = df_HAWA_ONLYdata.groupby('Unnamed: 2')[['Unnamed: 5','Unnamed: 6','Unnamed: 7','Unnamed: 8','Unnamed: 9','Unnamed: 10','Unnamed: 11','Unnamed: 12','Unnamed: 13','Unnamed: 14','Unnamed: 15','Unnamed: 16','Unnamed: 17', 'Unnamed: 18','Unnamed: 19','Unnamed: 20','Unnamed: 21','Unnamed: 22','Unnamed: 23','Unnamed: 24','Unnamed: 25','Unnamed: 26','Unnamed: 27','Unnamed: 28']].sum()
df_HAWA_SUMIF.to_excel('MMreport_HAWA_SUMIF_2ND.xlsx',encoding='utf-8-sig')





















# # XLS -> XLSX로 변환
# for i in glob.glob("FERT_2ND\*.xls"):

#     fname = i
#     excel = win32.gencache.EnsureDispatch('Excel.Application')
#     wb = excel.Workbooks.Open(fname)

#     wb.SaveAs(fname+"x", FileFormat = 51) # fileformat 51이 xlsx뜻함
#     wb.Close()
# excel.Application.Quit()

# # 우선 모든 자료들을 FTA_ROH에 저장
# all_data = pd.DataFrame()
# all_data1 = pd.DataFrame()
# for f in glob.glob("FERT_2ND\*.xlsx"): #해당 디렉토리에 있는 xlsx파일들 불러옴
#     df=pd.read_excel(f, sheet_name=None, engine='openpyxl')
#     all_data1 = pd.concat(df, ignore_index=True)
#     all_data=all_data.append(all_data1,ignore_index=True)
#     print(all_data)
# all_data.to_excel('MMreport_merged_FERT_2ND.xlsx',encoding='utf-8-sig') #저장
# #엑셀로 저장하는게 오래 걸릴 때에는, csv로 저장할 수도 있습니다.
# #all_data.to_csv('합본.csv',encoding='utf-8-sig') 

# df_FERT = pd.read_excel('MMreport_merged_FERT_2ND.xlsx', engine='openpyxl')
# print(df_FERT)
# condition = df_FERT['Unnamed: 1'] =='FERT'
# print(df_FERT[condition])
# df_FERT_ONLYdata = df_FERT[condition]
# df_FERT_ONLYdata.to_excel('MMreport_FERT_onlydata_2ND.xlsx',encoding='utf-8-sig')

# df_FERT_ONLYdata = pd.read_excel('MMreport_FERT_onlydata_2ND.xlsx', engine='openpyxl')
# df_FERT_SUMIF = df_FERT_ONLYdata.groupby('Unnamed: 2')[['Unnamed: 5','Unnamed: 6','Unnamed: 7','Unnamed: 8','Unnamed: 9','Unnamed: 10','Unnamed: 11','Unnamed: 12','Unnamed: 13','Unnamed: 14','Unnamed: 15','Unnamed: 16','Unnamed: 17', 'Unnamed: 18','Unnamed: 19','Unnamed: 20','Unnamed: 21','Unnamed: 22','Unnamed: 23','Unnamed: 24','Unnamed: 25','Unnamed: 26','Unnamed: 27','Unnamed: 28']].sum()
# df_FERT_SUMIF.to_excel('MMreport_FERT_SUMIF_2ND.xlsx',encoding='utf-8-sig')












# # XLS -> XLSX로 변환
# for i in glob.glob("HALB_2ND\*.xls"):

#     fname = i
#     excel = win32.gencache.EnsureDispatch('Excel.Application')
#     wb = excel.Workbooks.Open(fname)

#     wb.SaveAs(fname+"x", FileFormat = 51) # fileformat 51이 xlsx뜻함
#     wb.Close()
# excel.Application.Quit()

# # 우선 모든 자료들을 FTA_ROH에 저장
# all_data = pd.DataFrame()
# all_data1 = pd.DataFrame()
# for f in glob.glob("HALB_2ND\*.xlsx"): #해당 디렉토리에 있는 xlsx파일들 불러옴
#     df=pd.read_excel(f, sheet_name=None, engine='openpyxl')
#     all_data1 = pd.concat(df, ignore_index=True)
#     all_data=all_data.append(all_data1,ignore_index=True)
#     print(all_data)
# all_data.to_excel('MMreport_merged_HALB_2ND.xlsx',encoding='utf-8-sig') #저장
# #엑셀로 저장하는게 오래 걸릴 때에는, csv로 저장할 수도 있습니다.
# #all_data.to_csv('합본.csv',encoding='utf-8-sig') 

# df_FERT = pd.read_excel('MMreport_merged_HALB_2ND.xlsx', engine='openpyxl')
# print(df_FERT)
# condition = df_FERT['Unnamed: 1'] =='HALB'
# print(df_FERT[condition])
# df_FERT_ONLYdata = df_FERT[condition]
# df_FERT_ONLYdata.to_excel('MMreport_HALB_onlydata_2ND.xlsx',encoding='utf-8-sig')

# df_FERT_ONLYdata = pd.read_excel('MMreport_HALB_onlydata_2ND.xlsx', engine='openpyxl')
# df_FERT_SUMIF = df_FERT_ONLYdata.groupby('Unnamed: 2')[['Unnamed: 5','Unnamed: 6','Unnamed: 7','Unnamed: 8','Unnamed: 9','Unnamed: 10','Unnamed: 11','Unnamed: 12','Unnamed: 13','Unnamed: 14','Unnamed: 15','Unnamed: 16','Unnamed: 17', 'Unnamed: 18','Unnamed: 19','Unnamed: 20','Unnamed: 21','Unnamed: 22','Unnamed: 23','Unnamed: 24','Unnamed: 25','Unnamed: 26','Unnamed: 27','Unnamed: 28']].sum()
# df_FERT_SUMIF.to_excel('MMreport_HALB_SUMIF_2ND.xlsx',encoding='utf-8-sig')


