# -*- coding: utf-8 -*-
"""
Created on Tue Dec 26 16:46:37 2017

@author: ayaakobi
"""


# Import pandas
import pandas as pd
from openpyxl import load_workbook
import datetime
import numpy as np


pd.set_option('max_rows', 100)
pd.set_option('max_colwidth',100)
pd.set_option('display.max_columns', 15)
pd.set_option('display.max_rows', 10)

# Assign spreadsheet filename to `file`
limiter_file = "C:/Users/ayaakobi/alon/WW50'17 1274 Limiters POR - REV9.57 - P3e 23 - 4K Aligment.xlsm"
limiters_inc_usc_config='C:/Users/ayaakobi/alon/limiterUpload.xlsx'

# Load spreadsheet
#xl_limtier = pd.ExcelFile(limiter_file)
# Load sheets
df_master_sheet = xl_limtier.parse("Master Sheet")
df_config_sheet = xl_limtier.parse("CEID Config")



#P print the sheet names
#print(xl_limtier.sheet_names)

#filter the sheet as inv and save it in uploas file
df_filtered_inv = df_master_sheet[df_master_sheet['CEID'].notnull() & (df_master_sheet['Metric'] == "INV") & (df_master_sheet['Detail'] == "MRCL")]
writer_to_upLoad_file = pd.ExcelWriter(limiters_inc_usc_config)
df_filtered_inv.to_excel(writer_to_upLoad_file,'inv')

#filter the sheet as usc and save it in uploas file
df_filtered_usc = df_master_sheet[df_master_sheet['CEID'].notnull() & (df_master_sheet['Metric'] == "USC") & (df_master_sheet['Detail'] != "Agg") 
                                    & (df_master_sheet['RunKey'].notnull())]
# when dont have value in group put rhe ceid value
df_filtered_usc.loc[df_filtered_usc['Group'].isnull(), ['Group']] = df_filtered_usc['CEID']
#spilit the process and dot 
df_filtered_usc.rename(columns = {'Detail':'Process'}, inplace = True)
df_filtered_usc['Process'] = df_filtered_usc['Process'].astype(str)
df_filtered_usc.insert(8, 'dot', df_filtered_usc['Process'].str[5:], allow_duplicates=False)
df_filtered_usc['Process'] = df_filtered_usc['Process'].str[:4]

df_filtered_usc["RunKey"] = df_filtered_usc["CEID"].map(str) + '-'+ df_filtered_usc["Group"]


print(df_filtered_usc)

df_filtered_usc=df_filtered_usc.melt(df_filtered_usc, id_vars=['farm', 'fruit'], var_name='WW', value_name='value')




df_filtered_usc.to_excel(writer_to_upLoad_file,'usc')

#filter the sheet as config and save it in uploas file



df_filtered_config = df_config_sheet[['Area','Tool', 'Group', 'Sharing Type', 'Weight (from)' , 'with CEID' , 'Areq Factor (group 4)' , 'Tool configuration'
                                      , 'FE' , 'BE' , 'FIN', 'WELL-POLY' , 'ESD', 'ILD-PCT' , 'FTI', 'GATE' , 'MGD', 'CON1' , 'CON2', 'SAFFI_W1' , 'SAFFI_W2', 'SSAFI V0-M1' 
                                      , 'BE-LL', 'BE-UL' , 'First Occurance' ]]

df_filtered_config.rename(columns = {'Tool':'CEID'}, inplace = True)
df_filtered_config["RunKey"] = df_filtered_config["CEID"].map(str) + '-'+ df_filtered_config["Group"]





df_view_filtered_usc = pd.merge(df_filtered_usc, df_filtered_config , how='left', on=['CEID'], left_index=True, right_index=True)
print(df_view_filtered_usc)

df_view_filtered_usc.to_excel(writer_to_upLoad_file,'try1')

df_filtered_config = df_filtered_config.set_index('RunKey')
df_filtered_usc = df_filtered_usc.set_index('RunKey')



df_filtered_config.to_excel(writer_to_upLoad_file,'config')






df_view_filtered_usc = pd.merge(df_filtered_usc, df_filtered_config , how='left', on=['CEID'], left_index=True, right_index=True)


print(df_view_filtered_usc)


df_view_filtered_usc.drop_duplicates()

df_view_filtered_usc.to_excel(writer_to_upLoad_file,'try2')








# upload time
now_date = datetime.datetime.now()
now_date_to_excel = pd.DataFrame({"Date": str(now_date)}, index=[0])
print(now_date_to_excel)
now_date_to_excel.to_excel(writer_to_upLoad_file, 'time')






writer_to_upLoad_file.save()
writer_to_upLoad_file.close()



