# -*- coding: utf-8 -*-
"""
Created on Mon Jan  1 15:15:47 2018

@author: ayaakobi
"""

# Import pandas
import pandas as pd
import numpy as np


# Assign spreadsheet filename to `file`
PATH_to_Yview_RECL = 'C:/Users/ayaakobi/alon/yview_and_recl.xlsx'
PATH_to_start_file = "C:/Users/ayaakobi/alon/start_mor_por_metro.xlsx"

PATH_to_new_file_union_RECL_Yview = 'C:/Users/ayaakobi/alon/union_RECL_Yview.xlsx'
#PATH_SAVE_file = ""

Yview_RECL_REV_Entity_config = pd.ExcelFile(PATH_to_Yview_RECL)
start_file = pd.ExcelFile(PATH_to_start_file)

# Print the sheet names
start_sheet = start_file.parse("Starts")
Yview_sheet = Yview_RECL_REV_Entity_config.parse("Yview loaders")
RECL_sheet = Yview_RECL_REV_Entity_config.parse("RECL loaders")
Entity_config_sheet = Yview_RECL_REV_Entity_config.parse("Rev_Entity_Config loaders")

RECL_sheet_in_format = RECL_table_fix_to_format(RECL_sheet)
Yview_sheet_in_format = Yview_table_fix_to_format(Yview_sheet)


RECL_sheet_in_format = RECL_sheet_in_format[['OPERATION','OPER_SHORT_DES','FIST','Group','volume_factor']]
Yview_sheet_in_format = Yview_sheet_in_format[['OPERATION','OPER_SHORT_DES','FIST','Group','volume_factor']]

frames = [RECL_sheet_in_format, Yview_sheet_in_format]
result = pd.concat(frames)

union_table=pd.merge(result, start_sheet[['Type','YV fist']], left_on='FIST',right_on='YV fist',  how='left')
union_table=union_table[union_table['Group'].notnull()]

writer = pd.ExcelWriter(PATH_to_new_file_union_RECL_Yview)
union_table.to_excel(writer,'union_RECL_Yview')
writer.save()
writer.close()



# RECL table fix to format 
def RECL_table_fix_to_format(RECL_sheet):
    RECL_sheet['OPERATION'] = RECL_sheet['OPER_DESC'].apply(lambda row: row.split('-')[0])
    RECL_sheet['OPER_SHORT_DES'] = RECL_sheet['OPER_DESC'].apply(lambda row: row.rsplit('-',1)[1])
    RECL_sheet.rename(columns = {'CEID':'Group'}, inplace = True)
    RECL_sheet[['OPERATION','OPER_SHORT_DES','FIST','Group','DATA']]
    RECL_sheet['DATA']=RECL_sheet['DATA'].str.strip()
    RECL_sheet=RECL_sheet[RECL_sheet['DATA'].notnull() & (RECL_sheet['DATA'] == "Y") ]
    RECL_sheet['volume_factor'] = 1

    return RECL_sheet


# RECL table fix to format 
def Yview_table_fix_to_format(Yview_sheet):
        Yview_sheet.rename(columns = {'CEID_UDF':'Group'}, inplace = True)
        Yview_sheet=Yview_sheet[Yview_sheet['Group'].notnull()]
        Yview_sheet['FIST'] = Yview_sheet.apply(lambda row: changFistNameSSAFI(row), axis=1)

        return Yview_sheet


# function to change the ssafi fist as layer name 
def changFistNameSSAFI(row):
    if row['FIST'] == 'SSAFI':
        return row['LAYER_NAME'].split(' ')[0]
    else :  return row['FIST']







