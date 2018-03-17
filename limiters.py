# -*- coding: utf-8 -*-
"""
Created on 31.12.2017

@author: ayaakobi
"""



#'Import pandas'
import pandas as pd
import datetime
import numpy as np

pd.set_option('display.expand_frame_repr', False)
pd.set_option('max_rows', 100)
pd.set_option('max_colwidth', 100)
pd.set_option('display.max_columns', 20)
pd.set_option('display.max_rows', 1000)

# Assign spreadsheet filename to `file`
limiter_file = "C:/Users/ayaakobi/alon/WW50'17 1274 Limiters POR - REV9.57 - P3e 23 - 4K Aligment.xlsm"
limiters_inc_usc_config = 'C:/Users/ayaakobi/alon/limiterUpload.xlsx'

# Load spreadsheet
xl_limtier = pd.ExcelFile(limiter_file)
# Load sheets
df_master_sheet = xl_limtier.parse("Master Sheet")
df_config_sheet = xl_limtier.parse("CEID Config")


# P print the sheet names
# print(xl_limtier.sheet_names)

# filter the sheet as inv and save it in uploas file
df_filtered_inv = df_master_sheet[df_master_sheet['CEID'].notnull() & (df_master_sheet['Metric'] == "INV") & (df_master_sheet['Detail'] == "MRCL")]
writer_to_upLoad_file = pd.ExcelWriter(limiters_inc_usc_config)
df_filtered_inv.loc[df_filtered_inv['Group'].isnull(), ['Group']] = df_filtered_inv['CEID']
df_filtered_inv['RunKey']=df_filtered_inv["CEID"].map(str) + '-'+df_filtered_inv["Group"]
# unpivot 
df_filtered_inv=pd.melt(df_filtered_inv, id_vars=['RunKey', 'Index', 'Order', 'CEID', 'Group', 'from(ceid)','Metric', 'Detail'], var_name='ww', value_name='inv')
# write to excel as format 
df_filtered_inv.to_excel(writer_to_upLoad_file, 'inv')


# filter the sheet as usc and save it in uploas file
df_filtered_usc = df_master_sheet[df_master_sheet['CEID'].notnull() & (df_master_sheet['Metric'] == "USC") & (df_master_sheet['Detail'] != "Agg") 
                                    & (df_master_sheet['RunKey'].notnull())]
# when dont have value in group put rhe ceid value
df_filtered_usc.loc[df_filtered_usc['Group'].isnull(), ['Group']] = df_filtered_usc['CEID']
# spilit the process and dot 
df_filtered_usc.rename(columns = {'Detail':'Process'}, inplace = True)
df_filtered_usc['Process'] = df_filtered_usc['Process'].astype(str)
df_filtered_usc.insert(8, 'dot', df_filtered_usc['Process'].str[5:], allow_duplicates=False)
df_filtered_usc['Process'] = df_filtered_usc['Process'].str[:4]

df_filtered_usc["RunKey"] = df_filtered_usc["CEID"].map(str) + '-'+ df_filtered_usc["Group"]


df_filtered_usc=pd.melt(df_filtered_usc, id_vars=['RunKey','Index', 'Order','CEID','Group','from(ceid)' ,'Metric','Process','dot'], var_name='ww' )


df_filtered_usc.to_excel(writer_to_upLoad_file,'usc')

# filter the sheet as config and save it in uploas file

df_filtered_config = df_config_sheet[['Area','Tool', 'Group', 'Sharing Type', 'Weight (from)' , 'with CEID' , 'Areq Factor (group 4)' , 'Tool configuration'
                                      , 'FE' , 'BE' , 'FIN', 'WELL-POLY' , 'ESD', 'ILD-PCT' , 'FTI', 'GATE' , 'MGD', 'CON1' , 'CON2', 'SAFFI_W1' , 'SAFFI_W2', 'SSAFI V0-M1' 
                                      , 'BE-LL', 'BE-UL' , 'First Occurance' ]]

df_filtered_config.rename(columns={'Tool': 'CEID'}, inplace=True)

df_filtered_config["RunKey"]=df_filtered_config["CEID"].map(str)+'-'+df_filtered_config["Group"]

df_filtered_config.to_excel(writer_to_upLoad_file,'config')



# VIEW TABLES IN RGIHT FORMAT AND NAME  
Group_USC=df_filtered_usc
Group_USC=pd.merge(Group_USC, df_filtered_config,  how='left', on=['RunKey'], left_index=True, right_index=True)

Group_USC=(Group_USC[['RunKey', 'CEID_x', 'Group_x' ,'Weight (from)', 'Process' ,'dot', 'ww', 'with CEID' ,'value']])
Group_USC.to_excel(writer_to_upLoad_file, 'Group_USC')





Group_INV=(df_filtered_inv[['RunKey', 'CEID' ,'Group' ,'Detail' ,'ww' ,'inv']])
Group_INV.to_excel(writer_to_upLoad_file, 'Group_INV')



def check_fe_or_be(row):
    if row['FE'] == 1:
        return 'FE'
    elif row['BE'] == 1:
        return 'BE' 
    else:
        return 'not_good'

df_filtered_config['FE/BE'] = df_filtered_config.apply(check_fe_or_be, axis=1)


def get_col_name_to_fist(row):      
    if row["FIN"] == 1:
        return "FIN"
    elif row['WELL-POLY'] == 1:
        return 'WELL-POLY' 
    elif row['ESD'] == 1:
        return 'ESD'
    elif row['ILD-PCT'] == 1:
        return 'ILD-PCT' 
    elif row['FTI'] == 1:
       return 'FTI'
    elif row['GATE'] == 1:
       return 'GATE'
    elif row['MGD'] == 1:
       return 'MGD'
    elif row['CON1'] == 1:
       return 'CON1'
    elif row['CON2'] == 1:
       return 'CON2'
    elif row['SAFFI_W1'] == 1:
       return 'SAFFI_W1'
    elif row['SAFFI_W2'] == 1:
       return 'SAFFI_W2'
    elif row['SSAFI V0-M1'] == 1:
       return 'SSAFI V0-M1'
    elif row['BE-LL'] == 1:
       return 'BE-LL'
    elif row['BE-UL'] == 1:
        return 'BE-UL'
    else:
        return 'not_good'


df_filtered_config['Fist'] = df_filtered_config.apply(get_col_name_to_fist, axis=1)



Group_config = df_filtered_config[['RunKey','Area','CEID','Group','Sharing Type','Weight (from)'
                                   ,'with CEID','Areq Factor (group 4)','Tool configuration', 'FE/BE' 
                                   , 'Fist' , 'First Occurance' ]]

Group_config.loc[Group_config['Group'].isnull(), ['Group']] = Group_config['CEID']
Group_config['RunKey'] = Group_config["CEID"].map(str) + '-'+ Group_config["Group"]
Group_config=Group_config.dropna(subset =['Area'])   
Group_config.to_excel(writer_to_upLoad_file,'Group_config')


# upload time
now_date = datetime.datetime.now()
now_date_to_excel = pd.DataFrame({"Date": str(now_date)}, index=[0])
print(now_date_to_excel)
now_date_to_excel.to_excel(writer_to_upLoad_file, 'time')






writer_to_upLoad_file.save()
writer_to_upLoad_file.close()









    
 
