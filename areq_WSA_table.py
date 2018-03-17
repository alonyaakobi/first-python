# -*- coding: utf-8 -*-
"""
Created on Sun Jan 14 11:11:17 2018

@author: ayaakobi
"""


import pandas as pd
import numpy as np
import datetime as dt 
from datetime import datetime


pd.set_option('display.expand_frame_repr', False)
pd.set_option('display.max_rows', 1000)

PATH_limiters_inc_usc_config = 'C:/Users/ayaakobi/alon/limiterUpload.xlsx'
PATH_data_for_MOR_POR_file = "C:/Users/ayaakobi/alon/start_mor_por_metro.xlsx"
PATH_WW = "C:/Users/ayaakobi/alon/WWs.xlsx"
PATH_to_WSA_file = "C:/Users/ayaakobi/alon/wsa.xlsx"   
PATH_to_Inventory_file = "C:/Users/ayaakobi/alon/Inventory.xlsx"  



config_sheet = pd.ExcelFile(PATH_limiters_inc_usc_config)
MOR_POR_sheet = pd.ExcelFile(PATH_data_for_MOR_POR_file)
ww_sheet = pd.ExcelFile(PATH_WW)
Inventory_file = pd.ExcelFile(PATH_to_Inventory_file)


# Print the sheet names
Group_config = config_sheet.parse("Group_config")
MOR_table = MOR_POR_sheet.parse("Target Availability")
POR_table = MOR_POR_sheet.parse("POR Availability")

MOR_table = MOR_table[['RunKey', 'Target Avail']]
POR_table = POR_table[['RunKey', 'Target Avail']]

Group_config = pd.merge(Group_config, MOR_table, on=['RunKey'], how='left')
Group_config.rename(columns={'Target Avail': 'MOR_Target_Avail'}, inplace=True)

Group_config = pd.merge(Group_config, POR_table, on=['RunKey'], how='left')
Group_config.rename(columns={'Target Avail': 'POR_Target_Avail'}, inplace=True)

Group_config = Group_config.drop_duplicates(subset=['RunKey'], keep=False)
Group_config = Group_config.reset_index()





def avg_group_mor(row):
    
    PATH_data_for_MOR_POR_file = "C:/Users/ayaakobi/alon/start_mor_por_metro.xlsx"
    POR_MOR_sheet = pd.ExcelFile(PATH_data_for_MOR_POR_file)
    MOR_table = POR_MOR_sheet.parse("Target Availability")
    MOR_table = MOR_table.drop_duplicates(subset=['Group'], keep=False)
    MOR_table.set_index('RunKey')
    
    result = (MOR_table.loc[MOR_table['Group'] ==  row[4] , ['Target Avail']]).mean()
    result = pd.to_numeric(result)
    result=result.reset_index()
    result = (result.loc[0, 0])
    return result
    
def avg_group_por(row):
    
    PATH_data_for_MOR_POR_file = "C:/Users/ayaakobi/alon/start_mor_por_metro.xlsx"
    POR_POR_sheet = pd.ExcelFile(PATH_data_for_MOR_POR_file)
    POR_table = POR_POR_sheet.parse("POR Availability")
    POR_table = POR_table.drop_duplicates(subset=['Group'], keep=False)
    POR_table.set_index('RunKey')    
    
    result = (POR_table.loc[POR_table['Group'] ==  row[4] , ['Target Avail']]).mean()
    result = pd.to_numeric(result)
    result=result.reset_index()
    result = (result.loc[0, 0])
    
    return result     

def return_as_intel_calnder(row):
    today_date = dt.datetime.today()
    if ( (row['WW_Start_Time'])<= today_date and (row['WW_End_Time'])> today_date ):
        return (row['WW_Format'])

  
    
    
     
Group_config['MOR_Target_Avail'] = (Group_config.apply(lambda row: avg_group_mor(row) if pd.isnull(row['MOR_Target_Avail']) else row['MOR_Target_Avail'], axis=1))

Group_config['POR_Target_Avail'] = (Group_config.apply(lambda row: avg_group_por(row) if pd.isnull(row['POR_Target_Avail']) else row['POR_Target_Avail'], axis=1))

ww_table = ww_sheet.parse("ww_table")

date_by_intel_calnder = (ww_table.apply(lambda row: return_as_intel_calnder(row) , axis=1))

date_by_intel_calnder = date_by_intel_calnder.dropna()
date_by_intel_calnder = pd.Series(date_by_intel_calnder, name="ww")
date_by_intel_calnder = pd.DataFrame(date_by_intel_calnder)
date_by_intel_calnder.set_index('ww')


Group_USC = config_sheet.parse("Group_USC") 

Group_USC.set_index('ww')

Group_USC_fillter_by_nwxt_ww = Group_USC.loc[(Group_USC['ww'] == date_by_intel_calnder['ww'].squeeze()  )]

Group_USC_fillter_by_nwxt_ww = Group_USC_fillter_by_nwxt_ww.drop_duplicates()

Group_config.set_index('RunKey')

Group_USC_fillter_by_nwxt_ww.set_index('RunKey')


wsa_table = pd.merge(Group_config, Group_USC_fillter_by_nwxt_ww, on='RunKey', how='left')

wsa_table = wsa_table[[ 'RunKey', 'Area', 'CEID', 'Group', 'Sharing Type',
       'Weight (from)_x', 'with CEID_x',
       'Tool configuration', 'FE/BE', 'Fist', 'First Occurance',
       'MOR_Target_Avail', 'POR_Target_Avail', 'CEID_x', 'Group_x',
       'Process', 'dot', 'ww', 'value']]


Inventory_sheet = Inventory_file.parse("Inventory") 
Inventory_sheet.rename(columns={'group': 'Group'}, inplace=True)
Inventory_sheet['RunKey'] = Inventory_sheet["CEID"].map(str)+'-'+Inventory_sheet["Group"]
Inventory_sheet['Is_prode'] ='Prode'

wsa_table = pd.merge(wsa_table, Inventory_sheet[[ 'RunKey', 'inv' , 'Is_prode']], on='RunKey', how='left')
wsa_table = wsa_table.loc[(wsa_table['Process'] == 1274) & (wsa_table['dot'] == 0)]
wsa_table['Is_prode'] = wsa_table['Is_prode'].fillna('Prode')


wsa_table.rename(columns={'Weight (from)_x': 'Weight'}, inplace=True)
wsa_table.rename(columns={'with CEID_x': 'From CEID'}, inplace=True)
wsa_table.rename(columns={'CEID_x': 'CEID'}, inplace=True)
wsa_table.rename(columns={'Group_x': 'Group'}, inplace=True)
wsa_table.rename(columns={'value': 'USC'}, inplace=True)
wsa_table['inv'] = wsa_table['inv'].fillna(0)


wsa_table['WSA'] =  wsa_table['inv'] * wsa_table['USC']



#metro tools 

metro_sheet = pd.ExcelFile(PATH_data_for_MOR_POR_file)
metro_table = metro_sheet.parse("Metro")

metro_table = metro_table[['RunKey', 'team', 'CEID' ,'Group', 'Type' , 'Prod_metro']]
metro_table['Sharing Type'] = 'Regular Sum'
metro_table['Weight'] = 1
metro_table['From CEID'] = metro_table['CEID']
metro_table['First Occurance'] = 1
metro_table['Fist'] = 'Metro'
metro_table['USC'] = 0
metro_table['inv'] = 1
metro_table.rename(columns={'team': 'Area'}, inplace=True)
metro_table.rename(columns={'Prod_metro': 'Is_prode'}, inplace=True)

metro_table = pd.merge(metro_table, MOR_table, on=['RunKey'], how='left')
metro_table.rename(columns={'Target Avail': 'MOR_Target_Avail'}, inplace=True)

metro_table = pd.merge(metro_table, POR_table, on=['RunKey'], how='left')
metro_table.rename(columns={'Target Avail': 'POR_Target_Avail'}, inplace=True)


starts_table = metro_sheet.parse("Starts")
starts_table = starts_table[['ww starts']]
starts = starts_table.max() + starts_table.max()*0.15

metro_table['WSA'] = starts[0]


wsa_table = wsa_table.T.drop_duplicates().T

wsa_table = pd.concat([wsa_table, metro_table]).drop_duplicates()


wsa_table = wsa_table[[    'RunKey', 'Area', 'CEID' , 'Group', 'Sharing Type' , 
                           'Weight' , 'From CEID' , 'First Occurance' , 'Fist' , 'Type' ,'MOR_Target_Avail',
                           'POR_Target_Avail','USC', 'Is_prode' , 'FE/BE',
                            'Process', 'Tool configuration', 'dot', 'inv',
                           'ww' ,  'WSA']]


writer = pd.ExcelWriter(PATH_to_WSA_file)
wsa_table.to_excel(writer,'WSA')
writer.save()
writer.close()














