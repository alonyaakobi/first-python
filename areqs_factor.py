# -*- coding: utf-8 -*-
"""
Created on Sun Jan  7 14:49:37 2018

@author: ayaakobi
"""

import pandas as pd
import numpy as np


pd.set_option('display.expand_frame_repr', False)


# Assign spreadsheet filename to `file`
PATH_to_union_Yview_RECL = 'C:/Users/ayaakobi/alon/union_RECL_Yview.xlsx'
PATH_to_new_file_union_RECL_Yview = 'C:/Users/ayaakobi/alon/Areqs_EWS_with_factor.xlsx'


Yview_RECL_union = pd.ExcelFile(PATH_to_union_Yview_RECL)


# Print the sheet names
Yview_RECL_union = Yview_RECL_union.parse("union_RECL_Yview")

factor_value = create_factor_value(Yview_RECL_union)

num_op_in_fist = function_num_op_in_fist(Yview_RECL_union)

volume_factor_group = Yview_RECL_union.groupby('Group')['volume_factor'].mean() #its can be change in the future

result_agg_all_table_to_EWS = function_calc_EWS_no_factor(num_op_in_fist,volume_factor_group ,factor_value)

result_agg_all_table_to_EWS=result_agg_all_table_to_EWS[['sum_Start_In_Fist','OPERATION','EWS_no_factor','EWS_Factor','volume_factor','EWS_with_factor']]


writer = pd.ExcelWriter(PATH_to_new_file_union_RECL_Yview)
result_agg_all_table_to_EWS.to_excel(writer,'Areqs_EWS_with_factor')
writer.save()
writer.close()




def create_factor_value(Yview_RECL_union):
        count_all = Yview_RECL_union.groupby('Group')['volume_factor'].sum()
        count_all.rename(columns={'volume_factor': 'volume_factor_count_all'}, inplace=True)
        only_yview=Yview_RECL_union[(Yview_RECL_union['Type'] == "Yview") ]
        only_yview = only_yview.groupby('Group')['volume_factor'].sum()
        
        result = pd.concat([count_all, only_yview], axis=1)
        result = result.applymap(lambda x:1 if pd.isnull(x) else x)
        result.columns.values[0] = 'Sum_all'
        result.columns.values[1] = 'Sum_only_yview'
        result['EWS_Factor'] = result.apply(lambda row: row['Sum_all']/row['Sum_only_yview'], axis=1)
        
        return result
    
    
def     function_num_op_in_fist(Yview_RECL_union):

        count_number_opr_in_FIST = Yview_RECL_union.groupby(['Group','FIST'])['OPERATION'].count()
        count_number_opr_in_FIST=pd.DataFrame(count_number_opr_in_FIST)
        count_number_opr_in_FIST=count_number_opr_in_FIST.reset_index()        
        data_start_file = "C:/Users/ayaakobi/alon/start_mor_por_metro.xlsx"
        data_start = pd.ExcelFile(data_start_file)
        starts_table = data_start.parse("Starts")
        starts_table.rename(columns={'YV fist': 'FIST'}, inplace=True)
        
        result = pd.merge(count_number_opr_in_FIST, starts_table, how='left', on=['FIST'])
        result=result[['Group','FIST','OPERATION','ww starts','Type','volume_factor']]
        result['sum_Start_In_Fist'] = result.apply(lambda row: row['OPERATION'] * row['ww starts'], axis=1)
               
        return result    
    

def     function_calc_EWS_no_factor(num_op_in_fist , volume_factor_group , factor_value):

        sum_opr_and_start_in_group = num_op_in_fist.groupby(['Group'])['OPERATION','sum_Start_In_Fist'].sum()
        sum_opr_and_start_in_group['EWS_no_factor']= sum_opr_and_start_in_group.apply(lambda row: row['sum_Start_In_Fist'] / row['OPERATION'], axis=1)
        result = pd.concat([sum_opr_and_start_in_group, volume_factor_group], axis=1)      
        result = pd.concat([result, factor_value], axis=1)
        result['EWS_with_factor'] = result.apply(lambda row: row['EWS_no_factor']*row['EWS_Factor']*row['volume_factor'], axis=1)
           
        return  result
    
    
    
    
    
    
    
    
    
    
    
    