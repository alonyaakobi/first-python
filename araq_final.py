# -*- coding: utf-8 -*-
"""
Created on Mon Feb 26 09:14:07 2018

@author: ayaakobi


final areq :) 

"""

import pandas as pd
import numpy as np
import datetime as dt 
from datetime import datetime

pd.set_option('display.expand_frame_repr', False)
pd.set_option('display.max_rows', 1000)

PATH_to_WSA_file = "C:/Users/ayaakobi/alon/wsa.xlsx"   
PATH_to_EWS_with_factor = "C:/Users/ayaakobi/alon/Areqs_EWS_with_factor.xlsx"  
PATH_to_Areq_Final = "C:/Users/ayaakobi/alon/Areq_Final.xlsx" 

WSA_sheet = pd.ExcelFile(PATH_to_WSA_file)
EWS_with_factor_sheet = pd.ExcelFile(PATH_to_EWS_with_factor)




WSA_table = WSA_sheet.parse("WSA")
EWS_with_factor_table = EWS_with_factor_sheet.parse("Areqs_EWS_with_factor")

wsa_ews_table = pd.merge(WSA_table, EWS_with_factor_table[[ 'Group', 'OPERATION' , 'EWS_with_factor']], on='Group', how='left')

wsa_ews_table = wsa_ews_table.drop_duplicates(subset=['Group'])



def buildTheUSC (x,y):
    num1 = x*y 
    if (x==0):
        return num1/1
    else: 
        return num1/x




wsa_ews_table = wsa_ews_table.groupby(['Group', 'Area' , 'Is_prode'] ).agg({'POR_Target_Avail': np.max, 'MOR_Target_Avail': np.max 
                                         ,'Type': np.max , 'CEID' : np.size  , 'inv' : np.sum , 'USC': np.sum 
                                         ,'WSA': np.sum , 'EWS_with_factor':np.sum })



wsa_ews_table['USC'] = np.vectorize(buildTheUSC)(wsa_ews_table['inv'], wsa_ews_table['USC'])

wsa_ews_table['EWS1'] = np.vectorize(buildTheUSC)(wsa_ews_table['inv'], wsa_ews_table['EWS_with_factor'])

wsa_ews_table = wsa_ews_table.reset_index()

#--------------------------------------------------------

def calcsTheAreq(row):
    
    if (row['EWS1'] == 0) | (row['WSA'] == 0):
        return 0
    elif pd.isnull(row['EWS1']) == True & (row['Is_prode'] == 'Prode'):
        if row['Area']=='CMS' or row['Area']=='TM' or row['Area']=='LI3' or row['Area']=='DM':
            return row['POR_Target_Avail']
        else: return 0.5
    elif row['USC']==0 and (row['Is_prode'] == 'Prode'):
        if row['Area']=='CMS' or row['Area']=='TM' or row['Area']=='LI3' or row['Area']=='DM':
            return row['POR_Target_Avail']
        else: return 0.5 
    elif (pd.isnull(row['EWS1']) == True or row['USC']==0 ) and row['Is_prode'] == 'Metro':
        return row['POR_Target_Avail']
    else : return (row['EWS1']/row['WSA'] == 0)*row['POR_Target_Avail']



wsa_ews_table['Areq_calc'] = wsa_ews_table.apply(lambda row: calcsTheAreq(row), axis=1)


def areqFinal(row):
    
    if pd.isnull(row['EWS1']) == True and (row['Is_prode'] == 'Metro'):
        return row['POR_Target_Avail']
    elif (row['Areq_calc'] > row['POR_Target_Avail']) & (row['POR_Target_Avail']<=0.8):
        if (row['Areq_calc']-row['POR_Target_Avail'])> 0.1:
            return 0.1
        else: return row['Areq_calc']-row['POR_Target_Avail']
    
    elif ((row['Areq_calc']) > row['POR_Target_Avail']) & (row['POR_Target_Avail']>0.8):
        if (row['Areq_calc']-row['POR_Target_Avail'])>0.5:
            num1=0.05
        else: num1=row['Areq_calc']-row['POR_Target_Avail']
        if row['POR_Target_Avail']+num1>row['MOR_Target_Avail']:
            return row['MOR_Target_Avail']
        else :
            if (row['Areq_calc']-row['POR_Target_Avail'])>(row['POR_Target_Avail']+0.05):
                return 0.05
            else: return (row['POR_Target_Avail']+row['Areq_calc']-row['POR_Target_Avail'])
    elif ( (row['Areq_calc'] <= row['POR_Target_Avail']) & (row['Areq_calc']>0.5)  ):
        return row['Areq_calc']
    elif (row['Areq_calc']<=0.5):
        return 0.5
    else: return 0 
            
            
            
            

wsa_ews_table['Areq_Final'] = wsa_ews_table.apply(lambda row: areqFinal(row), axis=1)


writer = pd.ExcelWriter(PATH_to_Areq_Final)
wsa_ews_table.to_excel(writer,'Areq_Final')
writer.save()
writer.close()









