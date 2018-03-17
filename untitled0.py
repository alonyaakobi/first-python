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






print (wsa_ews_table)


print(wsa_ews_table.keys())





