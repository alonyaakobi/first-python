# -*- coding: utf-8 -*-
"""
Created on Tue Dec 26 15:20:40 2017

@author: ayaakobi
"""


def 

# Import pandas
import pandas as pd
from openpyxl import load_workbook
import numpy as np



# Assign spreadsheet filename to `file`
data_for_all_file = "C:/Users/ayaakobi/alon/start_mor_por_metro.xlsx"



# Load spreadsheet
limiters_mor_por_start_metro = pd.ExcelFile(data_for_all_file)
# Load sheets

# Print the sheet names


mor_table = limiters_mor_por_start_metro.parse("Target Availability")

por_table = limiters_mor_por_start_metro.parse("POR Availability")

starts_table = limiters_mor_por_start_metro.parse("Starts")

print(type(starts_table))
print(starts_table)
print(starts_table.groupby('Type').agg({'ww starts': 'sum'}))

