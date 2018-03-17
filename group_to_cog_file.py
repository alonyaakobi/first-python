# -*- coding: utf-8 -*-
"""
Created on Sun Dec 31 22:00:48 2017

@author: ayaakobi
"""

# Import pandas
import pandas as pd
import numpy as np



# Assign spreadsheet filename to `file`
PATH_limiters_inv_usc_config = 'C:/Users/ayaakobi/alon/limiterUpload.xlsx'

COG_file = "C:/Users/ayaakobi/alon/A View FAME COG 3.01 (light).xlsm"

writer_to_COG_file = pd.ExcelWriter(COG_file)

# Load spreadsheet
Group = pd.ExcelFile(PATH_limiters_inv_usc_config)


# Print the sheet names
#print(Group.sheet_names)

Group_config = Group.parse("Group_config")

Group_config = Group_config['Group'].drop_duplicates()

writer = pd.ExcelWriter(PATH_limiters_inv_usc_config, engine='xlsxwriter')

Group_config.to_excel(writer, sheet_name='Distinct Groups')

workbook  = writer.book
workbook.filename = COG_file
workbook.add_vba_project('./vbaProject.bin')
# !! Won't load in Excel !!

writer.save()