
"""
Created on Fri Jan 13 21:57:00 2023

@author: massimo
"""

from spcsv import *
import pandas as pd

for sheet in pd.ExcelFile('LCI_Example.xlsx').sheet_names:

    spcsv_ready = pd.read_excel('LCI_Example.xlsx', sheet, index_col = False, header = None)

    to_spcsv(spcsv_ready, sheet +'.csv')
