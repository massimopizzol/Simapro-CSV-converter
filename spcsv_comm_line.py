"""
Created on Fri Jan 13 22:00:00 2023

@author: massimo
"""

from spcsv import *
from sys import argv
import pandas as pd

script, input_name = argv
"""Specify the name of the file to be converted in .csv"""
"""This script allows converting multiple sheets of an Excel file"""

for sheet in pd.ExcelFile(input_name).sheet_names:

    spcsv_ready = pd.read_excel(input_name, sheet, index_col = False, header = None)

    to_spcsv(spcsv_ready, sheet +'.csv')
