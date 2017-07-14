#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Tue Jul 11 06:56:43 2017

@author: massimo
"""

import pandas as pd
import os
import numpy as np


def to_spcsv(dataframe, name):

    """
    Super inefficient function 
    adapted from previous script to convert excel files into csv
    Converts a dataframe in simapro csv format
    'dataframe' [pd.DataFrame] object in same structure as the Excel files
    'name' [string] name of the output file, e.g. 'myLCI.csv'
    """
    output = open(name, 'w')

    # Open destination file and print the standard heading
    output.write('{CSV separator: Semicolon}\n')
    output.write('{CSV Format version: 7.0.0}\n')
    output.write('{Decimal separator: .}\n')
    output.write('{Date separator: /}\n')
    output.write('{Short date format: dd/MM/yyyy}\n\n')

    # List of fields required
    fields = ["Process", "Category type", "Time Period", "Geography",
              "Technology", "Representativeness", "Multiple output allocation",
              "Substitution allocation", "Cut off rules", "Capital goods",
              "Boundary with nature", "Record", "Generator", "Literature references",
              "Collection method", "Data treatment", "Verification",
              "Products", "Materials/fuels", "Resources", "Emissions to air",
              "Emissions to water", "Emissions to soil", "Final waste flows",
              "Non material emission", "Social issues", "Economic issues",
              "Waste to treatment", "End"
              ]

    # Standard value of these fields
    fields_value = ['', '', "Unspecified", "Unspecified", "Unspecified",
                    "Unspecified", "Unspecified", "Unspecified", "Unspecified",
                    "Unspecified", "Unspecified", '', '', '', '', '',
                    "Comment", '', '', '', '', '', '', '', '', '', '', '', ''
                    ]

    # Identify the processes
    LCI = dataframe.copy()
    LCI = LCI.replace(0.0, np.nan)
    LCI = LCI.replace(np.nan, '')
    processes = LCI.iloc[0,4:]

    # Screen through the processes
    for i in range(0, len(processes)):
        process = LCI.iloc[:,(i+4)]
        fields_value[1] = process[5]
        products = (6 * "\"%s\";") % (str(process[0]), str(process[1]),
                    str(process[2]), "100%", "not defined", str(process[4]))
        fields_value[17] = products
        matfuel_list = []
        raw_list = []
        air_list = []
        water_list = []
        soil_list = []
        finalwaste_list = []
        social_list = []
        economic_list = []
        wastetotreatment_list = []

        # Screen through the inputs and outputs of each process
        for j in range(6, len(process)):

            if LCI.iloc[j, 0] == '' and LCI.iloc[j,i+4] != '':
                matfuel = (7 * "\"%s\";") % (LCI.iloc[j,1], LCI.iloc[j,3], LCI.iloc[j,i+4], "Undefined", 0, 0, 0)
                matfuel_list.append(matfuel)

            elif LCI.iloc[j, 0] == ("Raw") and LCI.iloc[j,i+4] != '':
                raw = (8 * "\"%s\";") % (LCI.iloc[j,1], '', LCI.iloc[j,3], LCI.iloc[j,i+4], "Undefined", 0, 0, 0)
                raw_list.append(raw)

            elif LCI.iloc[j, 0] == ("Air") and LCI.iloc[j,i+4] != '':
                air = (8 * "\"%s\";") % (LCI.iloc[j,1], '', LCI.iloc[j,3], LCI.iloc[j,i+4], "Undefined", 0, 0, 0)
                air_list.append(air)

            elif LCI.iloc[j, 0] == ("Water") and LCI.iloc[j,i+4] != '':
                water = (8 * "\"%s\";") % (LCI.iloc[j,1], '', LCI.iloc[j,3], LCI.iloc[j,i+4], "Undefined", 0, 0, 0)
                water_list.append(water)

            elif LCI.iloc[j, 0] == ("Soil") and LCI.iloc[j,i+4] != '':
                soil = (8 * "\"%s\";") % (LCI.iloc[j,1], '', LCI.iloc[j,3], LCI.iloc[j,i+4], "Undefined", 0, 0, 0)
                soil_list.append(soil)

            elif LCI.iloc[j, 0] == ("Waste") and LCI.iloc[j,i+4] != '':
                finalwaste = (8 * "\"%s\";") % (LCI.iloc[j,1], '', LCI.iloc[j,3], LCI.iloc[j,i+4], "Undefined", 0, 0, 0)
                finalwaste_list.append(finalwaste)

            elif LCI.iloc[j, 0] == ("Social") and LCI.iloc[j,i+4] != '':
                social = (8 * "\"%s\";") % (LCI.iloc[j,1], '', LCI.iloc[j,3], LCI.iloc[j,i+4], "Undefined", 0, 0, 0)
                social_list.append(social)

            elif LCI.iloc[j, 0] == ("Economic") and LCI.iloc[j,i+4] != '':
                economic = (8 * "\"%s\";") % (LCI.iloc[j,1], '', LCI.iloc[j,3], LCI.iloc[j,i+4], "Undefined", 0, 0, 0)
                economic_list.append(economic)

            elif LCI.iloc[j, 0] == "Wastetotreatment" and LCI.iloc[j,i+4] != '':
                wastetotreatment = (7 * "\"%s\";") % (LCI.iloc[j,1], LCI.iloc[j,3], LCI.iloc[j,i+4], "Undefined", 0, 0, 0)
                wastetotreatment_list.append(wastetotreatment)

            # Assign the inputs and outputs to a list
            fields_value[18] = matfuel_list
            fields_value[19] = raw_list
            fields_value[20] = air_list
            fields_value[21] = water_list
            fields_value[22] = soil_list
            fields_value[23] = finalwaste_list
            fields_value[25] = social_list
            fields_value[26] = economic_list
            fields_value[27] = wastetotreatment_list
    
        for el in range (0, len(fields)):  # Important, note the indentation here
            output.write("%s\n" % fields[el])
            
            if not isinstance(fields_value[el], list):
                output.write("%s\n" % fields_value[el])	
            else:
                for j in fields_value[el]:
                    variable = "%s" % j
                    output.write("%s\n" % variable)
            output.write("\n")
    output.close()