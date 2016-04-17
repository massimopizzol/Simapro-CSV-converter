import xlrd #to open excel files
import os
from sys import argv 


script, input_name = argv 
"""Specify the name of the file to be converted in .csv"""
"""This script allows converting multiple sheets of an Excel file"""

#open workbook and find data
Data = xlrd.open_workbook(input_name)

#scroll across sheets

for sheet in range(0,Data.nsheets):
	LCI = Data.sheet_by_index(sheet)

	output_name = LCI.name + ".csv"
	output = open(output_name, 'w') 

#Open destination file and print the standard heading
	output.write("""{CSV separator: Semicolon}
{CSV Format version: 7.0.0}
{Decimal separator: .}
{Date separator: /}
{Short date format: dd/MM/yyyy}

""")

#List of fields required
	fields = ["Process", "Category type", "Time Period", 
		"Geography", "Technology", "Representativeness", 
		"Multiple output allocation", "Substitution allocation", "Cut off rules", 
		"Capital goods", "Boundary with nature", "Record", 
		"Generator", "Literature references", "Collection method", 
		"Data treatment", "Verification", "Products", 
		"Materials/fuels", "Resources", "Emissions to air", 
		"Emissions to water", "Emissions to soil", "Final waste flows", 
		"Non material emission", "Social issues", "Economic issues", 
		"Waste to treatment", "End"
		]

#Standard value of these fields
	fields_value = ['', '', "Unspecified", 
		"Unspecified", "Unspecified", "Unspecified", 
		"Unspecified", "Unspecified", "Unspecified", 
		"Unspecified", "Unspecified", '', 
		'', '', '', 
		'', "Comment", '', 
		'', '', '', 
		'', '', '', 
		'', '', '', 
		'', ''
		]

#Identify the processes
	processes = LCI.row_slice(rowx=0, start_colx=4, end_colx=LCI.ncols) 

#Screen through the processes
	for i in range(0, len(processes)):
		process = LCI.col_slice(colx=(i+4), start_rowx=0, end_rowx=LCI.nrows)
	
		fields_value[1] = process[5].value
	#print fields_value[1]  # Use this as a double check
	
		products = (6 * "\"%s\";") % (
			str(process[0].value), str(process[1].value), 
			str(process[2].value), "100%", "not defined", str(process[4].value)) 
		fields_value[17] = products
	#print fields_value[17] 
	#print fields_value 
	
		matfuel_list = []
		raw_list = []
		air_list = []
		water_list = []
		soil_list = []
		finalwaste_list = []
		social_list = []
		economic_list = []
		wastetotreatment_list = []
	
	#Screen through the inputs and outputs of each process
		for j in range(6, len(process)):
	
			if LCI.cell(j, 0).value == '' and LCI.cell(j, i+4).value != '':
				matfuel = (7 * "\"%s\";") % (
					LCI.cell(j,1).value, LCI.cell(j,3).value, 
					LCI.cell(j,i+4).value, "Undefined", 0, 0, 0
					)
				matfuel_list.append(matfuel)
			
			elif LCI.cell(j, 0).value == ("Raw") and LCI.cell(j, i+4).value != '':
				raw = (8 * "\"%s\";") % (
					LCI.cell(j,1).value, '', LCI.cell(j,3).value, 
					LCI.cell(j,i+4).value, "Undefined", 0, 0, 0
					)
				raw_list.append(raw)
		
			elif LCI.cell(j, 0).value == ("Air") and LCI.cell(j, i+4).value != '':
				air = (8 * "\"%s\";") % (
					LCI.cell(j,1).value, '', LCI.cell(j,3).value, 
					LCI.cell(j,i+4).value, "Undefined", 0, 0, 0
					)
				air_list.append(air)
		
			elif LCI.cell(j, 0).value == ("Water") and LCI.cell(j, i+4).value != '':
				water = (8 * "\"%s\";") % (
					LCI.cell(j,1).value, '', LCI.cell(j,3).value, 
					LCI.cell(j,i+4).value, "Undefined", 0, 0, 0
					)
				water_list.append(water)
		
			elif LCI.cell(j, 0).value == ("Soil") and LCI.cell(j, i+4).value != '':
				soil = (8 * "\"%s\";") % (
					LCI.cell(j,1).value, '', LCI.cell(j,3).value, 
					LCI.cell(j,i+4).value, "Undefined", 0, 0, 0
					)
				soil_list.append(soil)
			
			elif LCI.cell(j, 0).value == ("Waste") and LCI.cell(j, i+4).value != '':
				finalwaste = (8 * "\"%s\";") % (
					LCI.cell(j,1).value, '', LCI.cell(j,3).value, 
					LCI.cell(j,i+4).value, "Undefined", 0, 0, 0
					)
				finalwaste_list.append(finalwaste)
			
			elif LCI.cell(j, 0).value == ("Social") and LCI.cell(j, i+4).value != '':
				social = (8 * "\"%s\";") % (
					LCI.cell(j,1).value, '', LCI.cell(j,3).value, 
					LCI.cell(j,i+4).value, "Undefined", 0, 0, 0
					)
				social_list.append(social)
			
			elif LCI.cell(j, 0).value == ("Economic") and LCI.cell(j, i+4).value != '':
				economic = (8 * "\"%s\";") % (
					LCI.cell(j,1).value, '', LCI.cell(j,3).value, 
					LCI.cell(j,i+4).value, "Undefined", 0, 0, 0
					)
				economic_list.append(economic)
				
			elif LCI.cell(j, 0).value == "Wastetotreatment" and LCI.cell(j, i+4).value != '':
				wastetotreatment = (7 * "\"%s\";") % (
					LCI.cell(j,1).value, LCI.cell(j,3).value, 
					LCI.cell(j,i+4).value, "Undefined", 0, 0, 0
					)
				wastetotreatment_list.append(wastetotreatment)
						
			
			#Assign the inputs and outputs to a list
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