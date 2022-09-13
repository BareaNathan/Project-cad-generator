import xlrd
import csv

#Function to return the sheet object
def get_spreadsheet(name):
	book = xlrd.open_workbook(name)
	sh = book.sheet_by_index(0)
	return sh

#Function to get the columns of references and parnumber from XLS
def get_columns(sheet):
	col_b = sheet.col(1)
	col_b.insert(0," ")
	str_col_b = str(col_b)
	list_col_b = str_col_b.split(', ')

	col_e = sheet.col(4)
	str_col_e = str(sheet.col(4))
	list_col_e = str_col_e.split(',')

	return list_col_e, list_col_b

#Function to treat the list, getting just the important information
def clean_lists(ref_list, part_list):
	raw_part = []
	raw_ref = []

	#Removing the empty strings
	for x in ref_list:
		if 'text' in x:
			index_raw_ref = ref_list.index(x)
			raw_part.append(part_list[index_raw_ref].replace("text:'","").replace("'","")[0:6])
			raw_ref.append(x.replace(" text:'","").replace(" '",""))

	part = []
	ref = []

	#Spliting the group of references
	for x in raw_ref:
		index_ref = raw_ref.index(x)
		y = x.split('-')
		for a in y:
			part.append(raw_part[index_ref])
			ref.append(a)
	#Inserting the header
	ref.insert(0,"Designator")
	part.insert(0,"Partnumber")
	return ref, part

#Take the CSV CAD data
def get_csv_data(name, template_model):
	csv_data = []
	with open(name) as file:
		csv_sheet = csv.DictReader(file,fieldnames=('A','B','C','D','E','F','G','H','I','J'))
		if template_model == 'Altium':
			for row in csv_sheet:
				csv_data.append([row['A'],row['E'],row['F'],row['G']])
			del csv_data[:9]
		elif template_model == 'Protel':
			for row in csv_sheet:
				csv_data.append([row['A'],row['C'],row['D'],row['J']])
			del csv_data[1]

	return csv_data

#Unificate it all
def join_csv_and_xls(csv_data,xls_ref_data,xls_part_data):
	for row in csv_data:
		if row[0] in xls_ref_data:
			xls_ref_index = xls_ref_data.index(row[0])
			row.append(xls_part_data[xls_ref_index])
		else:
			row.append('Not found in XLS file')
	return csv_data

#Write a new CSV file with the output data
def output_csv(filename, list_of_contents):
	with open(filename, 'w', newline='') as new_csv:
		write = csv.writer(new_csv, quoting=csv.QUOTE_ALL)
		for row in list_of_contents:
			write.writerow(row)
	return print("File created in the Directory of the .xls file")

#############################
