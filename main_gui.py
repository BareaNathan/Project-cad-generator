import tkinter as tk
import copy
from tkinter import ttk
from getPartAndRefs import *
from getFileList import *

def main_function(chosen_xls,chosen_csv):
	index_for_xls = xls_files_name.index(chosen_xls)
	path_for_xls = xls_files[index_for_xls]

	index_for_csv = csv_files_name.index(chosen_csv)
	path_for_csv = csv_files[index_for_csv]

	call_write_function(path_for_xls,path_for_csv)
	
def call_write_function(path_xls,path_csv):
	#Calling the Part and Refs functions
	sheet_name = get_spreadsheet(path_xls)	#xls_files[40] 

	cad_type = CAD.get()

	csv_data = get_csv_data(path_csv,cad_type)	#csv_files[80]


	board_layer_conf = layer.get()

	if board_layer_conf == 'Single':
		ref_list, part_list = get_columns(sheet_name,1,4)
		ref, part = clean_lists(ref_list, part_list)

		joined_data = join_csv_and_xls(csv_data,ref,part)
		joined_data_sorted = sorted(joined_data[1:],key = lambda x: x[4])
		joined_data_sorted.insert(0,joined_data[0])


		if FUJI.get() == '1':
			fuji_list = copy.deepcopy(joined_data_sorted)
			for row in fuji_list:
				if row[1][-2:] == 'mm':
					row[1] = row[1][:-2]
					row[2] = row[2][:-2]
				try:
					row.append(csv_comp[row[4]][3])
					row.append(csv_comp[row[4]][4])
				except:
					row.append('Not Found')
					row.append('Not Found')

			write_output_file(path_xls,fuji_list,'FUJI.csv')

		if SAKI3D.get() == '1':
			saki3d_list = copy.deepcopy(joined_data_sorted)
			for row in saki3d_list:
				if row[1][-2:] == 'mm':
					row[1] = row[1][:-2]
					row[2] = row[2][:-2]
				try:
					row.append(csv_comp[row[4]][3])
					row.append(csv_comp[row[4]][4])
				except:
					row.append('Not Found')
					row.append('Not Found')
					
			write_output_file(path_xls,saki3d_list,'SAKI3D.csv')

		if MIRAE.get() == '1':
			mirae_list = copy.deepcopy(joined_data_sorted)
			for row in mirae_list:
				try:
					row.append(csv_comp[row[4]][3])
					row.append(csv_comp[row[4]][4])
					row[4] = csv_comp[row[4]][2]
				except:
					row.append('Not Found')
					row.append('Not Found')
					
			write_output_file(path_xls,mirae_list,'MIRAE.csv')

		if SAKI2D.get() == '1':
			saki2d_list = copy.deepcopy(joined_data_sorted)
			for row in saki2d_list:
				try:
					row.append(csv_comp[row[4]][3])
					row.append(csv_comp[row[4]][4])
					row[4] = csv_comp[row[4]][1]
				except:
					row.append('Not Found')
					row.append('Not Found')
					
			write_output_file(path_xls,saki2d_list,'SAKI2D.csv')


	elif board_layer_conf == 'TOP+BOT':
		temp_csv_top = copy.deepcopy(csv_data)
		print(temp_csv_top)

		ref_list, part_list = get_columns(sheet_name,1,5)
		ref, part = clean_lists(ref_list, part_list)

		joined_data = join_csv_and_xls(temp_csv_top,ref,part)
		joined_data_sorted = sorted(joined_data[1:],key = lambda x: x[4])
		joined_data_sorted.insert(0,joined_data[0])

		if FUJI.get() == '1':
			fuji_list = copy.deepcopy(joined_data_sorted)
			for row in fuji_list:
				if row[1][-2:] == 'mm':
					row[1] = row[1][:-2]
					row[2] = row[2][:-2]
				try:
					row.append(csv_comp[row[4]][3])
					row.append(csv_comp[row[4]][4])
				except:
					row.append('Not Found')
					row.append('Not Found')

			write_output_file(path_xls,fuji_list,'FUJI.csv')

		if SAKI3D.get() == '1':
			saki3d_list = copy.deepcopy(joined_data_sorted)
			for row in saki3d_list:
				if row[1][-2:] == 'mm':
					row[1] = row[1][:-2]
					row[2] = row[2][:-2]
				try:
					row.append(csv_comp[row[4]][3])
					row.append(csv_comp[row[4]][4])
				except:
					row.append('Not Found')
					row.append('Not Found')
					
			write_output_file(path_xls,saki3d_list,'SAKI3D.csv')

		if MIRAE.get() == '1':
			mirae_list = copy.deepcopy(joined_data_sorted)
			for row in mirae_list:
				try:
					row.append(csv_comp[row[4]][3])
					row.append(csv_comp[row[4]][4])
					row[4] = csv_comp[row[4]][2]
				except:
					row.append('Not Found')
					row.append('Not Found')
					
			write_output_file(path_xls,mirae_list,'MIRAE.csv')

		if SAKI2D.get() == '1':
			saki2d_list = copy.deepcopy(joined_data_sorted)
			for row in saki2d_list:
				try:
					row.append(csv_comp[row[4]][3])
					row.append(csv_comp[row[4]][4])
					row[4] = csv_comp[row[4]][1]
				except:
					row.append('Not Found')
					row.append('Not Found')
					
			write_output_file(path_xls,saki2d_list,'SAKI2D.csv')

		#write_output_file(path_xls,joined_data_sorted,'FUJI.csv')

		temp_csv_bot = copy.deepcopy(csv_data)
		print(temp_csv_bot)
		ref_list_bot, part_list_bot = get_columns(sheet_name,3,6)
		ref_bot, part_bot = clean_lists(ref_list_bot, part_list_bot)

		joined_data_bot = join_csv_and_xls(temp_csv_bot,ref_bot,part_bot)
		joined_data_bot_sorted = sorted(joined_data_bot[1:],key = lambda x: x[4])
		joined_data_bot_sorted.insert(0,joined_data_bot[0])

		if FUJI.get() == '1':
			fuji_list_bot = copy.deepcopy(joined_data_bot_sorted)
			for row in fuji_list_bot:
				if row[1][-2:] == 'mm':
					row[1] = row[1][:-2]
					row[2] = row[2][:-2]
				try:
					row.append(csv_comp[row[4]][3])
					row.append(csv_comp[row[4]][4])
				except:
					row.append('Not Found')
					row.append('Not Found')

			write_output_file(path_xls,fuji_list_bot,'FUJI-BOT.csv')

		if SAKI3D.get() == '1':
			saki3d_list_bot = copy.deepcopy(joined_data_bot_sorted)
			for row in saki3d_list_bot:
				if row[1][-2:] == 'mm':
					row[1] = row[1][:-2]
					row[2] = row[2][:-2]
				try:
					row.append(csv_comp[row[4]][3])
					row.append(csv_comp[row[4]][4])
				except:
					row.append('Not Found')
					row.append('Not Found')
					
			write_output_file(path_xls,saki3d_list_bot,'SAKI3D-BOT.csv')

		if MIRAE.get() == '1':
			mirae_list_bot = copy.deepcopy(joined_data_bot_sorted)
			for row in mirae_list_bot:
				try:
					row.append(csv_comp[row[4]][3])
					row.append(csv_comp[row[4]][4])
					row[4] = csv_comp[row[4]][2]
				except:
					row.append('Not Found')
					row.append('Not Found')
					
			write_output_file(path_xls,mirae_list_bot,'MIRAE-BOT.csv')

		if SAKI2D.get() == '1':
			saki2d_list_bot = copy.deepcopy(joined_data_bot_sorted)
			for row in saki2d_list_bot:
				try:
					row.append(csv_comp[row[4]][3])
					row.append(csv_comp[row[4]][4])
					row[4] = csv_comp[row[4]][1]
				except:
					row.append('Not Found')
					row.append('Not Found')
					
			write_output_file(path_xls,saki2d_list_bot,'SAKI2D-BOT.csv')

		#write_output_file(path_xls,joined_data_bot_sorted,'FUJI-BOT.csv')



def write_output_file(path_xls,joined_data,sulfix):
	
	#Edit Path and name to output CSV file
	ref_file_path = path_xls	#xls_files[40]
	new_file_path = ref_file_path[:-4] + '-' + sulfix
	print(new_file_path)

	#Output data
	output_csv(new_file_path,joined_data)

#PATH for all the CSV ans XLS files
files_path_origin = '\\\\10.0.0.3\\producao\\01 - PRODUCAO__BACKUP\\CEP\\Indicadores da Produção\\SMT & THT\\Arquivos de Programas\\*'

#Getting the PATH of each file
csv_files = get_csv_files_path(files_path_origin)
xls_files = get_xls_files_path(files_path_origin)

#Extracting the name of the files
csv_files_name = extract_file_names(csv_files)
csv_files_name_sorted = csv_files_name.copy()
csv_files_name_sorted.sort()

xls_files_name = extract_file_names(xls_files)
xls_files_name_sorted = xls_files_name.copy()
xls_files_name_sorted.sort()

#Getting the Components data
csv_comp = {}
with open(files_path_origin[:-1]+'Componentes SMT - Componentes.csv') as comps:
	csv_sheet = csv.DictReader(comps,fieldnames = ('FUJI','SAKI2D','MIRAE','Shape','Package'))

	for row in csv_sheet:
		csv_comp[row['FUJI']] = [row['FUJI'],row['SAKI2D'],row['MIRAE'],row['Shape'],row['Package']]

csv_comp['Partnumber'] = ['FUJI', 'SAKI2D', 'MIRAE', 'Shape', 'Package']

root = tk.Tk()

#Root container configure
root.rowconfigure(index=0,weight=1)
root.rowconfigure(index=1,weight=1)
root.rowconfigure(index=2,weight=1)
root.columnconfigure(index=0,weight=1)
root.columnconfigure(index=1,weight=1)

#Frames and Widgets
top_cont = tk.Frame(root)
top_cont_label_title = tk.Label(top_cont,text='CSV Generator 1.2').grid()


mid_cont_1 = tk.Frame(root)
mid_cont_1_label_title = tk.Label(mid_cont_1,text='XLS').grid()
mid_cont_1_listbox = ttk.Combobox(
	mid_cont_1,
	state = 'readonly',
	width = 25,
	values = xls_files_name_sorted
	)
mid_cont_1_listbox.set('Selecione o Arquivo...')
mid_cont_1_listbox.grid()

layer = tk.StringVar(mid_cont_1,value='',name = 'layer')

mid_cont_1_checkbutton_1 = ttk.Radiobutton(
	mid_cont_1,
	text = 'Single Layer Board',
	variable = 'layer',
	value = 'Single'
	).grid()

mid_cont_1_checkbutton_3 = ttk.Radiobutton(
	mid_cont_1,
	text = 'TOP+BOT Layer Board',
	variable = 'layer',
	value = 'TOP+BOT'
	).grid()

mid_cont_2 = tk.Frame(root)
mid_cont_2_label_title = tk.Label(mid_cont_2,text='CAD').grid()
mid_cont_2_listbox = ttk.Combobox(
	mid_cont_2,
	state = 'readonly',
	width = 50,
	values = csv_files_name_sorted
	)
mid_cont_2_listbox.set('Selecione o Arquivo...')
mid_cont_2_listbox.grid()

CAD = tk.StringVar(mid_cont_2,value="",name = 'CAD')

mid_cont_2_checkbutton_1 = ttk.Radiobutton(
	mid_cont_2,
	text = 'Altium CAD',
	variable = 'CAD',
	value = 'Altium'
	).grid()

mid_cont_2_checkbutton_2 = ttk.Radiobutton(
	mid_cont_2,
	text = 'Protel CAD',
	variable = 'CAD',
	value = 'Protel'
	).grid()


bot_cont = tk.Frame(root)

FUJI = tk.StringVar(bot_cont,value='',name='FUJI')
MIRAE = tk.StringVar(bot_cont,value='',name='MIRAE')
SAKI2D = tk.StringVar(bot_cont,value='',name='SAKI2D')
SAKI3D = tk.StringVar(bot_cont,value='',name='SAKI3D')

bot_cont_checkbutton_1 = ttk.Checkbutton(
	bot_cont,
	text = 'FUJI',
	variable = 'FUJI',
	).grid()

bot_cont_checkbutton_2 = ttk.Checkbutton(
	bot_cont,
	text = 'MIRAE',
	variable = 'MIRAE',
	).grid()

bot_cont_checkbutton_3 = ttk.Checkbutton(
	bot_cont,
	text = 'SAKI2D',
	variable = 'SAKI2D',
	).grid()

bot_cont_checkbutton_4 = ttk.Checkbutton(
	bot_cont,
	text = 'SAKI3D',
	variable = 'SAKI3D',
	).grid()

bot_cont_label_title = tk.Label(bot_cont,text='Clique em OK para gerar o(s) arquivo(s) CSV').grid()
bot_cont_button = tk.Button(bot_cont,
	text='OK',
	command = lambda: main_function(
		mid_cont_1_listbox.get(),
		mid_cont_2_listbox.get()
		)
	).grid()


#Grid the containers
top_cont.grid(row=0,columnspan=2)
mid_cont_1.grid(row=1,column=0)
mid_cont_2.grid(row=1,column=1)
bot_cont.grid(row=2,columnspan=2)

root.mainloop()