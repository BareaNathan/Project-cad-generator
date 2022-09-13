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
		write_output_file(path_xls,joined_data,'FUJI.csv')

	elif board_layer_conf == 'TOP+BOT':
		temp_csv_top = copy.deepcopy(csv_data)
		print(temp_csv_top)

		ref_list, part_list = get_columns(sheet_name,1,5)
		ref, part = clean_lists(ref_list, part_list)
		joined_data = join_csv_and_xls(temp_csv_top,ref,part)
		write_output_file(path_xls,joined_data,'FUJI.csv')

		temp_csv_bot = copy.deepcopy(csv_data)
		print(temp_csv_bot)
		ref_list_bot, part_list_bot = get_columns(sheet_name,3,6)
		ref_bot, part_bot = clean_lists(ref_list_bot, part_list_bot)
		joined_data_bot = join_csv_and_xls(temp_csv_bot,ref_bot,part_bot)
		write_output_file(path_xls,joined_data_bot,'FUJI-BOT.csv')



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