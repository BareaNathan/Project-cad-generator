import tkinter as tk
from tkinter import ttk
from getPartAndRefs import *
from getFileList import *

def main_function(chosen_xls,chosen_csv):
	index_for_xls = xls_files_name.index(chosen_xls)
	path_for_xls = xls_files[index_for_xls]

	index_for_csv = csv_files_name.index(chosen_csv)
	path_for_csv = csv_files[index_for_csv]

	write_output_file(path_for_xls,path_for_csv)
	

def write_output_file(path_xls,path_csv):
	#Calling the Part and Refs functions
	sheet_name = get_spreadsheet(path_xls)	#xls_files[40] 
	ref_list, part_list = get_columns(sheet_name)
	ref, part = clean_lists(ref_list, part_list)

	cad_type = CAD.get()

	csv_data = get_csv_data(path_csv,cad_type)	#csv_files[80]
	joined_data = join_csv_and_xls(csv_data,ref,part)


	#Edit Path and name to output CSV file
	ref_file_path = path_xls	#xls_files[40]
	new_file_path = ref_file_path[:-4] + '-FUJI.csv'
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
top_cont_label_title = tk.Label(top_cont,text='Label').grid()


mid_cont_1 = tk.Frame(root)
mid_cont_1_label_title = tk.Label(mid_cont_1,text='Label').grid()
mid_cont_1_listbox = ttk.Combobox(
	mid_cont_1,
	state = 'readonly',
	width = 25,
	values = xls_files_name_sorted
	)
mid_cont_1_listbox.set('texto')
mid_cont_1_listbox.grid()


mid_cont_2 = tk.Frame(root)
mid_cont_2_label_title = tk.Label(mid_cont_2,text='Label').grid()
mid_cont_2_listbox = ttk.Combobox(
	mid_cont_2,
	state = 'readonly',
	width = 50,
	values = csv_files_name_sorted
	)
mid_cont_2_listbox.set('texto')
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
bot_cont_label_title = tk.Label(bot_cont,text='Label').grid()
bot_cont_button = tk.Button(bot_cont,
	text='Clique-me',
	##Set these values in main.py scrpit
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