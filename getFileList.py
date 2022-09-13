import glob

#Take the csv files and path in the origin directory
def get_csv_files_path(directory):
	files_path = glob.glob(directory + '\\_CAD\\*.csv')

	return files_path

#Take the xls files and path in the origin directory
def get_xls_files_path(directory):
	files_path = glob.glob(directory + '\\_ESTRUTURA\\*.xls')

	return files_path

#Extract just the file name, for the input form
def extract_file_names(list_of_path):
	files_name = []

	for path in list_of_path:
		start_caracter = path.rindex('\\') + 1
		files_name.append(path[start_caracter:])

	return files_name
