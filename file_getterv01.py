import os
import time
from config_manager import get_config

def get_list_of_all_files():
    path_to_watch = get_config()['excel_path']

    files_dict = {}
    for files in os.listdir(path_to_watch):
            files_dict[files] = path_to_watch + r'\\' + files
    return files_dict

def get_list_of_excel_files():
    all_files_dict = get_list_of_all_files()
    excel_files_dict = {}
    for file_name in all_files_dict:
        if file_name.endswith('.xlsx'):
            # print('its excel!')
            excel_files_dict[file_name] = all_files_dict[file_name]

        # print(file_name)
    return excel_files_dict

def check_if_json_record_exists(file_name):
    all_files_dict = get_list_of_all_files()
    # print(all_files_dict)
    json_file_name = create_json_filename_from_excel(file_name)
    # print(json_file_name)
    if json_file_name in all_files_dict.keys():
        return True
    else:
        return False

def create_json_filename_from_excel(file_name):
    no_extension_file_name = os.path.splitext(file_name)[0]
    json_file_name = no_extension_file_name + '.json'
    return json_file_name

# print(check_if_json_record_exists('Liberia.xlsx'))