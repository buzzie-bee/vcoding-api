import json
import os

from file_getterv01 import check_if_json_record_exists, create_json_filename_from_excel, get_list_of_excel_files
from parserv05 import get_initial_data_from_excel
from config_manager import get_config

def get_data(file_name):
    folder_path = get_config()['excel_path']
    excel_path = folder_path + file_name
    json_file_name = create_json_filename_from_excel(file_name)
    json_path = folder_path + json_file_name
    # print(json_path)
    if check_if_json_record_exists(file_name):
        # print('it exists!')
        with open(json_path) as json_file:
            data = json.load(json_file)
            return data
    else:
        raw_data = get_initial_data_from_excel(excel_path)
        with open(json_path, 'w') as output_file:
            json.dump(raw_data, output_file)
            return raw_data

def get_files():
    raw_file_dict = get_list_of_excel_files()
    file_dict = {}
    for file_name in raw_file_dict.keys():
        total, completed = get_progress_on_file(file_name)
        file_dict[file_name] = { 'file_name': file_name, 'path': raw_file_dict[file_name], 'total':total, 'completed': completed}
    return file_dict

def get_single_file(file_name):
    data = get_data(file_name)
    total, completed = get_progress_on_file(file_name)
    return {'file_name' : file_name, 'file_data': data,'total':total, 'completed': completed  }

def get_progress_on_file(file_name):
    data = get_data(file_name)
    # print(data)
    total_exchanges = 0
    completed_exchanges = 0
    for year in data.keys():
        if type(data[year]) == dict:
            total_exchanges = total_exchanges + data[year]['exchange_count']
            for conflict in data[year]['records'].keys():
                if data[year]['records'][conflict]['Processed']:
                    completed_exchanges = completed_exchanges + 1
    return(total_exchanges, completed_exchanges)
    # print("total exchanges = ", total_exchanges)
    # print('completed exchanges = ', completed_exchanges)

def fetch_steps(file_name, restart=None):
    data = get_data(file_name)
    next_step = None
    remaining_steps = []
    all_steps_list = []
    all_steps = {}
    restarting = False
    if restart == True:
        restarting == True
    for yearIndex in data.keys():
        if type(data[yearIndex]) == dict:
            for record_id in data[yearIndex]['records']:
                processed_bool = data[yearIndex]['records'][record_id]['Processed']
                flag_for_review_bool = data[yearIndex]['records'][record_id]['Needs_Reviewing']
                all_steps[record_id] = {'year': yearIndex, 'conflictId': record_id, 'reviewFlag': flag_for_review_bool, 'processed': processed_bool}
                all_steps_list.append(record_id)
                if restarting == True:
                    if next_step == None:
                        next_step == [yearIndex, record_id]
                    remaining_steps.append([yearIndex, record_id])
                else:
                    if processed_bool == False:
                        if next_step == None:
                            next_step = [yearIndex, record_id]
                        remaining_steps.append([yearIndex, record_id])
    print(next_step)
    return{'next_step': next_step, 'remaining_steps': remaining_steps, 'all_steps_info': all_steps, 'all_steps_list': all_steps_list}

def update_record_json(file_name, year, exchangeID, answers, review_flag):
    try:
        folder_path = get_config()['excel_path']
        json_file_name = create_json_filename_from_excel(file_name)
        json_path = folder_path + json_file_name

        data = get_data(file_name)
        record = data[year]['records'][exchangeID]
        record['Processed'] = True
        record['Needs_Reviewing'] = review_flag
        record['answers'] = answers
        data[year]['records'][exchangeID] = record
        with open(json_path, 'w') as output_file:
            json.dump(data, output_file)

        # temp_dict = { 'file_name': file_name, 'year': year, 'exchangeID': exchangeID, 'answers': answers}
        # with open('temp4.json', 'w') as output_file:
        #      json.dump(temp_dict, output_file)
        return 'Success'
    except:
        return 'Error'

def fix_record_json(file_name, year, exchangeID, review_flag):
    # try:
    #     folder_path = get_config()['excel_path']
    #     json_file_name = create_json_filename_from_excel(file_name)
    #     json_path = folder_path + json_file_name

    #     data = get_data(file_name)
    #     record = data[year]['records'][exchangeID]
    #     record['Processed'] = True
    #     record['Needs_Reviewing'] = review_flag
    #     record['answers'] = data[year]['records']['answers']
    #     data[year]['records'][exchangeID] = record
    #     with open(json_path, 'w') as output_file:
    #         json.dump(data, output_file)
    #         print('fixed!')
    #     return 'Success'
    # except:
    #     return 'Error'

    folder_path = get_config()['excel_path']
    json_file_name = create_json_filename_from_excel(file_name)
    json_path = folder_path + json_file_name

    data = get_data(file_name)
    record = data[year]['records'][exchangeID]
    record['Processed'] = True
    record['Needs_Reviewing'] = review_flag
    data[year]['records'][exchangeID] = record
    with open(json_path, 'w') as output_file:
        json.dump(data, output_file)
        print('fixed!')
    return 'Success'


def find_and_fix_incorrect_json_review_flags(file_name):
    print('fixing: ')
    print(file_name)
    print('\n\n')
    fixed = 0
    remaining = 0
    data = get_data(file_name)
    for yearIndex in data.keys():
        if type(data[yearIndex]) == dict:
            for record_id in data[yearIndex]['records']:
                exchange_id = record_id
                processed_bool = data[yearIndex]['records'][record_id]['Processed']
                flag_for_review_bool = data[yearIndex]['records'][record_id]['Needs_Reviewing']
                
                if processed_bool == True:
                    answers = data[yearIndex]['records'][record_id]['answers']
                    if flag_for_review_bool == True:
                        if bool(answers):
                            # print(yearIndex + '  ' + record_id + ' ------------------------- Wrong!')
                            fix_record_json(file_name, yearIndex, exchange_id, False)
                            fixed = fixed + 1
                        else:
                            remaining = remaining + 1
                            # print(yearIndex + '  ' + record_id + '  is correctly flagged for review')
    print(file_name + '    fixed: ' + str(fixed) + '   remaining: ' + str(remaining))
                            

# def update_all_json_files():
#     raw_file_dict = get_list_of_excel_files()
#     for file_name in raw_file_dict.keys():

#         find_and_fix_incorrect_json_review_flags(file_name)

# update_all_json_files()

# file_name = "Angola_Master-File.xlsx"
# year = "1980"
# exchangeID = "ANG19801"    
# answers =  {"CMA": ["Present", 1], "Client": ["National Gov", 1], "DClient": ["Defense Ministry", 1], "ForThird": ["Yes", 1], "Consumer": ["National Gov", 1], "CoOrigin": ["Europe", 1], "OpOrigin": ["Europe", 1], "Task": ["Military Training", 3], "AgentSt": ["International Company", 3], "OwnSt": ["Not applicable", 998]}

# update_record_json(file_name, year, exchangeID, answers)

# fetch_steps('Angola_Master-File.xlsx')

# get_progress_on_file('Liberia.xlsx')
# get_files()
# test_name = 'Liberia.xlsx'

# print(create_json_filename_from_excel(test_name))
# get_data(test_name)
