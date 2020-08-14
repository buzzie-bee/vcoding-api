# Reading an excel file using Python 3.x. or earlier
import xlrd as x
import os
import json
from random import randint
from country_code_getter import get_country_code



def get_initial_data_from_excel(file_path):
    exchange_count = "exchange_count"
    exchange_cell_ranges_key = "exchange_cell_ranges_key"
    records = "records"
    country = "country"
    Conflict_Number = "Conflict Number"
    Processed = "Processed"
    Needs_Reviewing = "Needs_Reviewing"
    Conflict_ID = "Conflict ID"
    Dyad_ID = "Dyad-ID"
    Year = "Year"
    Agent_or_CMA = "Agent/CMA"
    Client = "Client"
    DCLient = "DCLient"
    For_Third = "For Third"
    Consumer = "Consumer"
    Task = "Task"
    AgentSt = "AgentSt"
    Agent_ID = "Agent-ID"
    CompanyOrigin = "CompanyOrigin"
    OperatorOrigin = "OperatorOrigin"
    SizeBest = "SizeBest"
    SizeMin = "SizeMin"
    SIzeMax = "SIzeMax"
    Reliability = "Reliability"
    Note = "Note: "
    Information = "Information"
    Code = "Code"
    Source = "Source"
    cellRanges = "cellRanges"

    # Give the address of the file on the local computer, i.e, path location
    # country_code = "LBR" # TODO Lookup country code for relevant country file
    loc_file = (file_path)
    # To open Workbook we declare a hadling variable wb
    wb = x.open_workbook(loc_file)
    sheet_names = wb.sheet_names()
    data_worksheets = []

    for sheet_name in sheet_names:
        if len(sheet_name) == 4 and (sheet_name[0] == 1 or 2):
            data_worksheets.append(sheet_name)
    # print(data_worksheets)

    # Prints the value of element at row 0 and column 0
    # test = sheet.cell_value(0, 0)
    # print('Exchange' in test)

    country_code = get_country_code(os.path.basename(file_path))
    total_number_of_exchanges = 0
    workbook_status = {country: country_code}

    for index, sheet in enumerate(data_worksheets):
    # for index in range(1):
        sheet = wb.sheet_by_name(data_worksheets[index])
        number_of_exchanges = 0
        exchange_cell_ranges = []
        workbook_status[data_worksheets[index]] = {exchange_count: 0}

        #Check if any exchanges exist in table
        has_exchanges = False
        for i in range(sheet.nrows):
            if ('Exchange') in sheet.cell_value(i, 0):
                has_exchanges = True

        if has_exchanges:
            for i in range(sheet.nrows):
                if 'Exchange' in sheet.cell_value(i, 0):
                    if number_of_exchanges > 0 and exchange_cell_ranges[number_of_exchanges-1][1] == '':
                        exchange_cell_ranges[number_of_exchanges-1][1] = i
                    exchange_cell_ranges.append([i, ''])
                    number_of_exchanges = number_of_exchanges + 1
                
                if i == sheet.nrows - 1 and exchange_cell_ranges[number_of_exchanges-1][1] == '':
                    exchange_cell_ranges[number_of_exchanges-1][1] = i
            if exchange_cell_ranges[number_of_exchanges-1][1] == '':
                print('Error adding final position of index')
                exchange_cell_ranges[number_of_exchanges-1][1] = sheet.nrows

        # if index > 0:
            # print("data_worksheet: ", data_worksheets[index - 1])
            # print(exchange_cell_ranges)

            workbook_status[data_worksheets[index]][exchange_count] = number_of_exchanges
            workbook_status[data_worksheets[index]][exchange_cell_ranges_key] = exchange_cell_ranges
            total_number_of_exchanges = total_number_of_exchanges + number_of_exchanges
        else:
            workbook_status.pop(data_worksheets[index], None)

    # print("total: ", total_number_of_exchanges)
    # print(workbook_status)
    # print(number_of_exchanges)
    # print(exchange_cell_ranges)

    # for key in workbook_status:
    #     print(key)

    count = 0
    for wbkey in workbook_status:
        count = count + 1

        exchange_code = ''
        exchange_dictionary = {
            Conflict_Number: 0,
            Processed: False,
            Needs_Reviewing: False,
            Conflict_ID: {cellRanges: [], Information: [], Code: [], Source: []},
            Dyad_ID: {cellRanges: [], Information: [], Code: [], Source: []},
            Year: {cellRanges: [], Information: [], Code: [], Source: []},
            Agent_or_CMA: {cellRanges: [], Information: [], Code: [], Source: []},
            Client: {cellRanges: [], Information: [], Code: [], Source: []},
            DCLient: {cellRanges: [], Information: [], Code: [], Source: []},
            For_Third: {cellRanges: [], Information: [], Code: [], Source: []},
            Consumer: {cellRanges: [], Information: [], Code: [], Source: []},
            Task: {cellRanges: [], Information: [], Code: [], Source: []},
            AgentSt: {cellRanges: [], Information: [], Code: [], Source: []},
            Agent_ID: {cellRanges: [], Information: [], Code: [], Source: []},
            CompanyOrigin: {cellRanges: [], Information: [], Code: [], Source: []},
            OperatorOrigin: {cellRanges: [], Information: [], Code: [], Source: []},
            SizeBest: {cellRanges: [], Information: [], Code: [], Source: []},
            SizeMin: {cellRanges: [], Information: [], Code: [], Source: []},
            SIzeMax: {cellRanges: [], Information: [], Code: [], Source: []},
            Reliability: {cellRanges: [], Information: [], Code: [], Source: []},
            Note: {cellRanges: [], Information: [], Code: [], Source: []},
        }
        if type(workbook_status[wbkey]) == dict:
            sheet = wb.sheet_by_name(wbkey)
            current_key = ""
            for i in range(workbook_status[wbkey][exchange_count]):
                # print(workbook_status[wbkey][exchange_count])
                # for i in range(1):
                exchange_dictionary[Conflict_Number] = i + 1
                # for j in range(exchange_dictionary[wbkey][cellRanges][0], exchange_dictionary[key][cellRanges][1]):

                for j in range(workbook_status[wbkey][exchange_cell_ranges_key][i][0], workbook_status[wbkey][exchange_cell_ranges_key][i][1]):
                    # print(workbook_status[wbkey][exchange_cell_ranges_key][i][0])
                    # print(workbook_status[wbkey][exchange_cell_ranges_key][i][1])
                    # print(j)
                    value = sheet.cell_value(j, 0)
                    # Reliability - 1989 not getting second value in here
                    if value in exchange_dictionary:
                        if current_key != "":
                            exchange_dictionary[current_key][cellRanges].append(j)
                        # print(value)
                        exchange_dictionary[value][cellRanges].append(j)
                        current_key = value
                        if value == Note:
                            exchange_dictionary[value][cellRanges].append(
                                workbook_status[wbkey][exchange_cell_ranges_key][i][1])
                # Clean up any missing range values
                for ex_dic_key in exchange_dictionary.keys():
                    # print("Ex Dic Key")
                    # print(ex_dic_key)
                    # print("Ex Dic")
                    # print(exchange_dictionary)
                    temp_key_type = type(exchange_dictionary[ex_dic_key])

                    if temp_key_type == list or temp_key_type == dict or temp_key_type == tuple:
                        field_len = len(exchange_dictionary[ex_dic_key][cellRanges])
                        if field_len == 1:
                            exchange_dictionary[ex_dic_key][cellRanges].append(workbook_status[wbkey][exchange_cell_ranges_key][i][1])

                for key in exchange_dictionary.keys():
                    value = ''
                    key_type = type(exchange_dictionary[key])

                    if key_type == list or key_type == dict or key_type == tuple:
                        if len(exchange_dictionary[key]) > 2:
                            if len(exchange_dictionary[key][cellRanges]) > 0:
                                # print("WB KEY")
                                # print(wbkey)
                                # print('WB KEY DICT')
                                # print(workbook_status[wbkey])
                                # print("RANGES LIST")
                                # print(workbook_status[wbkey][exchange_cell_ranges_key])
                                # print("RANGE")
                                # print(workbook_status[wbkey]
                                #       [exchange_cell_ranges_key])
                                # print('\n')
                                # print("exchange dict key")
                                # print(key)
                                # print("Exchange cell ranges:")
                                # print(exchange_dictionary[key][cellRanges])
                                for k in range(exchange_dictionary[key][cellRanges][0], exchange_dictionary[key][cellRanges][1]):
                                    # print(k)
                                    value = sheet.cell_value(k, 1)
                                    if value != '':
                                        if key == Year and type(value) == float:
                                            exchange_dictionary[key][Information].append(
                                                int(value))
                                        else:
                                            exchange_dictionary[key][Information].append(
                                                value)
                                value = ''

                                for l in range(exchange_dictionary[key][cellRanges][0], exchange_dictionary[key][cellRanges][1]):
                                    value = sheet.cell_value(l, 2)
                                    if value != '':
                                        exchange_dictionary[key][Code].append(value)
                                value = ''

                                for m in range(exchange_dictionary[key][cellRanges][0], exchange_dictionary[key][cellRanges][1]):
                                    value = sheet.cell_value(m, 3)
                                    if value != '':
                                        exchange_dictionary[key][Source].append(value)
                                value = ''
                # if count < 3:
                #     print(exchange_dictionary)
                if len(exchange_dictionary[Year][Information]) > 0:
                    exchange_code = country_code + str(exchange_dictionary[Year][Information][0]) + str(exchange_dictionary[Conflict_Number])
                else:
                    exchange_code = wbkey + '_' + country_code + '_E_Y_MISS_' + str(randint(10,90))
                
                if records not in workbook_status[wbkey]:
                    workbook_status[wbkey][records] = {}
                workbook_status[wbkey][records][exchange_code] = exchange_dictionary

                exchange_dictionary = {
                    Conflict_Number: 0,
                    Processed: False,
                    Needs_Reviewing: False,
                    Conflict_ID: {cellRanges: [], Information: [], Code: [], Source: []},
                    Dyad_ID: {cellRanges: [], Information: [], Code: [], Source: []},
                    Year: {cellRanges: [], Information: [], Code: [], Source: []},
                    Agent_or_CMA: {cellRanges: [], Information: [], Code: [], Source: []},
                    Client: {cellRanges: [], Information: [], Code: [], Source: []},
                    DCLient: {cellRanges: [], Information: [], Code: [], Source: []},
                    For_Third: {cellRanges: [], Information: [], Code: [], Source: []},
                    Consumer: {cellRanges: [], Information: [], Code: [], Source: []},
                    Task: {cellRanges: [], Information: [], Code: [], Source: []},
                    AgentSt: {cellRanges: [], Information: [], Code: [], Source: []},
                    Agent_ID: {cellRanges: [], Information: [], Code: [], Source: []},
                    CompanyOrigin: {cellRanges: [], Information: [], Code: [], Source: []},
                    OperatorOrigin: {cellRanges: [], Information: [], Code: [], Source: []},
                    SizeBest: {cellRanges: [], Information: [], Code: [], Source: []},
                    SizeMin: {cellRanges: [], Information: [], Code: [], Source: []},
                    SIzeMax: {cellRanges: [], Information: [], Code: [], Source: []},
                    Reliability: {cellRanges: [], Information: [], Code: [], Source: []},
                    Note: {cellRanges: [], Information: [], Code: [], Source: []},
                }
                
    return workbook_status

# filepath = r'C:\Users\tom_b\Desktop\Documents\Python\vcoding\excel\Angola_Master-File.xlsx'
# data = get_initial_data_from_excel(filepath)
# with open('temp.json', 'w') as outfile:
#     json.dump(data, outfile)
# # print(json.dumps(data))




# print(json.dumps(workbook_status))

# print("\n\n\n")
# print(json.dumps(exchange_dictionary))


#########
# TODO Add a 'needs reviewing' flag to each exchange.
# TODO Add info such as country code, year, etc to the dictionary.
# TODO Once a dictionary is built, store it somewhere and clean up variables.
# TODO Maybe generate a json document with ALL of the years in an excel, then maybe do that for ALL of the cases.
# TODO Look into using flask to do the python - react side of things, or maybe just use react and import a json of all the documents?
