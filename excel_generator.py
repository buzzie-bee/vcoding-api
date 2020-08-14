import xlsxwriter
import json

from country_code_getter import get_country_code_id 
from config_manager import get_config
from file_getterv01 import check_if_json_record_exists, create_json_filename_from_excel
from record_keeperv01 import get_files

def generate_excel_CMA(file_name):
    # Definitions
    UCDPConflictID = 'UCDPConflictID'
    UCDPCountryID = 'UCDPCountryID'
    COWStateAbb = 'COWStateAbb'
    COWCCode = 'COWCCode'
    SideA = 'SideA'
    SideA_ID = 'SideA ID'
    SideB = 'SideB'
    SideB_ID = 'SideB ID'
    YEAR = 'YEAR'
    Region = 'Region'
    ExchangeID = 'ExchangeID'
    CMA = 'CMA'
    ExTCOWCode = 'ExTCOWCode'
    Client = 'Client'
    ClientCode = 'ClientCode'
    DClient = 'DClient'
    ForThird = 'ForThird'
    Consumer = 'Consumer'
    ConsumerID = 'ConsumerID'
    CoOrigin = 'CoOrigin'
    CoOriginCode = 'CoOriginCode'
    AgentId = 'AgentId'
    OpOrigin = 'OpOrigin'
    OpOriginCode = 'OpOriginCode'
    Task = 'Task'
    AgentSt = 'AgentSt'
    OwnSt = 'OwnSt'
    SizeBest = 'SizeBest'
    SizeMin = 'SizeMin'
    SizeMax = 'SIzeMax'
    Rly = 'Rly'

    headers = [
        UCDPConflictID,
        UCDPCountryID,
        COWStateAbb,
        COWCCode,
        SideA,
        SideA_ID,
        SideB,
        SideB_ID,
        YEAR,
        Region,
        ExchangeID,
        CMA,
        ExTCOWCode,
        Client,
        ClientCode,
        DClient,
        ForThird,
        Consumer,
        ConsumerID,
        CoOrigin,
        CoOriginCode,
        AgentId,
        OpOrigin,
        OpOriginCode,
        Task,
        AgentSt,
        OwnSt,
        SizeBest,
        SizeMin,
        SizeMax,
        Rly
    ]

    textBoxHeaders = {ClientCode: ClientCode, 'ConsumerCode' : ConsumerID, CoOriginCode: CoOriginCode, OpOriginCode: OpOriginCode, 'AgentIdCode': AgentId, }

    headerPosition = {}

    #Open up json file
    # json_path = r'C:\Users\tom_b\Desktop\Documents\Python\vcoding\Niger_New.json'
    # json_path = r'C:\Users\tom_b\Desktop\Documents\Python\vcoding\jsonfix\Chad_Master-File.json'
    folder_path = get_config()['excel_path']
    print(folder_path)
    destination_path = get_config()['destination_path']
    print(destination_path)


    #### IMPLEMENT A COMPLETED PATH IN CONFIG AND SETTINGS PAGES



    json_file_name = create_json_filename_from_excel(file_name)
    print(json_file_name)
    json_path = folder_path + json_file_name
    print(json_path)
    with open(json_path, 'r') as data_json_file:
            data = json.load(data_json_file)

    excel_completed_filename = 'CMA_TABLE_' + file_name
    excel_completed_filepath = destination_path + excel_completed_filename
    # Create an new Excel file and add a worksheet.yy
    # workbook = xlsxwriter.Workbook('test.xlsx')
    workbook = xlsxwriter.Workbook(excel_completed_filepath)
    worksheet = workbook.add_worksheet()

    # Add a bold format to use to highlight cells.
    bold = workbook.add_format({'bold': True})

    needs_review_format = workbook.add_format()
    needs_review_format.set_font_color('red')
    needs_review_format.set_bold()
    needs_review_format.set_font_size(30)


    for index, header in enumerate(headers):
        worksheet.write(0, index, header, bold)
        worksheet.write(0,index+1, 'REVIEW_FLAG', bold)
        headerPosition['REVIEW_FLAG'] = index + 1
        headerPosition[header] = index
        # worksheet.write(1, index, 'test!')
        # print(str(index) + ': ' + header)


    # print(headerPosition)

    countryCode = data['country']
    countryCodeId = get_country_code_id(countryCode)

    current_row = 1

    for index, yearIndex in enumerate(data.keys()):
        # print(index)
        # print(yearIndex)
        if type(data[yearIndex]) == dict:
            for record_index, record_id in enumerate(data[yearIndex]['records']):
                processed_bool = data[yearIndex]['records'][record_id]['Processed']
                review_bool = data[yearIndex]['records'][record_id]['Needs_Reviewing']
                if processed_bool:


                    worksheet.write(current_row, headerPosition[YEAR], yearIndex)
                    worksheet.write(current_row, headerPosition[ExchangeID], record_id)
                    worksheet.write(current_row, headerPosition[COWStateAbb], countryCode )
                    worksheet.write(current_row, headerPosition[COWCCode], countryCodeId )
                    size_best_arr = data[yearIndex]['records'][record_id][SizeBest]['Information']
                    size_min_arr = data[yearIndex]['records'][record_id][SizeMin]['Information']
                    size_max_arr = data[yearIndex]['records'][record_id][SizeMax]['Information']
                    size_best = ''
                    size_min = ''
                    size_max = ''
                    
                    for  index, line in enumerate(size_best_arr):
                        print(index)
                        if index > 1:
                            seperator = ' // '
                        else:
                            seperator = ''

                        size_best = size_best + str(line) + seperator
                    
                    for index, line  in enumerate(size_min_arr):
                        if index > 1:
                            seperator = ' // '
                        else:
                            seperator = ''
                        size_min = size_min + str(line) + seperator

                    for  index, line  in enumerate(size_max_arr):
                        if index > 1:
                            seperator = ' // '
                        else:
                            seperator = ''
                            size_max = size_max + str(line) + seperator


                    # if float - convert to string
                    # if array join

                    worksheet.write(current_row, headerPosition[SizeBest], size_best)
                    worksheet.write(current_row, headerPosition[SizeMin], size_min)
                    worksheet.write(current_row, headerPosition[SizeMax], size_max)
                    # print(review_bool)
                    if review_bool:
                        position = headerPosition['REVIEW_FLAG']
                        worksheet.write(current_row, position, '1', needs_review_format)
                        current_row = current_row + 1

                    else:
                        print(record_id)
                        for dataId in data[yearIndex]['records'][record_id]['answers']:
                            key = ''
                            value = ''
                            position = 50
                            # print(headerPosition[dataId])
                            if dataId in textBoxHeaders:
                                # print(dataId)
                                key = textBoxHeaders[dataId]
                                value = data[yearIndex]['records'][record_id]['answers'][dataId][0]
                                position = headerPosition[key]
                            elif dataId == 'AgentId':
                                print("")
                            else:
                                key = dataId
                                value = data[yearIndex]['records'][record_id]['answers'][dataId][1]
                                position = headerPosition[key]
                            worksheet.write(current_row, position, value )
                        current_row = current_row + 1

                            # if dataId == 'ConsumerCode':
                            #     print(dataId)
                            # worksheet.write(index + 1, headerPosition[dataId],data[yearIndex]['records'][record_id]['answers'][dataId][1] )



    workbook.close()

# excel_files = get_files()
# for file_name in excel_files.keys():
#     generate_excel_CMA(file_name)