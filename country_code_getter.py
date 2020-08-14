import json

def get_country_code(file_name):
    with open('country_codes.json', 'r') as country_code_file:
        country_codes = json.load(country_code_file)
    

    country = None
    if '.' in file_name.split('_')[0]:
        country = file_name.split('.')[0].lower()
    else:
        country = file_name.split('_')[0].lower()
        
    country_code = None
    for code in country_codes:
        if country_codes[code]['StateNme'].lower() == country:
            country_code = code
        else:
            # print(country)
            if country in code.lower() :
                # print(code, ' : ',  country, '  Found it!')
                country_code = code
            if country == 'car':
                country_code = 'CEN'
        

    if (country_code):
        return country_code
    else:
        return 'Error finding country code ' + file_name + '  , ' + country

def get_country_code_id(country_code):
    with open('country_codes.json', 'r') as country_code_file:
        country_codes = json.load(country_code_file)
        return int(country_codes[country_code]['CCode'])




def generate_country_code_json():
    StateAbb = 'StateAbb'
    CCode = 'CCode'
    StateNme = 'StateNme'

    loc_file = (r'C:\\Users\\tom_b\\Desktop\\Documents\\Python\\vcoding\\CMA-COW State acronyms.xlsx')
    wb = x.open_workbook(loc_file)
    sheet = wb.sheet_by_index(0)

    country_codes = {}

    for i in range(sheet.nrows - 1):
        Abb = sheet.cell_value(i + 1, 0)
        Code = sheet.cell_value(i + 1, 1)
        Name = sheet.cell_value(i + 1, 2)
        country_codes[Abb] = {StateAbb: Abb, CCode: Code, StateNme: Name}

    with open('country_codes.json', 'w') as output_file:
        json.dump(country_codes, output_file)

# generate_country_code_json()
# print(get_country_code('Angola_Master-File.xlsx'))