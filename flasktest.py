import flask
from flask import request, jsonify
from flask_cors import CORS
import json
from record_keeperv01 import get_progress_on_file, get_files, get_data, get_single_file, fetch_steps, update_record_json
from config_manager import get_config, update_config
from excel_generator import generate_excel_CMA


app = flask.Flask(__name__)
CORS(app)
app.config["DEBUG"] = True

@app.route('/', methods=['GET'])
def home():
    return "<h1>API ENDPOINT</h1><p>Ignore this page. Close it down.</p>"

@app.route('/api/check', methods=['GET'])
def heartbeat():
    return 'SERVER IS ACTIVE'

# # A route to return all of the available entries in our catalog.
# @app.route('/api/test', methods=['GET'])
# def api_test():
#     # filepath = r'C:\Users\tom_b\Desktop\Documents\Python\vcoding\Liberia.xlsx'
#     #filepath = r'C:\\Users\\tom_b\\Desktop\\Documents\\Python\\vcoding\\excel\\\\Liberia.xlsx'
#     # test_data = get_initial_data_from_excel(filepath)
#     test_data = get_data('Angola_Master-File.xlsx')
#     return jsonify(test_data)

# @app.route('/api/test2', methods=['GET'])
# def api_test2():
#     # filepath = r'C:\Users\tom_b\Desktop\Documents\Python\vcoding\Liberia.xlsx'
#     # filepath = r'C:\\Users\\tom_b\\Desktop\\Documents\\Python\\vcoding\\excel\\\\Liberia.xlsx'
#     test2_data = get_files()
#     return jsonify(test2_data)

@app.route('/api/get_config', methods=['GET'])
def get_settings():
    config = get_config()
    return jsonify(config)

@app.route('/api/change_config/', methods=['POST'])
def change_config():
    new_config = json.loads(request.args['config'])
    print(new_config)
    result = update_config(new_config)
    return result

@app.route('/api/files', methods=['GET'])
def get_all_files():
    dict_of_files = get_files()
    return jsonify(dict_of_files)

@app.route('/api/select_file/', methods=['GET'])
def select_a_file():
    if 'file_name' in request.args:
        file_name = str(request.args['file_name'])
        print(file_name)
        payload = get_single_file(file_name)
        return jsonify(payload)
    else:
        print('nothing')
        return request.args

@app.route('/api/file_steps/', methods=['GET'])
def get_steps():
    if 'file_name' in request.args:
        if 'restart' in request.args:
            restart_bool = json.loads(str(request.args['restart']).lower())
            file_name = str(request.args['file_name'])
            payload = fetch_steps(file_name, restart_bool)
            # return jsonify({'steps':[1,2,3], 'next_step':1})
            return jsonify(payload)
        else:
            return 'Error'
    else:
        return 'Error'

@app.route('/api/update_record/', methods=['POST'])
def update_record():
    file_name = str(request.args['file_name'])
    year = str(request.args['year'])
    exchangeID = str(request.args['exchangeID'])
    answers = json.loads(request.args['answers'])
    review_flag = 'true' == request.args['reviewFlag']
    print(review_flag)
    result = update_record_json(file_name, year, exchangeID, answers, review_flag )
    # with open('test3.json', 'w') as output_file:
    #     json.dump(answers, output_file)
    return result


@app.route('/api/generate_excel/', methods=['GET'])
def generate_excel():
    if 'file_name' in request.args:
        file_name = str(request.args['file_name'])
    try:
        generate_excel_CMA(file_name)
        return 'Success'
    except expression as identifier:
        return "failed"
        
    
    

## List all files

## Show data for one file



app.run()