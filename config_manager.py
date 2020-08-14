import json
from pathlib import Path

#static variable names for dicts
excel_path = 'excel_path'


def create_config():
    config = {
        excel_path: Path('C:/Users/tom_b/Desktop/Documents/Python/vcoding/excel')
    }

    with open('config.json', 'w') as output_file:
        json.dump(config, output_file)


def get_config():
    with open('config.json', 'r') as config_json_file:
        config = json.load(config_json_file)
        return(config)

def update_config(new_config):
    with open('config.json', 'w') as config_json_file:
        json.dump(new_config, config_json_file)
    return 'DONE'

# create_config()
# print(get_config().keys())
