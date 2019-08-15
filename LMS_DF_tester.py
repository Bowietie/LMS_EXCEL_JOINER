import pandas as pd
from glob import glob
import datetime as dt
'''
Program that reads files from work LMS reports and formates a
CSV with a preferred table build and all modules in on
'''


def name_locations(df_drop_na):
    trigger = list(df_drop_na.loc[:, 'one'])
    name_loc = [index for index in range(len(trigger))
                if trigger[index] == 'Client name']
    return name_loc


def range_finder(location_of_name):
    a = location_of_name[0]
    b = location_of_name[1]
    a_b_range = (b - a)
    return a_b_range


def list_maker(location_of_name, df_drop_na):
    names_list = []
    agent_names = list(df_drop_na.loc[:, 'two'])
    for loc in location_of_name:
        test = []
        agent_name = agent_names[loc]
        for name in range(range_finder(location_of_name)):
            test.append(agent_name)
        names_list.append(test)
    return names_list


def flatten(list_of_name):
    out = []
    for item in list_of_name:
        if isinstance(item, (list, tuple)):
            out.extend(flatten(item))
        else:
            out.append(item)
    return out


files = glob('*.xls')
all_data = pd.DataFrame()
for file in files:
    df = pd.read_excel(file, header=None, names=["one", "two", "three"])
    df_drop_na = df.dropna(how='all')
    location_of_name = name_locations(df_drop_na)
    for data in location_of_name:
        number_of_tasks = range_finder(location_of_name)
        list_of_name = list_maker(location_of_name, df_drop_na)
        flat_list = flatten(list_of_name)
        df_with_names = df_drop_na.assign(Four=flat_list)
    all_data = all_data.append(df_with_names, sort=False)
    filename1 = str(dt.datetime.now())
    test = str(file[:-4])
all_data.to_excel(excel_writer=('test' + filename1 + '.xlsx'),
                  sheet_name=test)
