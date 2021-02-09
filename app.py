import pandas as pd
from openpyxl import Workbook

from utils import save_bom

INPUT_EXCEL = 'data/input.xlsx'
OUTPUT_EXCEL = 'data/output.xlsx'

BASE_LEVEL = '.1'

# create pandas dataframe and ignore all empty rows.
df = pd.read_excel(INPUT_EXCEL, engine='openpyxl').dropna()

# create a workbok object to create sheets and add data to sheets
workbook = Workbook()

finished_goods = df['Item Name'].unique()

for finished_good in finished_goods:
    finished_good_df = df[df['Item Name'] == finished_good]
    # get unique levels for a finished good
    levels = finished_good_df['Level'].unique()

    composite_rawmaterials = {}
    # get all composite raw materials with their level
    for level_index in range(len(levels)-1):
        for row_index in range(len(finished_good_df.index)-1):
            if(finished_good_df.iloc[row_index]['Level'] == levels[level_index]
               and finished_good_df.iloc[row_index+1]['Level'] == levels[level_index+1]):
                composite_rawmaterials[levels[level_index+1]] = finished_good_df.iloc[row_index]['Raw material']

    for level in levels:
        for row_index in range(1, len(finished_good_df.index)+1):
            if level == finished_good_df.iloc[row_index]['Level']:
                raw_materials = finished_good_df.loc[finished_good_df['Level'] == level]
                if level == BASE_LEVEL:
                    item = finished_good_df.iloc[row_index-1]['Item Name']
                else:
                    item = finished_good_df.iloc[row_index-1]['Raw material']
                save_bom(item, raw_materials, workbook)
                break

# remove Sheet that is created by default
workbook.remove(workbook['Sheet']) 
workbook.save(filename=OUTPUT_EXCEL)
