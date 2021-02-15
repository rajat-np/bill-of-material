import pandas as pd
from openpyxl import Workbook

from utils import save_bom

INPUT_EXCEL = 'data/secondinput.xlsx'
OUTPUT_EXCEL = 'data/output.xlsx'

BASE_LEVEL = '.1'

# create pandas dataframe and ignore all empty rows.
df = pd.read_excel(INPUT_EXCEL, engine='openpyxl').dropna()

# create a workbok object to create sheets and add data to sheets
workbook = Workbook()

finished_goods = df['Item Name'].unique()

for finished_good in finished_goods:
    finished_good_df = df[df['Item Name'] == finished_good]
    # save bom for primary good, eg fan, toy etc
    raw_materials = finished_good_df.loc[finished_good_df['Level'] == BASE_LEVEL]
    save_bom(
        finished_good,
        raw_materials,
        workbook
    )

    # reverse the order of dataframe
    finished_good_df = finished_good_df.iloc[::-1]

    no_of_secondary_rawmaterial = 0

    for row_index in range(len(finished_good_df.index)-1):
        current_level = finished_good_df.iloc[row_index]['Level']
        next_level = finished_good_df.iloc[row_index+1]['Level']

        if current_level == BASE_LEVEL:
            continue
        else:
            if current_level != next_level:
                raw_material = finished_good_df.iloc[row_index +
                                                     1]['Raw material']
                raw_material_items = finished_good_df.iloc[(
                    row_index - no_of_secondary_rawmaterial):(row_index + 1)]
                save_bom(
                    raw_material,
                    raw_material_items,
                    workbook
                )
                no_of_secondary_rawmaterial = 0
            else:
                no_of_secondary_rawmaterial += 1

# remove Sheet that is created by default
workbook.remove(workbook['Sheet'])
workbook.save(filename=OUTPUT_EXCEL)
