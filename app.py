import pandas as pd

INPUT_EXCEL = 'data/input.xlsx'
OUTPUT_EXCEL = 'data/output.xlsx'

#create pandas dataframe and ignore all empty rows.
df = pd.read_excel(INPUT_EXCEL, engine='openpyxl').dropna()

print(df)
