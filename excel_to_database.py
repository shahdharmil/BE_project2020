import sqlite3

import openpyxl as xl
import pandas as pd

file = 'M_Hourly_Weather_Data_All_Station .xlsx'
out_path = 'outfile.xlsx'
database = 'data_16_17.db'

# FOR PREFORMATTING

wb = xl.load_workbook(file)
sheet = wb['Hourly Weather Data']
cell = sheet.cell(1, 1)
cell.value = 'DateTime'
sheet.delete_rows(3, 4)
wb.save(out_path)

# REFORMATTING DATAFRAME
# DON'T DELETE OUTPATH FILE

df = pd.read_excel(out_path, header=[0, 1])
# df.drop(index=[0, 1], inplace=True)
df.drop(columns=['PA.1', 'UV.1', 'UV'], level=1, inplace=True)
# df = df.set_index(('DateTime', 'Unnamed: 0_level_1'))

# SENDING TO DATABASE

conn = sqlite3.connect(database=database)
df.to_sql('Data', conn, if_exists='replace', index=('DateTime', 'Unnamed: 0_level_1'))
