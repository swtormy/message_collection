import pandas as pd
import glob
from config import *

names = glob.glob(attachment_folder + '\*.xlsx', recursive=True)

dataframes = []

for el in names:
    new_df = pd.read_excel(el)
    col_count = len([el for el in pd.read_excel(el).iloc[0]])
    col_empty_sum = sum([1 if el is True else 0 for el in pd.read_excel(el).iloc[0].isnull()])
    if col_count == col_empty_sum:
        header = 0
        while col_count == col_empty_sum \
            or 'Unnamed: 0' in new_df.columns: #проверка пустых заголовков 
            header += 1
            new_df = pd.read_excel(el, header=header)
            col_count = len([el for el in new_df.iloc[0]])
            col_empty_sum = sum([1 if el is True else 0 for el in new_df.iloc[0].isnull()])
    name = el.split('\\')[-1:][0].split('_')[0]
    email = el.split('\\')[-1:][0].split('_')[1]
    new_df['Name'] = name
    new_df['Email'] = email
    cols_to_move = ['Name', 'Email']
    new_df = new_df[cols_to_move + [x for x in new_df.columns if x not in cols_to_move]]
    
    dataframes.append(new_df)


dataframe = pd.concat(dataframes, ignore_index=True)
dataframe.drop(columns = array_columns_to_delete, inplace=True)
dataframe.to_excel(name_resulting_file, index=False)
