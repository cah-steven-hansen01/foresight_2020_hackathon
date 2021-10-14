import pandas as pd

current_so = pd.read_csv('./clean_data/current_so.csv')
print(current_so.head())
print(len(current_so['Lot Number'].unique()))