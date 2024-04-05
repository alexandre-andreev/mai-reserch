# This is a cod for load from EXCEL
# Используем для работы с EXCEL библиотеку Pandas
import pandas as pd
import os

# каталог хранения EXCEL файлов
# cwd = os.getcwd()

# каталог где хранятся файлы с данными парных сравнений
os.chdir('D:\\_project\excel-repo')

# название файла парных сравнений
file = 'compare-tables.xlsx'
xbook = pd.ExcelFile(file)
# df = pd.read_excel(file)
# print(df.head())

# делаем справочник листов в файле
# открытие листа "Лист1" (имеет нулевой индекс)
df1 = xbook.parse(xbook.sheet_names[0])

data_sheets = {}
for sheets in xbook.sheet_names:
    data_sheets[sheets] = xbook.parse(sheets)
print(data_sheets['Лист1'])
print(data_sheets['Лист2'])
print(data_sheets['Лист3'])
print(data_sheets['Лист1'].shape) # размерность таблицы с данными
