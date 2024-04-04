# This is a cod for load from EXCEL
# Используем для работы с EXCEL библиотеку Pandas
import pandas as pd
import os

# каталог хранения EXCRL файлов
cwd = os.getcwd()

# каталог где хранятся файлы с данными парных сравнений
os.chdir('D:\\_project\excel-repo')

# название файла парных сравнений
file = 'compare-tables.xlsx'
xl = pd.ExcelFile(file)

# открытие листа "Лист1" (имеет нулевой индекс)
df1 = xl.parse(xl.sheet_names[0])