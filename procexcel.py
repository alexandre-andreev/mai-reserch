# This is a cod for load from EXCEL
# Используем для работы с EXCEL библиотеку Pandas
import pandas as pd
import os
import numpy as np

# каталог хранения EXCEL файлов
# cwd = os.getcwd()

# каталог, где хранятся файлы с данными парных сравнений
os.chdir('../excel-repo')
cwd = os.getcwd()
print(cwd)

# название файла парных сравнений
file = 'compare-tables.xlsx'
xbook = pd.ExcelFile(file)
# df = pd.read_excel(file)
# print(df.head())

# делаем справочник листов в файле

data_sheets = {}
for sheets in xbook.sheet_names:
    data_sheets[sheets] = xbook.parse(sheets)
print(data_sheets['Лист1'])
print(data_sheets['Лист2'])
print(data_sheets['Лист3'])

# открытие листа "Лист1" (имеет нулевой индекс)
df1 = xbook.parse(xbook.sheet_names[0], usecols=[1, 2, 3, 4, 5])
# df1 = pd.read_excel(file, usecols=[1, 2, 3, 4, 5])

# размерность таблицы с данными на листе 1
print(data_sheets['Лист1'].shape)

# доступ к ячейкам дата фрейма
print(df1.iloc[1][2])
# print(df1.at[1, 'c2'])

# вычисление свойств матрицы
# преобразование серий дата фреймов в матрицу
# получение строк
print(df1.iloc[0].values)
print(df1.iloc[1].values)
print(df1.iloc[2].values)
print(df1.iloc[3].values)
print('-----')
# получение матрицы
new_array = df1.to_numpy()
print(new_array, new_array[4, 2], sep='   __')
det = np.linalg.det(new_array)
print('Определитель матрицы = ', det)
