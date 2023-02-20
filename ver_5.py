import os
import fitz
import pandas as pd
from pandas import ExcelWriter
from my_functions import search_for_text_in_block, search_for_text

all_files = []
dirpath = r"C:\Users\adilet.yensepov\Desktop\tabula_pdf_reader\pdf_files"

for dirpath, dirnames, filenames in os.walk(dirpath): #сохраняем все pdf файлы в список all_files чтобы в дальнейшем работать с каждым файлом по отдельности
    for filename in [f for f in filenames if f.endswith(".pdf")]:
        all_files.append(os.path.join(filename))
        dir_path = dirpath

def start_parsing(doc):     #функция расчет возвращает значение в виде df
    col_1 = search_for_text_in_block(doc, text='Роль субъекта:') #заполнение первого столбца
    col_1[:] = [item for item in col_1 if 'Кредитная линия' not in item]
    col_2 = search_for_text_in_block(doc, text='Общая сумма кредита') #заполнение второго столбца

    col_3 = [] #заполнение третьего столбца
    for page in doc:
        text_from_page = page.get_text()
        text_from_page = text_from_page.replace('\n', ' ')
        if "Максимальное количество дней просрочки с начала" in text_from_page in text_from_page:
            days = search_for_text('Максимальное количество дней просрочки с начала', text_from_page)
            col_3.append(days)

    col_4 = [] #заполнение четвертого столбца
    for page in doc:
        text_from_page = page.get_text()
        text_from_page = text_from_page.replace('\n', ' ')
        if 'Максимальная сумма просроченных взносов с начала' in text_from_page:
            sums = search_for_text('Максимальная сумма просроченных взносов с начала', text_from_page)
            col_4.append(sums)

    # объединение всех 4 столбцов в один dataframe
    df = pd.DataFrame(list(zip(col_1, col_2, col_3, col_4)),
                   columns =['Общая информация', 'Баланс', 'Максимальное количество дней просрочки', 'Максимальная сумма просроченных взносов'])
    df = df.astype({'Максимальное количество дней просрочки': 'int', 'Максимальная сумма просроченных взносов':'int'}, errors='ignore')
    df['Максимальное количество дней просрочки'] = pd.to_numeric(df['Максимальное количество дней просрочки'], errors='coerce').fillna('нет данных', downcast='infer')
    df['Максимальная сумма просроченных взносов'] = pd.to_numeric(df['Максимальная сумма просроченных взносов'], errors='coerce').fillna('нет данных', downcast='infer')
    return df

#сохраняем dataframe в эксель (с названием каждого файла в отдельном листе). для этого вызываем функцию start_parsing и туда по очереди передаем название каждого файла
writer = ExcelWriter('output.xlsx')
for filename in all_files:
    fn = filename
    filename = dir_path + '\\' + filename
    doc = fitz.Document(filename)
    df = start_parsing(doc)
    sheet_name = str(fn).replace('.pdf', '').replace('\\', '')
    sheet_name = sheet_name[-30:]
    df.to_excel(writer, f'{sheet_name}')
writer.close()