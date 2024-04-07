import pandas as pd
import xlrd
from openpyxl import Workbook


def print_text(text, theme=''):
    print(theme)
    print('________________________________________________________')
    print(text)
    print('________________________________________________________\n')


file = 'reestr.xls'
file_modified = 'reestr_modified.xlsx'
workbook = xlrd.open_workbook(file)
sheet = workbook.sheet_by_index(0)
wb = Workbook()
ws = wb.active

# Копирование данных из файла .xls в новую книгу .xlsx
for row in range(sheet.nrows):
    for col in range(sheet.ncols):
        ws.cell(row=row + 1, column=col + 1, value=sheet.cell_value(row, col))

# Удаление первых трех строк
for _ in range(3):
    ws.delete_rows(1)
print('Файл изменен')
wb.save(file_modified)
print('Файл сохранен')
df = pd.read_excel(file_modified)
replaced_fields = {
    'Исправно': ['Сезонное хранение', 'Консервация', 'Командировка'],
    'Неисправно': ['Авария/ДТП', 'Государственный техосмотр', 'Под списание', 'Техобслуживание и ремонт']
}
regions = ["ВАО", "ЗАО", "ЗелАО", "САО", "СВАО", "СЗАО", "ТинАО", "ЦАО", "ЮАО", "ЮВАО", "ЮЗАО"]
orgs_for_dzkh = ['АО «Мосгаз»', 'АО «Мослифт»', 'АО «ОЭК»', 'ГБУ «Озеленение»', 'ГУП "Моссвет"', 'КП МЭД',
                 'ПАО "МОЭСК"', 'ПАО «МОЭК»']
orgs_for_tinao = ['Администрация поселения Внуковское', 'МБУ «Городское благоустройство»', 'МБУ «ДХБ»',
                  'МБУ поселения Щаповское «КБС и ЖКХ»', 'Администрация поселения Марушкинское',
                  'Администрация поселения Мосрентген', 'МБУ «Жилищник поселения Марушкинское»',
                  'МБУ «СЕЗ «Мосрентген»']

df['Состояние'] = df['Состояние'].replace(replaced_fields.get('Неисправно'), 'Неисправно')
df['Состояние'] = df['Состояние'].replace(replaced_fields.get('Исправно'), 'Исправно')

df.dropna(subset=['Состояние'], inplace=True)

df.loc[df['Округ'].notna() & ~df['Округ'].isin(regions), 'Округ'] = 'ДЖКХ'

df.loc[df['Организация'].isin(orgs_for_dzkh) & df['Округ'].isna(), 'Округ'] = 'ДЖКХ'

df.loc[df['Организация'].isin(orgs_for_tinao) & df['Округ'].isna(), 'Округ'] = 'ТинАО'

df.dropna(subset=['Округ'], inplace=True)

# Данные по ДЖКХ

df_djkh = df[df['Округ'] == 'ДЖКХ']
all_djkh = len(df_djkh)
df_djkh_neispravno = df_djkh[df_djkh['Состояние'] == 'Неисправно']
all_neispravno_djkh = len(df_djkh_neispravno)
print_text(f'Всего: {all_djkh}\nНеисправно: {all_neispravno_djkh}\nПроцент: {all_neispravno_djkh / (all_djkh / 100)}',
           'ДЖКХ')
# Данные по ДЖКХ


# Данные по всему
all_vsego = len(df)
all_neispravno = len(df[df['Состояние'] == 'Неисправно'])
all_ispravno = len(df[df['Состояние'] == 'Исправно'])
print_text(
    f'Всего: {all_vsego}\nИсправно: {all_ispravno} ({all_ispravno / (all_vsego / 100)})\nНеисправно: {all_neispravno} ({all_neispravno / (all_vsego / 100)})',
    'Всего')
# Данные по всему

# Данные по всему кроме ДЖКХ
df_other = df[df['Округ'].isin(regions)]
all_other = len(df_other)
neisparvno_other = len(df_other[df_other['Состояние'] == 'Неисправно'])
print_text(f'Всего: {all_other}\nНеисправно: {neisparvno_other}\nПроцент: {neisparvno_other / (all_other / 100)}',
           "Префектуры")
# Данные по всему кроме ДЖКХ

# Левая диаграмма (ДЖКХ)

all_count_djkh = df_djkh.groupby('Организация').size().reset_index(name='Всего')
df_djkh_neispravno = df_djkh[df_djkh['Состояние'] == 'Неисправно']
count_neispravno = df_djkh_neispravno.groupby('Организация').size().reset_index(name='Неисправно')
result_djkh = pd.merge(all_count_djkh, count_neispravno, on='Организация', how='left')
result_djkh['Неисправно'] = result_djkh['Неисправно'].fillna(0)
result_djkh = result_djkh.sort_values(by='Всего', ascending=True)
result_djkh = result_djkh.sort_values(by='Неисправно', ascending=False)
other = result_djkh.tail(len(result_djkh) - 4)
top_4 = result_djkh.head(4)
sum_other = other.sum().to_frame().T
sum_other['Организация'] = 'Иные'
top_4 = pd.concat([top_4, sum_other], ignore_index=True)
top_4['Процент_неисправных'] = (top_4['Неисправно'] / top_4['Всего']) * 100
print_text(top_4, 'Организации')
# Левая диаграмма (ДЖКХ)


# Правая диаграмма (ДЖКХ)
grouped = df[df['Округ'].isin(regions)].groupby('Округ')
totals = grouped.size().reset_index(name='Всего')
faulty = df[df['Состояние'] == 'Неисправно'].groupby('Округ').size().reset_index(name='Неисправно')
result = pd.merge(totals, faulty, on='Округ', how='left')
result['Неисправно'] = result['Неисправно'].fillna(0)
result['Процент_неисправных'] = (result['Неисправно'] / result['Всего']) * 100

print_text(result, 'Округи')
# Правая диаграмма (ДЖКХ)

# Нижние диаграммы
unique_values = df['Группа техники'].unique()
for group in unique_values:
    df_ubor = df[df['Группа техники'] == group]
    grouped = df_ubor.groupby('Тип')
    totals = grouped.size().reset_index(name='Всего')
    faulty = df_ubor[df_ubor['Состояние'] == 'Неисправно'].groupby('Тип').size().reset_index(name='Неисправно')
    result = pd.merge(totals, faulty, on='Тип', how='left')
    result['Неисправно'] = result['Неисправно'].fillna(0)
    result = result.sort_values(by='Всего', ascending=True)
    result = result.sort_values(by='Неисправно', ascending=False)
    other = result.tail(len(result) - 3)
    top_3 = result.head(3)
    sum_other = other.sum().to_frame().T
    sum_other['Тип'] = 'Иные'
    top_3 = pd.concat([top_3, sum_other], ignore_index=True)
    top_3['Процент_неисправных'] = (top_3['Неисправно'] / top_3['Всего']) * 100
    top_3 = top_3.sort_values(by='Процент_неисправных', ascending=False)
    total_sum = result.sum().to_frame().T
    total_sum['Тип'] = 'Всего'
    total_sum['Процент_неисправных'] = (total_sum['Неисправно'] / total_sum['Всего']) * 100
    final_result = pd.concat([top_3, total_sum], ignore_index=True)
    print_text(final_result, group)
# Нижние диаграммы


df.to_excel('reestrnew.xlsx', index=False)
