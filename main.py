import pandas as pd
import xlrd
from openpyxl import Workbook
from python_pptx_text_replacer import TextReplacer
import locale
from datetime import datetime, timedelta
locale.setlocale(locale.LC_ALL, 'ru_RU.UTF-8')


def fint(x):
    return locale._format('%d', x, grouping=True)


def print_text(text, theme=''):
    print(theme)
    print('________________________________________________________')
    print(text)
    print('________________________________________________________\n')


replacer = TextReplacer("../../Downloads/Telegram Desktop/prez.pptx", slides='',
                        tables=True, charts=True, textframes=True)
file = 'Реестр_транспортных_средств_02_12_2024.xls'
file_1 = file.replace(' ', '_').replace('.', '_').replace('_xls', '.xls')
datereestr = '.'.join(file_1.split('.')[0].split('_')[3:])
date_obj = datetime.strptime(datereestr, "%d.%m.%Y")
old_date_str = date_obj.strftime("%d.%m.%Y")
date_obj += timedelta(days=1)
new_date_str = date_obj.strftime("%d.%m.%Y")
replacer.replace_text([('*dateuid1*', f'{old_date_str}'), ('*dateuid2*', f'{new_date_str}')])

file_modified = 'abcreestr_modified.xlsx'
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
                 'ПАО "МОЭСК"', 'ПАО «МОЭК»', 'МВК']
orgs_for_tinao = ['Администрация поселения Внуковское', 'МБУ «Городское благоустройство»', 'МБУ «ДХБ»',
                  'МБУ поселения Щаповское «КБС и ЖКХ»', 'Администрация поселения Марушкинское',
                  'Администрация поселения Мосрентген', 'МБУ «Жилищник поселения Марушкинское»',
                  'МБУ «СЕЗ «Мосрентген»']

symbolbygroup = {
    'Уборочная техника': 'u',
    'Сопутствующая техника': 's',
    'Дорожно-строительная техника': 'd',
    'Вывоз отходов': 'v',
    'Прочие': 'p'
}

orgs_replace = {
    'АвД': 'Автомобильные дороги',
    'ГУП «ЭКОТЕХПРОМ»': 'Экотехпром',
    'ГБУ «ЭВАЖД»': 'ЭВАЖД'
}

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
djkh_percent = int(round(all_neispravno_djkh / (all_djkh / 100), 0))
print_text(f'Всего: {all_djkh}\nНеисправно: {all_neispravno_djkh}\nПроцент: {djkh_percent}',
           'ДЖКХ')
replacer.replace_text(
    [('*djkhall*', f'{fint(all_djkh)}'), ('*djkh_not*', f'{fint(all_neispravno_djkh)}'),
     ('*djkh_percent*', f'{djkh_percent}')])
# Данные по ДЖКХ


# Данные по всему
all_vsego = len(df)
all_neispravno = len(df[df['Состояние'] == 'Неисправно'])
all_ispravno = len(df[df['Состояние'] == 'Исправно'])
all_ispr_percent = int(round(all_ispravno / (all_vsego / 100), 0))
all_not_percent = int(round(all_neispravno / (all_vsego / 100), 0))
print_text(
    f'Всего: {all_vsego}\nИсправно: {all_ispravno} ({all_ispr_percent})\nНеисправно: {all_neispravno} ({all_not_percent})',
    'Всего')
replacer.replace_text(
    [('*all*', f'{fint(all_vsego)} '), ('*all_ispr*', f'{fint(all_ispravno)}'),
     ('*all_ispr_percent*', f'{all_ispr_percent}'), ('*all_not*', f'{fint(all_neispravno)}'),
     ('*all_not_percent*', f'{all_not_percent}')])
# Данные по всему

# Данные по всему кроме ДЖКХ
df_other = df[df['Округ'].isin(regions)]
all_other = len(df_other)
neisparvno_other = len(df_other[df_other['Состояние'] == 'Неисправно'])
other_percent = int(round(neisparvno_other / (all_other / 100), 0))
print_text(f'Всего: {all_other}\nНеисправно: {neisparvno_other}\nПроцент: {neisparvno_other / (all_other / 100)}',
           "Префектуры")
replacer.replace_text([('*prefect_all*', f'{fint(all_other)}'), ('*prefect_not*', f'{fint(neisparvno_other)}'),
                       ('*prefect_percent*', f'{other_percent}')])
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
top_4 = top_4.sort_values(by='Всего', ascending=False)
sum_other = other.sum().to_frame().T
sum_other['Организация'] = 'Иные'
top_4 = pd.concat([top_4, sum_other], ignore_index=True)
top_4['Процент_неисправных'] = (top_4['Неисправно'] / top_4['Всего']) * 100
print_text(top_4, 'Организации')
org_lst = []
for i in range(5):
    new_name = orgs_replace.get(top_4.iloc[i, 0], top_4.iloc[i, 0])
    org_lst.append((f'*org_name{i + 1}*', f'{new_name}'))
    org_lst.append((f'*org_{i + 1}*', f'{fint(int(top_4.iloc[i, 1]))}'))
    org_lst.append((f'*org_not{i + 1}*', f'{fint(int(top_4.iloc[i, 2]))}'))
    org_lst.append((f'*org_not_p_{i + 1}*', f'{round(top_4.iloc[i, 3], 2):.2f}'.replace('.', ',')))

print(org_lst)

replacer.replace_text(org_lst)

top_2 = top_4.sort_values(by='Процент_неисправных', ascending=False)
new_name_1 = orgs_replace.get(top_2.iloc[0, 0], top_2.iloc[0, 0])
new_name_2 = orgs_replace.get(top_2.iloc[1, 0], top_2.iloc[1, 0])
replacer.replace_text(
    [('*p_top1_name*', f'{new_name_1}'), ('*p_top1_perc*', f'{round(top_2.iloc[0, 3], 2):.2f}%'.replace('.', ',')),
     ('*p_top2_name*', f'{new_name_2}'),
     ('*p_top2_perc*', f'{round(top_2.iloc[1, 3], 2):.2f}%'.replace('.', ','))])

# Левая диаграмма (ДЖКХ)


# Правая диаграмма (ДЖКХ)
grouped = df[df['Округ'].isin(regions)].groupby('Округ')
totals = grouped.size().reset_index(name='Всего')
faulty = df[df['Состояние'] == 'Неисправно'].groupby('Округ').size().reset_index(name='Неисправно')
result = pd.merge(totals, faulty, on='Округ', how='left')
result['Неисправно'] = result['Неисправно'].fillna(0)
result['Процент_неисправных'] = (result['Неисправно'] / result['Всего']) * 100
region_list = []
for i in range(11):
    region_i = str(result.iloc[i, 0]).upper()
    region_list.append((f'*{region_i}*', f'{fint(result.iloc[i, 1])}'))
    region_list.append((f'*{region_i}_n*', f'{fint(result.iloc[i, 2])}'))
    region_list.append((f'*{region_i}_p*', f'{round(result.iloc[i, 3], 2):.2f}'.replace('.', ',')))
print(region_list)
replacer.replace_text(region_list)

top_1 = result.sort_values(by='Процент_неисправных', ascending=False)
replacer.replace_text(
    [('*prfct1*', f'{top_1.iloc[0, 0]}'), ('*prfctp*', f'{round(top_1.iloc[0, 3], 2):.2f}%'.replace('.', ','))])
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
    top_3['Процент_неисправных'] = (top_3['Неисправно'] / top_3['Всего']) * 100
    top_3 = top_3.sort_values(by='Процент_неисправных', ascending=False)
    sum_other = other.sum().to_frame().T
    sum_other['Тип'] = 'Иные'
    top_3 = pd.concat([top_3, sum_other], ignore_index=True)
    top_3['Процент_неисправных'] = (top_3['Неисправно'] / top_3['Всего']) * 100
    total_sum = result.sum().to_frame().T
    total_sum['Тип'] = 'Всего'
    total_sum['Процент_неисправных'] = (total_sum['Неисправно'] / total_sum['Всего']) * 100
    final_result = pd.concat([top_3, total_sum], ignore_index=True)
    grp_replace = []
    ssymbol = symbolbygroup.get(group)
    for i in range(5):
        grp_replace.append((f'{ssymbol}name_{i + 1}', f'{final_result.iloc[i, 0]}'))
        grp_replace.append((f'{ssymbol}all_{i + 1}', f'{fint(final_result.iloc[i, 1])}'))
        grp_replace.append((f'{ssymbol}not_{i + 1}', f'{fint(final_result.iloc[i, 2])}'))
        grp_replace.append((f'{ssymbol}p{i + 1}', f'{int(round(final_result.iloc[i, 3], 0))}%'))
    print_text(final_result, group)
    print(grp_replace)
    replacer.replace_text(grp_replace)
# Нижние диаграммы


df.to_excel('abcnew.xlsx', index=False)
replacer.write_presentation_to_file("changed.pptx")
