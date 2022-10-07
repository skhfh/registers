import win32com.client
import os
from datetime import date, datetime


# функция возвращающая следующий месяц
def next_month(month):
    month_list = ['Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь', 'Июль',
                  'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь', 'Январь']
    month_index = month_list.index(month)
    month_next = month_list[month_index + 1]
    return month_next


# функция создающая форму для пустых реестров на 10 строк
def form_creation_10(register_number, register_name):
    sh_change = reg_wb.Sheets(f'Реестр {register_number}')
    tem_sh.Range('A1:E7').EntireColumn.Copy(sh_change.Range('A1:E7'))
    tem_sh.Range('A8:E8').Copy(sh_change.Range('A8:E17'))
    tem_sh.Range('A9:E15').Copy(sh_change.Range('A18:E24'))
    sh_change.Range('C18').value = '=СУММ(C8:C17)'

    sh_change.Range('A3').value = f'РЕЕСТР №{register_number}'
    sh_change.Range('A4').value = register_name
    sh_change.Range('A5').value = month_year


# функция проверяет указана ли дата в имени файла, выдает эту дату или error
def check_date(estimated_date):
    global test
    list_format = ['%d.%m.%Y', '%d%m%Y', '%Y%m%d']
    for format in list_format:
        try:
            test = datetime.strptime(estimated_date, format)
        except ValueError:
            test = 'error'
        else:
            break
    return test


# функция заполняющая реестр в экселе
def report_generation(register_number, register_name):
    this_path = os.getcwd()
    files_location = f'{this_path}\Подтверждающие документы\Реестр {register_number}'

    files_name = os.listdir(files_location)
    files_list = []
    files_date = []

    # создается форма для реестра на 200 строк
    sh_change = reg_wb.Sheets(f'Реестр {register_number}')
    tem_sh.Range('A1:E7').EntireColumn.Copy(sh_change.Range('A1:E7'))
    tem_sh.Range('A8:E8').Copy(sh_change.Range('A8:E207'))
    tem_sh.Range('A9:E15').Copy(sh_change.Range('A208:E214'))
    sh_change.Range('C208').value = '=СУММ(C8:C207)'

    # цикл заполняет список с именами файлов без расширения и список с датами изменения файлов
    for i in range(len(files_name)):
        z = files_name[i].rfind('.')
        files_list.append(files_name[i][:z])
        file_time = os.path.getmtime(f'{files_location}\{files_name[i]}')
        file_date = date.fromtimestamp(file_time).strftime('%d.%m.%Y')
        files_date.append(file_date)

    # часть кода, которая изменяет дату в списке на дату из имени файла, и удаляет ее в имени
    shear_list_date = [[-10, None], [-8, None], [None, 10], [None, 8]]
    shear_list_name = [[-11, None], [-9, None], [None, 11], [None, 9]]

    for i in range(len(files_list)):
        for shear in shear_list_date:
            check = check_date(files_list[i][shear[0]:shear[1]])
            if check != 'error':
                check_format = check.strftime('%d.%m.%Y')
                files_date[i] = check_format
                shear_name = shear_list_name[shear_list_date.index(shear)]
                removable_part = files_list[i][shear_name[0]:shear_name[1]]
                files_list[i] = files_list[i].strip(removable_part)

    # собираем сортированную по дате таблицу для экселя
    db = []
    for i in range(len(files_list)):
        db.append([files_list[i], files_date[i]])
    db.sort(key=lambda x: datetime.strptime(x[1], '%d.%m.%Y'))
    for i in range(len(db)):
        db[i].insert(0, i + 1)

    # записываем в файл эксель и удаляем лишние строки

    sh_change.Range('A3').value = f'РЕЕСТР №{register_number}'
    sh_change.Range('A4').value = register_name
    sh_change.Range('A5').value = month_year

    row_not_del = 8
    for row in range(len(db)):
        sh_change.Cells(row + 8, 1).value = db[row][0]
        sh_change.Cells(row + 8, 2).value = db[row][1]
        sh_change.Cells(row + 8, 4).value = db[row][2]
        row_not_del += 1

    sh_change.Range(sh_change.Cells(row_not_del, 1), sh_change.Cells(207, 1)).EntireRow.Delete()


#######################
# выполнение программы....

month_year = input('Введите текущий месяц, в формате "Июнь 2021 г."\n'
                   'и нажмите Enter: ')
print('\nВыполняется, ждите...')

month = month_year.split(' ')[0].capitalize()

# создаем книгу реестров
this_path = os.getcwd()
Excel = win32com.client.Dispatch("Excel.Application")
tem_wb = Excel.Workbooks.Open(fr'{this_path}\template.xlsx')
tem_sh = tem_wb.Sheets(1)

reg_wb = Excel.Workbooks.Add()
reg_sh = reg_wb.Sheets(1)
reg_sh.Name = 'Лишний'

for i in range(28, 12, -1):
    if tem_sh.Cells(i, 7).value != '-':
        sh_new = reg_wb.Sheets.Add()
        sh_new.Name = f'Реестр {i - 12}'

reg_sh.Delete()

# если есть заполняем первые четыре реестра (планы и отчеты)

# заполняем 1й реестр (план на следующий месяц)
try:
    sh_change = reg_wb.Sheets('Реестр 1')
    tem_sh.Range('A1:E15').EntireColumn.Copy(sh_change.Range('A1:E15'))
    sh_change.Range('A3').value = 'РЕЕСТР №1'
    sh_change.Range('A4').value = tem_sh.Range('H13').value
    sh_change.Range('A5').value = month_year
    sh_change.Range('A8').value = '1'
    sh_change.Range('B8').value = f'План работы на {next_month(month)}'
    sh_change.Range('D8').value = datetime.today().strftime('%d.%m.%Y')
except Exception:
    pass

# заполняем 2й реестр (план на следующий квартал)
try:
    sh_change = reg_wb.Sheets('Реестр 2')
    tem_sh.Range('A1:E15').EntireColumn.Copy(sh_change.Range('A1:E15'))
    sh_change.Range('A3').value = 'РЕЕСТР №2'
    sh_change.Range('A4').value = tem_sh.Range('H14').value
    sh_change.Range('A5').value = month_year
    sh_change.Range('A8').value = '1'
    sh_change.Range('B8').value = f'План работы на квартал'
    sh_change.Range('D8').value = datetime.today().strftime('%d.%m.%Y')
except Exception:
    pass

# заполняем 3й реестр (отчет работы за месяц)
try:
    sh_change = reg_wb.Sheets('Реестр 3')
    tem_sh.Range('A1:E15').EntireColumn.Copy(sh_change.Range('A1:E15'))
    sh_change.Range('A3').value = 'РЕЕСТР №3'
    sh_change.Range('A4').value = tem_sh.Range('H15').value
    sh_change.Range('A5').value = month_year
    sh_change.Range('A8').value = '1'
    sh_change.Range('B8').value = f'Отчет работы за {month}'
    sh_change.Range('D8').value = datetime.today().strftime('%d.%m.%Y')
except Exception:
    pass

# заполняем 4й реестр (отчет работы за квартал)
try:
    sh_change = reg_wb.Sheets('Реестр 4')
    tem_sh.Range('A1:E15').EntireColumn.Copy(sh_change.Range('A1:E15'))
    sh_change.Range('A3').value = 'РЕЕСТР №4'
    sh_change.Range('A4').value = tem_sh.Range('H16').value
    sh_change.Range('A5').value = month_year
    sh_change.Range('A8').value = '1'
    sh_change.Range('B8').value = f'Отчет работы за квартал'
    sh_change.Range('D8').value = datetime.today().strftime('%d.%m.%Y')
except Exception:
    pass

# заполняем оставшиеся реестры

reg_num_list = []
for i in range(17, 29):
    if tem_sh.Cells(i, 7).value != '-':
        reg_num = int(tem_sh.Cells(i, 7).value)
        reg_num_name = tem_sh.Cells(i, 8).value
        reg_num_list.append([reg_num, reg_num_name])

for register_number, register_name in reg_num_list:
    try:
        report_generation(register_number, register_name)
    except FileNotFoundError:
        form_creation_10(register_number, register_name)

# сохраняем файл

reg_wb.SaveAs(fr'{this_path}\Реестры.xlsx')
reg_wb.Close()

tem_wb.Save()
tem_wb.Close()
Excel.Quit()

print()
k = input('Готово! \nМожете закрыть программу \nby_skh')
