import datetime
import os
import string
import openpyxl
from openpyxl.styles import Alignment
from openpyxl.workbook import Workbook
from openpyxl.styles import Font
import win32com.client
import win32com
import win32com.client as win32

"""pyinstaller --onefile main.py """

number_of_Direction = {
    "Фонд / Фонд(рефинанс)": 0,
    "ЧИ (внутр)": 0
}  # Направления, дальше в них вкладываются стадии
Deal_stades = {
    "Всего поступило": 0,
    "Передано в Аналитический отдел": 0,
    "Вынесены на Кредитный комитет": 0,
    "Одобрено": 0,
    "Выдано": 0,
    "Основная": 0,
    "Доп.выдача": 0

}  # Стадии, подсчитывает в них и записывает с них в ячейки
my_columns = {
    "ID": 0,
    "Вид сделки": 0,
    "Стадия сделки": 0,
    "Направление сделки": 0,
    "Андеррайтер": 0,
    "Дата вынесения на КК": 0,
    "Результат (окончательно)": 0
}  # Названия колонок по которым идёт поиск

percent = " %"
chop = 0
direction_score = 1
bib = 1
other_Direction = 0
sheet = 0
wb = 0
percent_Error = "0 %"
filenames = "report.xlsx"


def File_Converter_xls_to_xlsx():
    global wb
    global sheet
    global filenames

    file = str(os.path.abspath(os.curdir)) + "\\report.xlsx"
    file_old = str(os.path.abspath(os.curdir)) + "\\report.xls"

    excel = win32.gencache.EnsureDispatch('Excel.Application')
    wbo = excel.Workbooks.Open(file_old)

    wbo.SaveAs(file, FileFormat=51)  # FileFormat = 51 is for .xlsx extension
    wbo.Close()  # FileFormat = 56 is for .xls extension
    del wbo
    excel.Application.Quit()

    wb = openpyxl.load_workbook(filename=filenames)
    sheet = wb.active

    Start()


def File_Saved():
    wb.save(filenames)


def File_Format():
    for i in range(1, 31):
        sheet.row_dimensions[i].height = 20

    for c in string.ascii_letters:
        sheet.column_dimensions[c].width = 25
        # print(c)
    sheet.column_dimensions['A'].width = 46

    for r in range(1, 50):

        for w in range(1, 50):
            vino = sheet.cell(row=r, column=w)
            vino.alignment = Alignment(horizontal='center')


def File_List_Creator():
    message = "Отчёт"
    # date = str(datetime.datetime.now())
    # title_for_repot = message + date
    wb.create_sheet(index=1, title=message)


def Convertering():
    for key in number_of_Direction:
        number_of_Direction[key] = dict(Deal_stades)


def Reading_Appraiser():
    nuuumer = 0

    while True:
        nuuumer = nuuumer + 1
        this_sources = sheet.cell(row=nuuumer, column=my_columns["ID"]).value

        global direction_score

        if this_sources != sheet.cell(row=999, column=1).value:
            direction_score = direction_score + 1

        if this_sources == sheet.cell(row=999, column=1).value:
            break


def Automatic_search_of_the_columns():
    num_columns_w = 0

    while True:
        num_columns_w = num_columns_w + 1
        this_columns = sheet.cell(row=1, column=num_columns_w).value

        if this_columns in my_columns:
            my_columns[this_columns] = num_columns_w
            # print("Успешно Найдено: "+this_columns + " - А её номер : = " + str(num_columns_w))

        if this_columns == sheet.cell(row=1, column=105).value:
            break
    Reading_Appraiser()


def Reading_WriterArray():
    print("Metod Reading Writer Array: - Done")

    empty_cell = sheet.cell(row=999, column=1).value
    # Пустая ячейка как понятие истинной пустоты

    colum = 0
    pole = " "
    this_cell = 0

    for i in range(2, direction_score):

        direction_key = 0
        this_cell_Direct = sheet.cell(row=i, column=my_columns[
            "Направление сделки"]).value  # Подсчёт общего количества сделок по направлениям

        """Это разграничения данных и условий для этих данных"""

        if this_cell_Direct == "ФОНД" or this_cell_Direct == "ФОНД (рефинанс)":
            number_of_Direction["Фонд / Фонд(рефинанс)"]["Всего поступило"] += 1
            direction_key = "Фонд / Фонд(рефинанс)"
        elif this_cell_Direct == "ЧИ (внутр)":
            number_of_Direction["ЧИ (внутр)"]["Всего поступило"] += 1
            direction_key = "ЧИ (внутр)"

        # print("Расчёт по направлениям" + str(number_of_Direction))

        colum = int(my_columns["Андеррайтер"])
        this_cell = sheet.cell(row=i, column=colum).value

        if this_cell != empty_cell:
            pole = "Передано в Аналитический отдел"
            number_of_Direction[direction_key][pole] += 1
            # Подсчёт переданных в Аналитический отдел (по заполнению ячейки)
        # print("Расчёт по анал отделу" + str(number_of_Direction))

        colum = int(my_columns["Дата вынесения на КК"])
        this_cell = sheet.cell(row=i, column=colum).value
        if this_cell != empty_cell:
            number_of_Direction[direction_key]["Вынесены на Кредитный комитет"] += 1
            # Подсчёт вынесенных на КК (по заполнению ячейки)
        # print("Расчёт по КК" + str(number_of_Direction))

        colum = int(my_columns["Результат (окончательно)"])
        this_cell = sheet.cell(row=i, column=colum).value
        this_cell = str(this_cell)
        this_cell = this_cell.lower()
        if this_cell == "одобрено" or this_cell == "положительно" or this_cell == "одобрили":
            number_of_Direction[direction_key]["Одобрено"] += 1
        # Всего по слову одобрено
        # print("Расчёт по одобрениям" + str(number_of_Direction))

        colum = int(my_columns["Стадия сделки"])
        this_cell = sheet.cell(row=i, column=colum).value
        this_cell = str(this_cell)
        this_cell = this_cell.lower()
        if this_cell == "обслуживание" or this_cell == "просрочка" or this_cell == "сделка успешно закрыта":
            number_of_Direction[direction_key]["Выдано"] += 1
        # Поиск по слома стадии для поля выдано
        # print("Расчёт по выдачам" + str(number_of_Direction))

        colum = int(my_columns["Вид сделки"])
        this_cell = sheet.cell(row=i, column=colum).value
        this_cell = str(this_cell)
        this_cell = this_cell.lower()
        if this_cell == "основная":
            number_of_Direction[direction_key]["Основная"] += 1
        elif this_cell == "доп. выдача":
            number_of_Direction[direction_key]["Доп.выдача"] += 1
        # Поиск по виду сделки для поля вид


def WriterFile():
    global bib
    sheet.cell(row=1, column=1).value = "Наименование направлений"
    sheet.cell(row=1, column=1).font = Font(bold=True)
    column_num_write = 2

    save_tottal_cell = 0

    save_tottal_cell_odobreno = 0

    for num_stades in Deal_stades:
        bib = bib + 1
        sheet.cell(row=bib, column=1).value = num_stades

    for direction in number_of_Direction:

        num = 1
        sheet.cell(row=1, column=column_num_write).value = direction
        sheet.cell(row=1, column=column_num_write).font = Font(bold=True)

        sheet.cell(row=1, column=column_num_write + 1).value = "% от входящих"
        sheet.cell(row=1, column=column_num_write + 1).font = Font(bold=True)

        sheet.cell(row=1, column=column_num_write + 2).value = "% выдачи из одобренных"
        sheet.cell(row=1, column=column_num_write + 2).font = Font(bold=True)

        for dil in number_of_Direction[direction]:
            num = num + 1

            sheet.cell(row=num, column=column_num_write).value = number_of_Direction[direction][dil]
            if num == 2:
                save_tottal_cell = number_of_Direction[direction][dil]
            if num > 2 and num < 6:
                del_a = int(number_of_Direction[direction][dil])
                del_b = save_tottal_cell

                if del_a > 0 and del_b > 0:
                    drop = del_a / del_b * 100
                    drop = str(drop)
                    drop_t = int(drop.find('.') + 2)
                    drops = drop[0:drop_t]
                    sheet.cell(row=num, column=column_num_write + 1).value = str(drops) + percent
                else:
                    sheet.cell(row=num, column=column_num_write + 1).value = percent_Error
            elif num > 6:
                del_a = int(number_of_Direction[direction][dil])
                del_b = int(number_of_Direction[direction]["Выдано"])

                if del_a > 0 and del_b > 0:
                    drop = del_a / del_b * 100
                    drop = str(drop)
                    drop_t = int(drop.find('.') + 2)
                    drops = drop[0:drop_t]
                    sheet.cell(row=num, column=column_num_write + 1).value = str(drops) + percent
                else:
                    sheet.cell(row=num, column=column_num_write + 1).value = percent_Error

            if dil == "Одобрено":
                save_tottal_cell_odobreno = number_of_Direction[direction][dil]

            if dil == "Выдано":
                del_C = int(number_of_Direction[direction][dil])
                del_V = save_tottal_cell_odobreno

                if del_C > 0 and del_V > 0:
                    drop = del_C / del_V * 100
                    drop = str(drop)
                    drop_t = int(drop.find('.') + 2)
                    drops = drop[0:drop_t]
                    sheet.cell(row=num, column=column_num_write + 2).value = str(drops) + percent
                else:
                    sheet.cell(row=num, column=column_num_write + 2).value = percent_Error

        column_num_write += 3


def Start():
    global sheet
    Automatic_search_of_the_columns()
    Convertering()
    Reading_WriterArray()
    print("Расчёт по направлениям ГОТОВ " + str(number_of_Direction))
    File_List_Creator()
    sheet = wb["Отчёт"]
    File_Format()
    WriterFile()
    File_Saved()


for i in range(1, 2):
    File_Converter_xls_to_xlsx()
