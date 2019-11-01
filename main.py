from tkinter import *
from tkinter import filedialog as fd

from docxtpl import DocxTemplate

import openpyxl
import os

# import pandas as pd


def cmd_select_file():
    cmd_clear_all()
    file_name = fd.askopenfilename(filetypes=(("Excel files", "*.xls"), ("Excel files", "*.xlsx")))
    fname = f"Имя файла: {os.path.basename(file_name)}"
    dname = f"Директория: {os.path.dirname(file_name)}"
    # dir_name["text"] = dname
    # f_name["text"] = fname
    f_name["text"] = f"Имя файла: {file_name}"
    patrn_name["text"] = ""
    func_add_text()
    # смена цвета кнопок
    # btnNEW["bg"] = "lightgreen"
    btnCLR["bg"] = "lightblue"
    btnGETTXT["bg"] = "lightblue"


def func_check_field():
    list_fields = [
        legalName,
        prim,
        registrationDate,
        inn,
        ogrn,
        legalAddress,
        principalActivity,
        statedCapitalSum,
        uchr,
        uchrEx,
        shareholderRegister,
        heads,
        oldHeads,
    ]
    total_length = 0
    for field in list_fields:
        total_length += len(field.get(1.0, END))
    # если все поля пустые, то total_length = 13
    # каждое пустое поле содержит какой-то символ, поэтому
    # длина каждого поля равна 1, даже если оно пустое
    return total_length


def func_add_text():
    if f_name["text"] and func_check_field() == 13:
        btnNEW["bg"] = "orange"
        btnCLR["bg"] = "lightblue"
        btnGETTXT["bg"] = "lightblue"

        # wb = openpyxl.load_workbook(filename=f"{dir_name['text'][12:]}/{f_name['text'][11:]}")
        wb = openpyxl.load_workbook(filename=f"{f_name['text'][11:]}")
        sheet = wb["Сводка и история"]

        def func_get_text(key):
            if dict_a.get(key):
                pos = "B" + str(dict_a.get(key))
                text = str(sheet[pos].value)
            else:
                text = ""
            return text

        dict_a = {}
        for i in range(2, 200):
            a = "A" + str(i)
            dict_a[sheet[a].value] = i
        # вывод в консоль для проверки
        # for k, v in dict_a.items():
        #     print(f'{k} - {v}')

        # == получение номеров строк и текста по этим номерам: == #
        # прим.: идут в том же порядке, что и поля в окне

        # ---- Наименование ----
        legalNameStart = dict_a.get("Полное наименование")
        pos1 = "B" + str(dict_a.get("Полное наименование"))
        pos1C = "C" + str(dict_a.get("Полное наименование"))
        if sheet[pos1C].value != "":
            date = str(sheet[pos1C].value)
            date = date[0:10]
            normDate = date[8:10] + "." + date[5:7] + "." + date[0:4]
            legalNameText = f"{sheet[pos1].value} ({normDate} г.)"
        else:
            legalNameText = sheet[pos1].value
        # ---- Наименование ----

        # ---- Краткое наименование ----
        legalNameEnd = dict_a.get("Краткое наименование")
        countRowsForPrim = legalNameEnd - legalNameStart
        if countRowsForPrim == 1:
            primText = ""
        else:
            if countRowsForPrim == 2:
                pos = "B" + str(legalNameEnd - 1)
                posC = "C" + str(legalNameEnd - 1)
                date = str(sheet[posC].value)
                date = date[0:10]
                normDate = date[8:10] + "." + date[5:7] + "." + date[0:4]
                primText = f"{str(sheet[pos].value)} ({normDate} г.)"
            else:
                primText = ""
                i = 1
                while i < countRowsForPrim:
                    pos = "B" + str(legalNameStart + i)
                    posC = "C" + str(legalNameStart + i)
                    date = str(sheet[posC].value)
                    date = date[0:10]
                    normDate = date[8:10] + "." + date[5:7] + "." + date[0:4]
                    if i == (countRowsForPrim - 1):
                        text = f"{str(sheet[pos].value)} ({normDate} г.)"
                    else:
                        text = f"{str(sheet[pos].value)} ({normDate} г.)\n"
                    primText += text
                    i = i + 1
        # ---- Краткое наименование ----

        # ---- Дата образования, ИНН, ОГРН и др. ----
        registrationDateText = func_get_text("Дата образования")
        innText = func_get_text("ИНН")
        ogrnText = func_get_text("ОГРН")
        legalAddressText = func_get_text("Юр. адрес")
        principalActivityText = func_get_text("Основной вид деятельности")
        statedCapitalSumText = func_get_text("Уставный капитал")
        # ---- Дата образования, ИНН, ОГРН и др. ----

        # ---- Учредители ----
        uchrPosStart = dict_a.get("Учредители")
        uchrPosEnd = dict_a.get("Конечные владельцы")
        if uchrPosEnd < uchrPosStart:

            # -- вот такое решение --
            # чтобы получить позицию первого вхождения "учредителей"
            dictTemp = {}
            for i in range(2, 200):
                a = "A" + str(i)
                dictTemp[i] = sheet[a].value
            for k, v in dictTemp.items():
                if v == "Учредители":
                    uchrPosStart = k
                    break
            # -- вот такое решение --

            uchrPosEnd = dict_a.get("Учредители (Росстат)")
            uchrRowsCount = uchrPosEnd - uchrPosStart
            if uchrRowsCount == 1:
                pos = "B" + str(uchrPosStart)
                posC = "C" + str(uchrPosStart)
                date = str(sheet[posC].value)
                date = date[0:10]
                normDate = date[8:10] + "." + date[5:7] + "." + date[0:4]
                text = f"{str(sheet[pos].value)} ({normDate} г.)"
                uchrText = text
            else:
                uchrText = ""
                i = 1
                while i <= uchrRowsCount:
                    if i == 1:
                        pos = "B" + str(uchrPosStart)
                        posC = "C" + str(uchrPosStart)
                    else:
                        pos = "B" + str(uchrPosStart + i - 1)
                        posC = "C" + str(uchrPosStart + i - 1)
                    date = str(sheet[posC].value)
                    date = date[0:10]
                    normDate = date[8:10] + "." + date[5:7] + "." + date[0:4]
                    if i == uchrRowsCount:
                        text = f"{str(sheet[pos].value)} ({normDate} г.)"
                    else:
                        text = f"{str(sheet[pos].value)} ({normDate} г.)\n"
                    uchrText += text
                    i = i + 1
        else:
            uchrRowsCount = uchrPosEnd - uchrPosStart
            if uchrRowsCount == 2:
                uchrText = func_get_text("Учредители")
            if uchrRowsCount > 2:
                uchrText = ""
                uchrRowsCount = uchrRowsCount
                i = 1
                while i < uchrRowsCount:
                    if i == 1:
                        pos = "B" + str(uchrPosStart)
                        posC = "C" + str(uchrPosStart)
                    else:
                        pos = "B" + str(uchrPosStart + i - 1)
                        posC = "C" + str(uchrPosStart + i - 1)
                    date = str(sheet[posC].value)
                    date = date[0:10]
                    normDate = date[8:10] + "." + date[5:7] + "." + date[0:4]
                    if i == (uchrRowsCount - 1):
                        text = f"{str(sheet[pos].value)} ({normDate} г.)"
                    else:
                        text = f"{str(sheet[pos].value)} ({normDate} г.)\n"
                    uchrText += text
                    i = i + 1

        # ---- Предыдущие собственники ----
        uchrExText = ""
        # ---- Предыдущие собственники ----

        # ---- Держатель реестра акционеров ----
        shareholderRegisterText = func_get_text("Держатель реестра акционеров АО")
        # ---- Держатель реестра акционеров ----

        # ---- Руководитель ----
        text = func_get_text("Генеральный директор")
        posC = "C" + str(dict_a.get("Генеральный директор"))
        if sheet[posC].value != "":
            date = str(sheet[posC].value)
            date = date[0:10]
            normDate = date[8:10] + "." + date[5:7] + "." + date[0:4]
            headsText = text + f" ({normDate} г.)"
        else:
            headsText = text
        # ---- Руководитель ----

        # ---- Прежние руководители ----
        oldHeadsStart = dict_a.get("Генеральный директор")
        oldHeadsEnd = dict_a.get("Основной вид деятельности")
        oldHeadsCountRows = oldHeadsEnd - oldHeadsStart
        if oldHeadsCountRows == 2:
            oldHeadsText = ""
        if oldHeadsCountRows == 3:
            pos = "B" + str(oldHeadsStart + 1)
            posC = "C" + str(oldHeadsStart + 1)
            date = str(sheet[posC].value)
            date = date[0:10]
            normDate = date[8:10] + "." + date[5:7] + "." + date[0:4]
            text = f"{str(sheet[pos].value)} ({normDate} г.)"
            oldHeadsText = text
        if oldHeadsCountRows > 3:
            oldHeadsText = ""
            oldHeadsCountRows = oldHeadsCountRows - 1
            i = 1
            while i < oldHeadsCountRows:
                pos = "B" + str(oldHeadsStart + i)
                posC = "C" + str(oldHeadsStart + i)
                date = str(sheet[posC].value)
                date = date[0:10]
                normDate = date[8:10] + "." + date[5:7] + "." + date[0:4]
                if i == (oldHeadsCountRows - 1):
                    text = f"{str(sheet[pos].value)} ({normDate} г.)"
                else:
                    text = f"{str(sheet[pos].value)} ({normDate} г.)\n"
                oldHeadsText += text
                i = i + 1
        # ---- Прежние руководители ----

        # == заполнение полей == #
        legalName.insert(1.0, legalNameText)
        prim.insert(1.0, primText)
        registrationDate.insert(1.0, registrationDateText)
        inn.insert(1.0, innText)
        ogrn.insert(1.0, ogrnText)
        legalAddress.insert(1.0, legalAddressText)
        principalActivity.insert(1.0, principalActivityText)
        statedCapitalSum.insert(1.0, statedCapitalSumText)
        uchr.insert(1.0, uchrText)
        uchrEx.insert(1.0, uchrExText)
        shareholderRegister.insert(1.0, shareholderRegisterText)
        heads.insert(1.0, headsText)
        oldHeads.insert(1.0, oldHeadsText)

        # можно обработать содержимое эксель-файла через pandas
        # как-то примерно так: https://python-scripts.com/question/8941

        # mydata = {}
        #         # mydata["legalName"] = legalNameText
        #         # mydata["prim"] = primText
        #         # mydata["registrationDate"] = registrationDateText
        #         # mydata["inn"] = innText
        #         # mydata["ogrn"] = ogrnText
        #         # mydata["legalAddress"] = legalAddressText
        #         # mydata["principalActivity"] = principalActivityText
        #         # mydata["statedCapitalSum"] = statedCapitalSumText
        #         # mydata["uchr"] = uchrText
        #         # mydata["uchrEx"] = uchrExText
        #         # mydata["shareholderRegister"] = shareholderRegisterText
        #         # mydata["heads"] = headsText
        #         # mydata["oldHeads"] = oldHeadsText
        #         # return mydata

        # return legalNameText, primText, registrationDateText, innText, ogrnText, legalAddressText, \
        #        principalActivityText, statedCapitalSumText, uchrText, uchrExText, shareholderRegisterText, \
        #        headsText, oldHeadsText,


def cmd_clear_all():
    legalName.delete(1.0, END)
    prim.delete(1.0, END)
    registrationDate.delete(1.0, END)
    inn.delete(1.0, END)
    ogrn.delete(1.0, END)
    legalAddress.delete(1.0, END)
    principalActivity.delete(1.0, END)
    statedCapitalSum.delete(1.0, END)
    uchr.delete(1.0, END)
    uchrEx.delete(1.0, END)
    shareholderRegister.delete(1.0, END)
    heads.delete(1.0, END)
    oldHeads.delete(1.0, END)

    # после нажатия цвет кнопок изменяется:
    btnNEW["bg"] = "lightgrey"
    btnCLR["bg"] = "lightgrey"
    btnGETTXT["bg"] = "lightgrey"


def cmd_select_pattern():
    file_name = fd.askopenfilename(filetypes=(("Word files", "*.doc"), ("Word files", "*.docx")))
    patrn_name["text"] = f"Шаблон: {file_name}"
    # смена цвета кнопок
    btnNEW["bg"] = "lightgreen"


def func_create_new():
    doc = DocxTemplate(patrn_name["text"][8:])
    context = {}
    context["legalName"] = legalName.get(1.0, END)
    context["prim"] = prim.get(1.0, END)
    context["registrationDate"] = registrationDate.get(1.0, END)
    context["INN"] = inn.get(1.0, END)
    context["OGRN"] = ogrn.get(1.0, END)
    context["legalAddress"] = legalAddress.get(1.0, END)
    context["principalActivity"] = principalActivity.get(1.0, END)
    context["statedCapitalSum"] = statedCapitalSum.get(1.0, END)
    context["uchr"] = uchr.get(1.0, END)
    context["uchrEx"] = uchrEx.get(1.0, END)
    context["shareholderRegister"] = shareholderRegister.get(1.0, END)
    context["heads"] = heads.get(1.0, END)
    context["oldHeads"] = oldHeads.get(1.0, END)
    doc.render(context)
    doc.save("шаблон-final.docx")


root = Tk()

# ======================= верхнее меню ===============================
mainmenu = Menu()
root["menu"] = mainmenu
mainmenu.add_command(label="Выбрать файл", command=cmd_select_file)
mainmenu.add_command(label="Очистить поля", command=cmd_clear_all)
mainmenu.add_command(label="Получить данные", command=func_add_text)
mainmenu.add_command(label="Новый отчёт", command=func_create_new)
mainmenu.add_command(label="Закрыть", command=root.destroy)
# ======================= верхнее меню ===============================

# ----------------------------------------- # мои настройки ----------
lpx = 45                                    # отступ по X, т.е. слева
hpy = 2                                     # отступ по Y, т.е. сверху/снизу
lblwdth = 52                                # ширина поля
myfont = ("Microsoft Sans Serif", "8")      # шрифт
# ----------------------------------------- # мои настройки ----------

# ======================== кнопки и поля =============================
Label(text="Выберите отчёт из Контур-Фокуса").grid(row=0, column=0, sticky=W, pady=10, padx=10)
Button(text="Выбрать файл", width=45, height=1, bg="lightgrey", command=cmd_select_file).grid(
    row=0, column=1, sticky=W, pady=10, padx=1)

# dir_name = Label(height=1)
# dir_name.grid(row=1, columnspan=2, column=0, sticky=W, pady=1, padx=10)
#
# f_name = Label(height=1)
# f_name.grid(row=2, columnspan=2, column=0, sticky=W, pady=1, padx=10)

# Label(text="Имя файла:").grid(row=2, column=0, sticky=W, pady=2, padx=10)
# f_name = Label(width=70, height=1)
# f_name.grid(row=2, columnspan=2, sticky=W, pady=2, padx=0)

f_name = Label(height=1)
f_name.grid(row=1, columnspan=2, column=0, sticky=W, pady=1, padx=10)

Label(text="Выберите шаблон для заполнения").grid(row=2, column=0, sticky=W, pady=10, padx=10)
Button(text="Выбрать шаблон", width=45, height=1, bg="lightgrey", command=cmd_select_pattern).grid(
    row=2, column=1, sticky=W, pady=10, padx=1)

patrn_name = Label(height=1)
patrn_name.grid(row=3, columnspan=2, column=0, sticky=W, pady=1, padx=10)

btnNEW = Button(text="Создать отчёт", width=45, height=1, bg="lightgrey", command=func_create_new)
btnNEW.grid(row=4, columnspan=2, sticky=E, pady=10, padx=1)

Label(text="____________________________________________")\
    .grid(row=5, columnspan=2, column=0, sticky=W, pady=1, padx=10)


Label(text="Наименование:").grid(row=6, column=0, sticky=E, pady=hpy, padx=lpx)
legalName = Text(width=lblwdth, height=2, font=myfont)
legalName.grid(row=6, column=1, sticky=W, pady=hpy, padx=1)

Label(text="Примечание:").grid(row=7, column=0, sticky=E, pady=hpy, padx=lpx)
prim = Text(width=lblwdth, height=4, font=myfont)
prim.grid(row=7, column=1, sticky=W, pady=hpy, padx=1)

Label(text="Дата образования:").grid(row=8, column=0, sticky=E, pady=hpy, padx=lpx)
registrationDate = Text(width=20, height=1, font=myfont)
registrationDate.grid(row=8, column=1, sticky=W, pady=hpy, padx=1)

Label(text="ИНН:").grid(row=9, column=0, sticky=E, pady=hpy, padx=lpx)
inn = Text(width=20, height=1, font=myfont)
inn.grid(row=9, column=1, sticky=W, pady=hpy, padx=1)

Label(text="ОГРН:").grid(row=10, column=0, sticky=E, pady=hpy, padx=lpx)
ogrn = Text(width=20, height=1, font=myfont)
ogrn.grid(row=10, column=1, sticky=W, pady=hpy, padx=1)

Label(text="Адрес:").grid(row=11, column=0, sticky=E, pady=hpy, padx=lpx)
legalAddress = Text(width=lblwdth, height=2, font=myfont)
legalAddress.grid(row=11, column=1, sticky=W, pady=hpy, padx=1)

Label(text="ОКВЭД:").grid(row=12, column=0, sticky=E, pady=hpy, padx=lpx)
principalActivity = Text(width=lblwdth, height=1, font=myfont)
principalActivity.grid(row=12, column=1, sticky=W, pady=hpy, padx=1)

Label(text="Уставной капитал:").grid(row=13, column=0, sticky=E, pady=hpy, padx=lpx)
statedCapitalSum = Text(width=lblwdth, height=1, font=myfont)
statedCapitalSum.grid(row=13, column=1, sticky=W, pady=hpy, padx=1)

Label(text="Учредители:").grid(row=14, column=0, sticky=E, pady=hpy, padx=lpx)
uchr = Text(width=lblwdth, height=2, font=myfont)
uchr.grid(row=14, column=1, sticky=W, pady=hpy, padx=1)

Label(text="Предыдущие собственники:").grid(row=15, column=0, sticky=E, pady=hpy, padx=lpx)
uchrEx = Text(width=lblwdth, height=2, font=myfont)
uchrEx.grid(row=15, column=1, sticky=W, pady=hpy, padx=1)

Label(text="Держатель реестра акционеров:").grid(row=16, column=0, sticky=E, pady=hpy, padx=lpx)
shareholderRegister = Text(width=lblwdth, height=1, font=myfont)
shareholderRegister.grid(row=16, column=1, sticky=W, pady=hpy, padx=1)

Label(text="Руководитель:").grid(row=17, column=0, sticky=E, pady=hpy, padx=lpx)
heads = Text(width=lblwdth, height=2, font=myfont)
heads.grid(row=17, column=1, sticky=W, pady=hpy, padx=1)

Label(text="Прежний руководитель:").grid(row=18, column=0, sticky=E, pady=hpy, padx=lpx)
oldHeads = Text(width=lblwdth, height=2, font=myfont)
oldHeads.grid(row=18, column=1, sticky=W, pady=hpy, padx=1)

Label(text="____________________________________________")\
    .grid(columnspan=2, row=19, column=0, sticky=W, pady=1, padx=10)

btnGETTXT = Button(text="Получить значения", width=31, height=1, bg="lightgrey", command=func_add_text)
btnGETTXT.grid(row=20, column=0, sticky=W, pady=15, padx=10)
btnCLR = Button(text="Очистить все поля", width=45, height=1, bg="lightgrey", command=cmd_clear_all)
btnCLR.grid(row=20, column=1, sticky=E, pady=15, padx=1)
# ======================== кнопки и поля =============================

root.title("Автоматическое создание отчёта")
root.geometry("620x700+300+50")
root.resizable(False, False)
root.mainloop()
