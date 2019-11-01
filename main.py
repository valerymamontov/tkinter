from tkinter import *
from tkinter import filedialog as fd

import openpyxl

# import pandas as pd


def cmd_select_file():
    cmd_clear_all()
    file_name = fd.askopenfilename(filetypes=(("Excel files", "*.xls"), ("Excel files", "*.xlsx")))
    f_name["text"] = file_name
    func_add_text()
    # смена цвета кнопок
    btnNEW["bg"] = "lightgreen"
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
        btnNEW["bg"] = "lightgreen"
        btnCLR["bg"] = "lightblue"
        btnGETTXT["bg"] = "lightblue"

        wb = openpyxl.load_workbook(filename=f_name["text"])
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

        # получение номеров строк и текста по этим номерам

        # ---- наименование ----
        legalNameStart = dict_a.get("Полное наименование")
        pos1 = "B" + str(dict_a.get("Полное наименование"))
        pos1C = "C" + str(dict_a.get("Полное наименование"))
        if sheet[pos1C].value != "":
            date = str(sheet[pos1C].value)
            date = date[0:10]
            normDate = date[8:10] + "." + date[5:7] + "." + date[0:4]
            legalNameText = sheet[pos1].value + f"(c {normDate} г.)"
        else:
            legalNameText = sheet[pos1].value
        # ---- наименование ----

        # ---- краткое наименование ----
        legalNameEnd = dict_a.get("Краткое наименование")
        countRowsForPrim = legalNameEnd - legalNameStart
        if countRowsForPrim == 1:
            primText = ""
        else:
            if countRowsForPrim == 2:
                pos2 = "B" + str(legalNameEnd - 1)
                pos2C = "C" + str(legalNameEnd - 1)
                date = str(sheet[pos2C].value)
                date = date[0:10]
                normDate = date[8:10] + "." + date[5:7] + "." + date[0:4]
                primText = "с " + normDate + " г." + " - " + str(sheet[pos2].value)
            else:
                primText = ""
                i = 1
                while i < countRowsForPrim:
                    pos2 = "B" + str(legalNameStart + i)
                    pos2C = "C" + str(legalNameStart + i)
                    date = str(sheet[pos2C].value)
                    date = date[0:10]
                    normDate = date[8:10] + "." + date[5:7] + "." + date[0:4]
                    text = "с " + normDate + " г." + " - " + str(sheet[pos2].value) + "\n"
                    primText += text
                    i = i + 1
        # ---- краткое наименование ----

        registrationDateText = func_get_text("Дата образования")
        innText = func_get_text("ИНН")
        ogrnText = func_get_text("ОГРН")
        legalAddressText = func_get_text("Юр. адрес")
        principalActivityText = func_get_text("Основной вид деятельности")
        statedCapitalSumText = func_get_text("Уставный капитал")

        uchrPosStart = dict_a.get("Учредители")
        uchrPosEnd = dict_a.get("Конечные владельцы")
        if uchrPosEnd < uchrPosStart:
            uchrPosEnd = dict_a.get("Учредители (Росстат)")
            uchrRowsCount = uchrPosEnd - uchrPosStart
        else:
            uchrRowsCount = uchrPosEnd - uchrPosStart
            if uchrRowsCount == 2:
                uchrText = func_get_text("Учредители")
            if uchrRowsCount > 2:
                uchrText = ""
                uchrRowsCount = uchrRowsCount - 1
                i = 1
                while i < uchrRowsCount:
                    pos = "B" + str(uchrRowsCount + i)
                    posC = "C" + str(uchrRowsCount + i)
                    date = str(sheet[posC].value)
                    date = date[0:10]
                    normDate = date[8:10] + "." + date[5:7] + "." + date[0:4]
                    text = "с " + normDate + " г." + " - " + str(sheet[pos].value) + "\n"
                    oldHeadsText += text
                    i = i + 1

        # else:
        #     print(f"start - {uchrPosStart}, end - {uchrPosEnd}")
            # to make empty

        shareholderRegisterText = func_get_text("Держатель реестра акционеров АО")

        # ---- руководитель ----
        text = func_get_text("Генеральный директор")
        posC = "C" + str(dict_a.get("Генеральный директор"))
        if sheet[posC].value != "":
            date = str(sheet[posC].value)
            date = date[0:10]
            normDate = date[8:10] + "." + date[5:7] + "." + date[0:4]
            headsText = f"c {normDate} г. - " + text
        else:
            headsText = text
        # ---- руководитель ----

        # ---- прежние руководители ----
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
            text = "с " + normDate + " г." + " - " + str(sheet[pos].value) + "\n"
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
                text = "с " + normDate + " г." + " - " + str(sheet[pos].value) + "\n"
                oldHeadsText += text
                i = i + 1
        # ---- прежние руководители ----

        legalName.insert(1.0, legalNameText)
        prim.insert(1.0, primText)
        registrationDate.insert(1.0, registrationDateText)
        inn.insert(1.0, innText)
        ogrn.insert(1.0, ogrnText)
        legalAddress.insert(1.0, legalAddressText)
        principalActivity.insert(1.0, principalActivityText)
        statedCapitalSum.insert(1.0, statedCapitalSumText)
        shareholderRegister.insert(1.0, shareholderRegisterText)
        heads.insert(1.0, headsText)
        oldHeads.insert(1.0, oldHeadsText)

        # это вариант через pandas
        # x_file = pd.ExcelFile(file_name)
        # x_sheet = x_file.parse(0)
        # var1 = x_sheet["B16"]


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
    btnNEW["bg"] = "lightgrey"
    btnCLR["bg"] = "lightgrey"
    btnGETTXT["bg"] = "lightgrey"


def func_create_new():
    pass


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
lpx = 40                                    # отступ по X, т.е. слева
hpy = 2                                     # отступ по Y, т.е. сверху/снизу
lblwdth = 52                                # ширина поля
myfont = ("Microsoft Sans Serif", "8")      # шрифт
# ----------------------------------------- # мои настройки ----------

# ======================== кнопки и поля =============================
Label(text="Выберите отчёт из Контур-Фокуса").grid(row=0, column=0, sticky=W, pady=20, padx=10)
Button(text="Выбрать файл", width=45, height=1, bg="lightgrey", command=cmd_select_file).grid(
    row=0, column=1, sticky=W, pady=20, padx=1)

Label(text="Имя файла:").grid(row=1, column=0, sticky=W, pady=2, padx=10)
f_name = Label(width=38, height=1)
f_name.grid(row=1, column=1, sticky=W, pady=2, padx=0)

btnNEW = Button(text="Создать отчёт", width=45, height=1, bg="lightgrey", command=func_create_new)
btnNEW.grid(row=2, columnspan=2, sticky=E, pady=10, padx=1)

Label(text="____________________________________________")\
    .grid(row=3, columnspan=2, column=0, sticky=W, pady=1, padx=10)

Label(text="Наименование:").grid(row=4, column=0, sticky=E, pady=hpy, padx=lpx)
legalName = Text(width=lblwdth, height=2, font=myfont)
legalName.grid(row=4, column=1, sticky=W, pady=hpy, padx=1)

Label(text="Примечание:").grid(row=5, column=0, sticky=E, pady=hpy, padx=lpx)
prim = Text(width=lblwdth, height=4, font=myfont)
prim.grid(row=5, column=1, sticky=W, pady=hpy, padx=1)

Label(text="Дата образования:").grid(row=6, column=0, sticky=E, pady=hpy, padx=lpx)
registrationDate = Text(width=20, height=1, font=myfont)
registrationDate.grid(row=6, column=1, sticky=W, pady=hpy, padx=1)

Label(text="ИНН:").grid(row=7, column=0, sticky=E, pady=hpy, padx=lpx)
inn = Text(width=20, height=1, font=myfont)
inn.grid(row=7, column=1, sticky=W, pady=hpy, padx=1)

Label(text="ОГРН:").grid(row=8, column=0, sticky=E, pady=hpy, padx=lpx)
ogrn = Text(width=20, height=1, font=myfont)
ogrn.grid(row=8, column=1, sticky=W, pady=hpy, padx=1)

Label(text="Адрес:").grid(row=9, column=0, sticky=E, pady=hpy, padx=lpx)
legalAddress = Text(width=lblwdth, height=2, font=myfont)
legalAddress.grid(row=9, column=1, sticky=W, pady=hpy, padx=1)

Label(text="ОКВЭД:").grid(row=10, column=0, sticky=E, pady=hpy, padx=lpx)
principalActivity = Text(width=lblwdth, height=1, font=myfont)
principalActivity.grid(row=10, column=1, sticky=W, pady=hpy, padx=1)

Label(text="Уставной капитал:").grid(row=11, column=0, sticky=E, pady=hpy, padx=lpx)
statedCapitalSum = Text(width=lblwdth, height=1, font=myfont)
statedCapitalSum.grid(row=11, column=1, sticky=W, pady=hpy, padx=1)

Label(text="Учредители:").grid(row=12, column=0, sticky=E, pady=hpy, padx=lpx)
uchr = Text(width=lblwdth, height=2, font=myfont)
uchr.grid(row=12, column=1, sticky=W, pady=hpy, padx=1)

Label(text="Предыдущие собственники:").grid(row=13, column=0, sticky=E, pady=hpy, padx=lpx)
uchrEx = Text(width=lblwdth, height=2, font=myfont)
uchrEx.grid(row=13, column=1, sticky=W, pady=hpy, padx=1)

Label(text="Держатель реестра акционеров:").grid(row=14, column=0, sticky=E, pady=hpy, padx=lpx)
shareholderRegister = Text(width=lblwdth, height=1, font=myfont)
shareholderRegister.grid(row=14, column=1, sticky=W, pady=hpy, padx=1)

Label(text="Руководитель:").grid(row=15, column=0, sticky=E, pady=hpy, padx=lpx)
heads = Text(width=lblwdth, height=2, font=myfont)
heads.grid(row=15, column=1, sticky=W, pady=hpy, padx=1)

Label(text="Прежний руководитель:").grid(row=16, column=0, sticky=E, pady=hpy, padx=lpx)
oldHeads = Text(width=lblwdth, height=2, font=myfont)
oldHeads.grid(row=16, column=1, sticky=W, pady=hpy, padx=1)

Label(text="____________________________________________")\
    .grid(columnspan=2, row=17, column=0, sticky=W, pady=1, padx=10)

btnGETTXT = Button(text="Получить значения", width=31, height=1, bg="lightgrey", command=func_add_text)
btnGETTXT.grid(row=18, column=0, sticky=W, pady=15, padx=10)
btnCLR = Button(text="Очистить все поля", width=45, height=1, bg="lightgrey", command=cmd_clear_all)
btnCLR.grid(row=18, column=1, sticky=E, pady=15, padx=1)
# ======================== кнопки и поля =============================

root.title("Автоматическое создание отчёта")
root.geometry("600x700+300+50")
root.resizable(False, False)
root.mainloop()
