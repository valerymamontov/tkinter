from tkinter import *
from tkinter import filedialog as fd

from docxtpl import DocxTemplate

import openpyxl

import subprocess


# Команда на кнопку "Выбрать файл"
def cmd_select_file():
    cmd_clear_all()
    file["text"] = ""
    file_name = fd.askopenfilename(filetypes=(("Excel files", "*.xls"), ("Excel files", "*.xlsx")))
    if file_name:
        file["text"] = f"Имя файла: {file_name}"
        cmd_add_text()
        # смена цвета кнопок
        btn_5["bg"] = "lightblue"
        btn_6["bg"] = "lightblue"


# Отдельная фукция по проверке полей
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


# Команда на кнопку "Заполнить все поля"
def cmd_add_text():
    if file["text"] and func_check_field() == 13:
        if pattern["text"]:
            btn_4["bg"] = "lightgreen"
        else:
            btn_4["bg"] = "orange"
        btn_5["bg"] = "lightblue"
        btn_6["bg"] = "lightblue"

        wb = openpyxl.load_workbook(filename=f"{file['text'][11:]}")
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
            nameOrg = sheet[pos1].value
            nameOrg = nameOrg.replace("Общество с ограниченной ответственностью", "ООО")
            nameOrg = nameOrg.replace("ОБЩЕСТВО С ОГРАНИЧЕННОЙ ОТВЕТСТВЕННОСТЬЮ", "ООО")
            nameOrg = nameOrg.replace("Товарищество с ограниченной ответственностью", "ТОО")
            nameOrg = nameOrg.replace("ТОВАРИЩЕСТВО С ОГРАНИЧЕННОЙ ОТВЕТСТВЕННОСТЬЮ", "ТОО")
            nameOrg = nameOrg.replace("Открытое акционерное общество", "ОАО")
            nameOrg = nameOrg.replace("ОТКРЫТОЕ АКЦИОНЕРНОЕ ОБЩЕСТВО", "ОАО")
            nameOrg = nameOrg.replace("Закрытое акционерное общество", "ЗАО")
            nameOrg = nameOrg.replace("ЗАКРЫТОЕ АКЦИОНЕРНОЕ ОБЩЕСТВО", "ЗАО")
            nameOrg = nameOrg.replace("Публичное акционерное общество", "ПАО")
            nameOrg = nameOrg.replace("ПУБЛИЧНОЕ АКЦИОНЕРНОЕ ОБЩЕСТВО", "ПАО")
            nameOrg = nameOrg.replace("Акционерное общество", "АО")
            nameOrg = nameOrg.replace("АКЦИОНЕРНОЕ ОБЩЕСТВО", "АО")
            legalNameText = f"{nameOrg} ({normDate} г.)"
        else:
            nameOrg = sheet[pos1].value
            nameOrg = nameOrg.replace("Общество с ограниченной ответственностью", "ООО")
            nameOrg = nameOrg.replace("ОБЩЕСТВО С ОГРАНИЧЕННОЙ ОТВЕТСТВЕННОСТЬЮ", "ООО")
            nameOrg = nameOrg.replace("Товарищество с ограниченной ответственностью", "ТОО")
            nameOrg = nameOrg.replace("ТОВАРИЩЕСТВО С ОГРАНИЧЕННОЙ ОТВЕТСТВЕННОСТЬЮ", "ТОО")
            nameOrg = nameOrg.replace("Открытое акционерное общество", "ОАО")
            nameOrg = nameOrg.replace("ОТКРЫТОЕ АКЦИОНЕРНОЕ ОБЩЕСТВО", "ОАО")
            nameOrg = nameOrg.replace("Закрытое акционерное общество", "ЗАО")
            nameOrg = nameOrg.replace("ЗАКРЫТОЕ АКЦИОНЕРНОЕ ОБЩЕСТВО", "ЗАО")
            nameOrg = nameOrg.replace("Публичное акционерное общество", "ПАО")
            nameOrg = nameOrg.replace("ПУБЛИЧНОЕ АКЦИОНЕРНОЕ ОБЩЕСТВО", "ПАО")
            nameOrg = nameOrg.replace("Акционерное общество", "АО")
            nameOrg = nameOrg.replace("АКЦИОНЕРНОЕ ОБЩЕСТВО", "АО")
            legalNameText = nameOrg
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
                nameOrg = str(sheet[pos].value)
                nameOrg = nameOrg.replace("Общество с ограниченной ответственностью", "ООО")
                nameOrg = nameOrg.replace("ОБЩЕСТВО С ОГРАНИЧЕННОЙ ОТВЕТСТВЕННОСТЬЮ", "ООО")
                nameOrg = nameOrg.replace("Товарищество с ограниченной ответственностью", "ТОО")
                nameOrg = nameOrg.replace("ТОВАРИЩЕСТВО С ОГРАНИЧЕННОЙ ОТВЕТСТВЕННОСТЬЮ", "ТОО")
                nameOrg = nameOrg.replace("Открытое акционерное общество", "ОАО")
                nameOrg = nameOrg.replace("ОТКРЫТОЕ АКЦИОНЕРНОЕ ОБЩЕСТВО", "ОАО")
                nameOrg = nameOrg.replace("Закрытое акционерное общество", "ЗАО")
                nameOrg = nameOrg.replace("ЗАКРЫТОЕ АКЦИОНЕРНОЕ ОБЩЕСТВО", "ЗАО")
                nameOrg = nameOrg.replace("Публичное акционерное общество", "ПАО")
                nameOrg = nameOrg.replace("ПУБЛИЧНОЕ АКЦИОНЕРНОЕ ОБЩЕСТВО", "ПАО")
                nameOrg = nameOrg.replace("Акционерное общество", "АО")
                nameOrg = nameOrg.replace("АКЦИОНЕРНОЕ ОБЩЕСТВО", "АО")
                primText = f"{nameOrg} ({normDate} г.)"
            else:
                primText = ""
                i = 1
                while i < countRowsForPrim:
                    pos = "B" + str(legalNameStart + i)
                    posC = "C" + str(legalNameStart + i)
                    date = str(sheet[posC].value)
                    date = date[0:10]
                    normDate = date[8:10] + "." + date[5:7] + "." + date[0:4]
                    nameOrg = str(sheet[pos].value)
                    nameOrg = nameOrg.replace("Общество с ограниченной ответственностью", "ООО")
                    nameOrg = nameOrg.replace("ОБЩЕСТВО С ОГРАНИЧЕННОЙ ОТВЕТСТВЕННОСТЬЮ", "ООО")
                    nameOrg = nameOrg.replace("Товарищество с ограниченной ответственностью", "ТОО")
                    nameOrg = nameOrg.replace("ТОВАРИЩЕСТВО С ОГРАНИЧЕННОЙ ОТВЕТСТВЕННОСТЬЮ", "ТОО")
                    nameOrg = nameOrg.replace("Открытое акционерное общество", "ОАО")
                    nameOrg = nameOrg.replace("ОТКРЫТОЕ АКЦИОНЕРНОЕ ОБЩЕСТВО", "ОАО")
                    nameOrg = nameOrg.replace("Закрытое акционерное общество", "ЗАО")
                    nameOrg = nameOrg.replace("ЗАКРЫТОЕ АКЦИОНЕРНОЕ ОБЩЕСТВО", "ЗАО")
                    nameOrg = nameOrg.replace("Публичное акционерное общество", "ПАО")
                    nameOrg = nameOrg.replace("ПУБЛИЧНОЕ АКЦИОНЕРНОЕ ОБЩЕСТВО", "ПАО")
                    nameOrg = nameOrg.replace("Акционерное общество", "АО")
                    nameOrg = nameOrg.replace("АКЦИОНЕРНОЕ ОБЩЕСТВО", "АО")
                    if i == (countRowsForPrim - 1):
                        text = f"{nameOrg} ({normDate} г.)"
                    else:
                        text = f"{nameOrg} ({normDate} г.)\n"
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
        if dict_a.get("Учредители") is None:
            uchrPosStart = dict_a.get("Учредители и участники")
            uchrPosEnd = dict_a.get("Конечные владельцы")
            if uchrPosEnd < uchrPosStart:

                # -- вот такое решение --
                # чтобы получить позицию первого вхождения "учредителей"
                dictTemp = {}
                for i in range(2, 200):
                    a = "A" + str(i)
                    dictTemp[i] = sheet[a].value
                for k, v in dictTemp.items():
                    if v == "Учредители и участники":
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
                    uchrText = func_get_text("Учредители и участники")
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
        else:
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


# Команда на кнопку "Очистить все поля"
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
    btn_4["bg"] = "lightgrey"
    btn_5["bg"] = "lightgrey"
    btn_6["bg"] = "lightgrey"


# Команда на кнопку "Выбрать файл"
def cmd_select_pattern():
    pattern["text"] = ""
    report["text"] = ""
    file_name = fd.askopenfilename(filetypes=(("Word files", "*.doc"), ("Word files", "*.docx")))
    if file_name:
        pattern["text"] = f"Шаблон: {file_name}"
        if file["text"]:
            # смена цвета кнопок
            btn_3["bg"] = "lightgrey"
            btn_4["bg"] = "lightgreen"


# Комнада на кнопку "Создать отчёт"
def cmd_create_new():
    if file["text"]:
        if pattern["text"]:
            doc = DocxTemplate(pattern["text"][8:])
            if doc != "":
                report["text"] = ""
                report_name = fd.asksaveasfilename(defaultextension="*.*", filetypes=(
                    ("Документ Word", "*.docx"),
                    ("Документ Word 97-2003", "*.doc"),
                    ("Текст в формате RTF", "*.rtf"),
                ))
                context = dict()
                context["legalName"] = legalName.get(1.0, END)
                context["prim"] = prim.get(1.0, END).replace("\n", "<w:br/>")
                context["registrationDate"] = registrationDate.get(1.0, END)
                context["INN"] = inn.get(1.0, END)
                context["OGRN"] = ogrn.get(1.0, END)
                context["legalAddress"] = legalAddress.get(1.0, END)
                context["principalActivity"] = principalActivity.get(1.0, END)
                context["statedCapitalSum"] = statedCapitalSum.get(1.0, END)
                context["uchr"] = uchr.get(1.0, END).replace("\n", "<w:br/>")
                context["uchrEx"] = uchrEx.get(1.0, END)
                context["shareholderRegister"] = shareholderRegister.get(1.0, END)
                context["heads"] = heads.get(1.0, END)
                context["oldHeads"] = oldHeads.get(1.0, END).replace("\n", "<w:br/>")
                doc.render(context)
                if report_name:
                    report["text"] = f"Отчёт: {report_name}"
                    doc.save(report_name)
                    # смена цвета кнопок
                    btn_3["bg"] = "lightgreen"
                    btn_4["bg"] = "lightgrey"


# Комнада на кнопку "Посмотреть отчёт"
def cmd_open_report():
    if report["text"]:
        report_name = report["text"][7:]
        subprocess.Popen(('start', report_name), shell=True)




root = Tk()
# ======================= верхнее меню ===============================
main_menu = Menu(root)
root["menu"] = main_menu
file_m = Menu(main_menu, tearoff=0)
fields_m = Menu(main_menu, tearoff=0)
file_m.add_command(label="Выбрать файл", command=cmd_select_file)
file_m.add_command(label="Выбрать шаблон", command=cmd_select_pattern)
fields_m.add_command(label="Очистить все поля", command=cmd_clear_all)
fields_m.add_command(label="Заполнить все поля", command=cmd_add_text)
main_menu.add_cascade(label="Файл", menu=file_m)
main_menu.add_cascade(label="Документ", menu=fields_m)
main_menu.add_command(label="Создать отчёт", command=cmd_create_new)
main_menu.add_command(label="Посмотреть отчёт", command=cmd_open_report)
main_menu.add_command(label="Закрыть", command=root.destroy)
# ======================= верхнее меню ===============================

# ----------------------------------------- # мои настройки ----------
lpx = 45                                    # отступ по X, т.е. слева
hpy = 2                                     # отступ по Y, т.е. сверху/снизу
lblwdth = 52                                # ширина поля
myfont = ("Microsoft Sans Serif", "8")      # шрифт
# ----------------------------------------- # мои настройки ----------

# ======================== кнопки и поля =============================
lbl_1 = Label(text="Выберите отчёт из Контур-Фокуса")
lbl_1.grid(row=0, column=0, sticky=W, pady=10, padx=10)
btn_1 = Button(text="Выбрать файл", width=45, height=1, bg="lightgrey", command=cmd_select_file)
btn_1.grid(row=0, column=1, sticky=W, pady=10, padx=1)

file = Label(height=1)
file.grid(row=1, columnspan=2, column=0, sticky=W, pady=1, padx=10)

lbl_2 = Label(text="Выберите шаблон для заполнения")
lbl_2.grid(row=2, column=0, sticky=W, pady=10, padx=10)
btn_2 = Button(text="Выбрать шаблон", width=45, height=1, bg="lightgrey", command=cmd_select_pattern)
btn_2.grid(row=2, column=1, sticky=W, pady=10, padx=1)

pattern = Label(height=1)
pattern.grid(row=3, columnspan=2, column=0, sticky=W, pady=1, padx=10)

btn_3 = Button(text="Посмотреть отчёт", width=31, bg="lightgrey", command=cmd_open_report)
btn_3.grid(row=4, column=0, sticky=W, pady=1, padx=10)
btn_4 = Button(text="Создать отчёт", bg="lightgrey", command=cmd_create_new)
btn_4.grid(row=4, column=1, sticky=W+E, pady=1, padx=1)

report = Label(height=1)
report.grid(row=5, columnspan=2, column=0, sticky=W, pady=1, padx=10)

Label(text="____________________________________________")\
    .grid(row=6, columnspan=2, column=0, sticky=W, pady=1, padx=10)


Label(text="Наименование:").grid(row=7, column=0, sticky=E, pady=hpy, padx=lpx)
legalName = Text(width=lblwdth, height=2, font=myfont)
legalName.grid(row=7, column=1, sticky=W, pady=hpy, padx=1)

Label(text="Примечание:").grid(row=8, column=0, sticky=E, pady=hpy, padx=lpx)
prim = Text(width=lblwdth, height=3, font=myfont)
prim.grid(row=8, column=1, sticky=W, pady=hpy, padx=1)

Label(text="Дата образования:").grid(row=9, column=0, sticky=E, pady=hpy, padx=lpx)
registrationDate = Text(width=20, height=1, font=myfont)
registrationDate.grid(row=9, column=1, sticky=W, pady=hpy, padx=1)

Label(text="ИНН:").grid(row=10, column=0, sticky=E, pady=hpy, padx=lpx)
inn = Text(width=20, height=1, font=myfont)
inn.grid(row=10, column=1, sticky=W, pady=hpy, padx=1)

Label(text="ОГРН:").grid(row=11, column=0, sticky=E, pady=hpy, padx=lpx)
ogrn = Text(width=20, height=1, font=myfont)
ogrn.grid(row=11, column=1, sticky=W, pady=hpy, padx=1)

Label(text="Адрес:").grid(row=12, column=0, sticky=E, pady=hpy, padx=lpx)
legalAddress = Text(width=lblwdth, height=2, font=myfont)
legalAddress.grid(row=12, column=1, sticky=W, pady=hpy, padx=1)

Label(text="ОКВЭД:").grid(row=13, column=0, sticky=E, pady=hpy, padx=lpx)
principalActivity = Text(width=lblwdth, height=1, font=myfont)
principalActivity.grid(row=13, column=1, sticky=W, pady=hpy, padx=1)

Label(text="Уставной капитал:").grid(row=14, column=0, sticky=E, pady=hpy, padx=lpx)
statedCapitalSum = Text(width=lblwdth, height=1, font=myfont)
statedCapitalSum.grid(row=14, column=1, sticky=W, pady=hpy, padx=1)

Label(text="Учредители и участники:").grid(row=15, column=0, sticky=E, pady=hpy, padx=lpx)
uchr = Text(width=lblwdth, height=2, font=myfont)
uchr.grid(row=15, column=1, sticky=W, pady=hpy, padx=1)

Label(text="Предыдущие собственники:").grid(row=16, column=0, sticky=E, pady=hpy, padx=lpx)
uchrEx = Text(width=lblwdth, height=2, font=myfont)
uchrEx.grid(row=16, column=1, sticky=W, pady=hpy, padx=1)

Label(text="Держатель реестра акционеров:").grid(row=17, column=0, sticky=E, pady=hpy, padx=lpx)
shareholderRegister = Text(width=lblwdth, height=1, font=myfont)
shareholderRegister.grid(row=17, column=1, sticky=W, pady=hpy, padx=1)

Label(text="Руководитель:").grid(row=18, column=0, sticky=E, pady=hpy, padx=lpx)
heads = Text(width=lblwdth, height=2, font=myfont)
heads.grid(row=18, column=1, sticky=W, pady=hpy, padx=1)

Label(text="Прежний руководитель:").grid(row=19, column=0, sticky=E, pady=hpy, padx=lpx)
oldHeads = Text(width=lblwdth, height=2, font=myfont)
oldHeads.grid(row=19, column=1, sticky=W, pady=hpy, padx=1)

Label(text="____________________________________________")\
    .grid(columnspan=2, row=20, column=0, sticky=W, pady=1, padx=10)

btn_5 = Button(text="Заполнить все поля", width=31, height=1, bg="lightgrey", command=cmd_add_text)
btn_5.grid(row=21, column=0, sticky=W, pady=15, padx=10)
btn_6 = Button(text="Очистить все поля", width=45, height=1, bg="lightgrey", command=cmd_clear_all)
btn_6.grid(row=21, column=1, sticky=E, pady=15, padx=1)
# ======================== кнопки и поля =============================

root.title("Автоматическое создание отчёта")
root.geometry("620x700+300+20")
root.resizable(False, False)
root.mainloop()
