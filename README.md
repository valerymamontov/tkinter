## Автоматический перенос данных их файла MS Excel в файл MS Word
Ежедневно готовится большое количество справок, в которые вставляются данные из Excel-отчётов. Приходится много "копипастить".
В отчётах содержится название организации, дата образования, ИНН, виды деятельности, учредители и множество других сведений.  
Нужная информация из отчёта переносится в новый документ в формате MS Word. В итоге получается краткая справка об организации.  

Используя встроенную в Python библиотеку Tkinter, написал небольшую программу с графическим интерфейсом.
<p align="center">
<img src="https://github.com/valerymamontov/screenshots/blob/master/tkinter.gif">
</p>

Чтобы открыть Excel-файл и считать его содержимое, я использовал [openpyxl][1].  
Чтобы передать данные в документ формата Word, использовал [docxtpl][2]. На использование библиотеки docxtpl натолкнула [статья на хабре][3].  
Чтобы программу можно было запустить под ОС Windows, файл program.py я конвертировал в program.exe при помощи [pyinstaller][4].  

## Содержание репозитория:

"Report by Contur-Focus.xlsx" - Отчёт  
"Pattern.docx" - Шаблон  
"MyReport.docx" - Итоговая справка  


program.py - скрипт
requirements.txt - список зависимостей  
program.exe - исполняемый файл

[1]: https://openpyxl.readthedocs.io/en/stable/
[2]: https://docxtpl.readthedocs.io/en/latest/
[3]: https://habr.com/ru/post/456534/
[4]: https://pypi.org/project/PyInstaller/
