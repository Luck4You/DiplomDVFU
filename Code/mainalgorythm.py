import workwithdocs
import workwithexcel
import pandas as pd
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx import Document
import re

def SumOfZe(dataframe, index_of_disc):
    sum_of_ze = 0
    for index, row in dataframe.iterrows():
        if index_of_disc[:3] == str(dataframe.at[index, 'DISCINDEX'])[:3]:
            sum_of_ze += re.search(r'\((\d+)\)', str(row['DISCHOURS']))
    return sum_of_ze

def MakeOne(df_student,file_save, shablon_path):
    # Открываем документ Word по указанному пути
    doc = Document(shablon_path)

    # Инициализация переменных
    index_Table = 0
    DiscType_now = "1"
    leftTable = None
    rightTable = None
    mainTable = None

    # Создаем пустой DataFrame для хранения информации о курсовых проектах и работах
    KursJob = pd.DataFrame(columns=['TYPECONTROL', 'DISCNAME', 'DISCHOURS', 'OCENKA'])

    # Поиск таблиц в документе Word
    for table in doc.tables:
        if table.cell(0, 0).text == "TABL1TABL":
            leftTable = table
        if table.cell(0, 0).text == "TABL2TABL":
            rightTable = table

    mainTable = leftTable

    # Обработка каждой строки в DataFrame df_student
    for index, row in df_student.iterrows():
        # Проверяем тип контроля, если это курсовой проект или курсовая работа, добавляем в KursJob
        if row['TYPECONTROL'] == 'Курсовой проект' or row['TYPECONTROL'] == 'Курсовая работа':
            new_row = {'TYPECONTROL': row['TYPECONTROL'],
                    'DISCNAME': row['DISCNAME'],
                    'DISCHOURS': row['DISCHOURS'],
                    'OCENKA': row['OCENKA']}
            KursJob.loc[len(KursJob.index)] = new_row
        else:
            # Логика обработки других типов дисциплин в зависимости от их индекса
            # Проверяем, не изменился ли тип дисциплины с предыдущей итерации
            if not DiscType_now == str(row['DISCINDEX'])[1]:
                # Проверяем, помещается ли новая информация в текущую таблицу или нужно переключиться на другую
                if index_Table + 1 > len(mainTable.rows) - 1:
                    mainTable = rightTable
                    index_Table = 0
                
                # Объединяем ячейки таблицы, если необходимо
                workwithdocs.MergeTable(mainTable, index_Table, 1)
                
                Sum_Of_Ze = 0
                # Определяем тип дисциплины и обновляем текущий тип
                match(str(row['DISCINDEX'])[1]):
                    case "2":
                        DiscType = "Практики\nв том числе:"
                        DiscType_now = "2"
                        Sum_Of_Ze = SumOfZe(df_student,"Б2")
                    case "3":
                        DiscType = "Государственная итоговая аттестация\nв том числе:"
                        DiscType_now = "3"
                        Sum_Of_Ze = SumOfZe(df_student,"Б3")
                    case "Т":
                        # Циклически обрабатываем данные из DataFrame KursJob для типа "Т"
                        for jindex, row1 in KursJob.iterrows():
                            discname = f"{row1['TYPECONTROL']}, {row1['DISCNAME']}"
                            # Объединяем ячейки таблицы при необходимости
                            if len(discname) >= 55 and index_Table + (len(discname) // 55) <= len(mainTable.rows) - 1:
                                workwithdocs.MergeTable(mainTable, index_Table, (len(discname) // 55))
                                index_Table = index_Table + (len(discname) // 55)
                            elif index_Table + (len(discname) // 55) > len(mainTable.rows) - 1:
                                mainTable = rightTable
                                index_Table = 0
                            
                            # Вставляем информацию в таблицу из DataFrame KursJob
                            workwithdocs.Insert_Str_from(mainTable, index_Table, discname, "", str(row1['OCENKA']), WD_ALIGN_VERTICAL.BOTTOM)
                            index_Table = index_Table + 1
                            
                        DiscType = "Факультативные дисциплины\nв том числе:"
                        DiscType_now = "Т"
                
                # Устанавливаем значение типа дисциплины в соответствующую ячейку таблицы
                workwithdocs.Insert_Str_from(mainTable, index_Table, DiscType, f"{Sum_Of_Ze} з.е.", "", WD_ALIGN_VERTICAL.BOTTOM)
                
                mainTable.cell(index_Table, 0).text = DiscType
                cell = mainTable.cell(index_Table, 0)
                # Применяем форматирование к ячейке
                workwithdocs.MakeFMT(cell, 11, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_VERTICAL.BOTTOM, 0, False)
                
                # Проверяем, нужно ли переключиться на другую таблицу или продолжить заполнение текущей
                if index_Table + 2 > len(mainTable.rows) - 1:
                    mainTable = rightTable
                    index_Table = 0
                else:
                    index_Table = index_Table + 2
            # Обработка информации о дисциплине и ее вставка в таблицу
            # Преобразование названия дисциплины в строку
            discname = str(row['DISCNAME'])

            # Если длина названия дисциплины больше или равна 55 символам и добавление строки не выйдет за пределы таблицы
            if len(discname) >= 55 and index_Table+(len(discname)//55) <= len(mainTable.rows) - 1:
                # Объединение ячеек таблицы в зависимости от длины названия дисциплины
                workwithdocs.MergeTable(mainTable,index_Table,(len(discname)//55))
                index_Table = index_Table + (len(discname)//55)
            # Если добавление строки выйдет за пределы таблицы, переходим к правой таблице
            elif index_Table+(len(discname)//55) > len(mainTable.rows) - 1:
                mainTable = rightTable
                index_Table = 0

            # Вставка информации о дисциплине в таблицу
            workwithdocs.Insert_Str_from(mainTable, 
                            index_Table, 
                            discname, 
                            str(row['DISCHOURS']), 
                            str(row['OCENKA']),
                            WD_ALIGN_VERTICAL.BOTTOM)

            # Проверка, не выйдем ли за пределы таблицы после вставки строки, если да, переходим к правой таблице
            if index_Table + 1 > len(mainTable.rows) - 1:
                mainTable = rightTable
                index_Table = 0
            else:
                index_Table = index_Table + 1
    doc.save(f"C:/Users/pavlov.sa/Desktop/Software/DiplomDVFU-1/Test/Unmerged/{file_save}")

def MakeAll():
    # Указываем пути к файлам
    shablon_path = "C:/Users/pavlov.sa/Desktop/Software/DiplomDVFU-1/Shablons/Diplom.docx"  # Путь к шаблону диплома
    data_file_path = "C:/Users/pavlov.sa/Desktop/Software/DiplomDVFU-1/Test/DataSet.xlsx"  # Путь к файлу с данными

    # Чтение данных из файла Excel в DataFrame
    df = pd.read_excel(data_file_path)

    # Преобразование DataFrame с помощью функции MakeDf из модуля workwithexcel
    df = workwithexcel.MakeDf(df)

    # Получение уникальных значений номеров зачётных книжек
    unique_zachbooks = df['ZACHBOOK'].unique()

    # Создание словаря zk_dfs, содержащего подмножества DataFrame для каждой зачётной книжки
    zk_dfs = {zachbook: df[df['ZACHBOOK'] == zachbook] for zachbook in unique_zachbooks}

    # Для каждой зачётной книжки создаём отдельный файл диплома
    for zachbook, df_part in zk_dfs.items():
        filename = f"{zachbook}.docx"  # Генерируем имя файла на основе номера зачётной книжки
        MakeOne(df_part, filename, shablon_path)  # Вызываем функцию MakeOne для создания файла диплома

    # Записываем отсортированные данные обратно в файл Excel
    df.to_excel("C:/Users/pavlov.sa/Desktop/Software/DiplomDVFU-1/Test/SortedDataSet.xlsx")