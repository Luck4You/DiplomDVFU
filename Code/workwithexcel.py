import pandas as pd

def MakeDf(df):
    df = Rename_Columns(df)
    df = Duplicates(df)
    df['DISCINDEX_STR'] = df['DISCINDEX'].apply(Sort_by_index) 
    df = df.sort_values(by=['ZACHBOOK', 'DISCINDEX_STR', 'STUDYYEAR', 'POLUGOD'])
    df = df.drop(columns=['DISCINDEX_STR'])
    df.to_excel("SortedDataSet.xlsx")
    return df

def Sort_by_index(name):
    # Функция для сортировки по индексу
    
    # Удаляем скобки из имени
    name = name.replace(")", "")
    name = name.replace("(", ".")
    
    # Разбиваем строку по точкам
    array = str(name).split(".")
    
    SumNum = 0
    
    # Проходим по элементам массива и обрабатываем каждый
    for index, element in enumerate(array):
        if element.isdigit():
            # Если элемент является числом, учитываем его в сумме
            match index:
                case 2:
                    SumNum += int(element)*100
                case 3:
                    SumNum += int(element)*10
            SumNum += int(element)
        elif element == "ФТД":
            # Если элемент равен "ФТД", присваиваем максимальное значение
            SumNum = 999999
            continue
        else:
            # Обрабатываем различные случаи для других значений
            match element:
                case "Б":
                    SumNum += 0
                case "О":
                    SumNum += 0
                case "В":
                    SumNum += 10000
                case "ДВ":
                    SumNum += 20000
                case "У":
                    SumNum += 30000
                case "П":
                    SumNum += 40000
                case "Б1":
                    SumNum += 500000
                case "Б2":
                    SumNum += 600000
                case "Б3":
                    SumNum += 700000
    
    return SumNum

def Rename_Columns(dataframe):
    dataframe = dataframe.rename(columns= 
    {
        'Школа студента' : ',MAINSCHOOL',
        'ФИО' : 'FIO',
        'Зачетная книга' : 'ZACHBOOK',
        'Наименование дисциплины' : 'DISCNAME',
        'Школа реализующая дисциплину' : 'DISCSCHOOL',
        'Учебная группа' : 'GROUP',
        'Учебный год' : 'STUDYYEAR',
        'Группа периодов контроля' : 'POLUGOD',
        'Система оценивания' : 'TYPEMARK',
        'Вид контроля' : 'TYPECONTROL',
        'Тип ведомости' : 'TYPEVEDOMOSTI',
        'Оценка' : 'OCENKA',
        'Номер ведомости' : 'NUMVEDOMOSTI',
        'Количество часов (Количество ЗЕТ)' : 'DISCHOURS',
        'Индекс' : 'DISCINDEX',
        'Дата проставления' : 'DATAPROST',
        'Кем проставлена' : 'WHOPROST',
        'Семестр' : 'SEMESTR',
        'Учебный план' : 'STUDYPLAN',
        'Рабочий план' : 'JOBPLAN',
        'Состояние' : 'SOSTOYANIE',
        'Выпускной курс' : 'VIPUSK',
        'Преподаватель ведомости' : 'PREPODVESOMOSTI'
    })
    return dataframe

def Duplicates(dataframe):
    # Функция для обработки дубликатов в датафрейме
    
    # Определяем ранжирование оценок
    grades_ranking = {
        'Неявка': 0,
        'Недопуск': 1,
        'Неудовлетворительно': 2,
        'Не зачтено': 3,
        'Удовлетворительно': 4,
        'Зачтено': 5,
        'Хорошо': 6,
        'Отлично': 7
    }

    # Присваиваем каждой оценке ранг
    dataframe['OCENKA_RANK'] = dataframe['OCENKA'].map(grades_ranking)
    
    # Сортируем данные по нескольким столбцам
    dataframe = dataframe.sort_values(by=['ZACHBOOK', 'DISCNAME', 'STUDYYEAR', 'POLUGOD', 'TYPECONTROL', 'OCENKA_RANK'], ascending=[True, True, True, True, True, False])

    # Находим дубликаты по определенным столбцам и сохраняем их в Excel
    duplicates = dataframe.duplicated(subset=['ZACHBOOK', 'DISCNAME', 'STUDYYEAR', 'POLUGOD', 'TYPECONTROL'])
    duplicates.to_excel("Test/Duplicates.xlsx")
    
    # Удаляем дубликаты, оставляя первое вхождение
    dataframe = dataframe.drop_duplicates(subset=['ZACHBOOK', 'DISCNAME', 'STUDYYEAR', 'POLUGOD', 'TYPECONTROL'], keep='first')
    
    # Удаляем временный столбец с рангом оценок
    dataframe = dataframe.drop('OCENKA_RANK', axis=1)

    # Удаляем дубликаты по расширенному набору столбцов
    dataframe = dataframe.drop_duplicates(subset=['ZACHBOOK', 'DISCNAME', 'STUDYYEAR', 'POLUGOD', 'TYPECONTROL', 'OCENKA'])
    
    # Удаляем строки, где название предмета - "Элективные курсы по физической культуре и спорту"
    dataframe = dataframe[dataframe['DISCNAME'] != "Элективные курсы по физической культуре и спорту"]
    
    return dataframe


def process_bad_marks(dataframe):
    # Функция для обработки плохих оценок в датафрейме
    # Выбираем строки с оценками 'Неудовлетворительно', 'Не зачтено', 'Неявка', 'Недопуск'
    df_bad_marks = dataframe[dataframe['OCENKA'].isin(['Неудовлетворительно', 'Не зачтено', 'Неявка', 'Недопуск'])]
    
    # Получаем столбец 'ZACHBOOK' для строк с плохими оценками
    zachbook_bad_marks = df_bad_marks['ZACHBOOK']
    
    # Преобразуем столбец 'ZACHBOOK' в список
    zachbook_bad_marks_list = zachbook_bad_marks.tolist()
    
    # Удаляем строки с плохими оценками из исходного датафрейма
    dataframe = dataframe[~dataframe['ZACHBOOK'].isin(zachbook_bad_marks)]
    
    # Создаем новый датафрейм с 'ZACHBOOK' плохих оценок и сохраняем его в Excel
    zachbook_bad_marks_df = pd.DataFrame(zachbook_bad_marks_list, columns=['ZACHBOOK'])
    zachbook_bad_marks_df.to_excel('zachbook_bad_marks.xlsx', index=False)
    
    return dataframe


