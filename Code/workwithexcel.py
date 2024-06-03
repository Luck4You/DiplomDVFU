def MakeDf(df):
    df = Rename_Columns(df)
    
    df['DISCINDEX_STR'] = df['DISCINDEX'].apply(Sort_by_index) 
    df = df.sort_values(by=['ZACHBOOK', 'DISCINDEX_STR', 'STUDYYEAR', 'POLUGOD'])
    df = df.drop(columns=['DISCINDEX_STR'])
    return df

def Sort_by_index(name):
    name = name.replace(")", "")
    name = name.replace("(", ".")
    array = str(name).split(".")
    SumNum = 0
    for index, element in enumerate(array):
        if element.isdigit():
            match index:
                case 2:
                    SumNum += int(element)*100
                case 3:
                    SumNum += int(element)*10
            SumNum += int(element)
        elif element == "ФТД":
            SumNum = 99999
            continue
        else:
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
                    SumNum += 50000
                case "Б2":
                    SumNum += 60000
                case "Б3":
                    SumNum += 70000
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
    duplicates = dataframe.duplicated(subset=['ZACHBOOK', 
                                    'DISCNAME', 
                                    'STUDYYEAR', 
                                    'POLUGOD', 
                                    'TYPECONTROL'])
    duplicates.to_excel("Test/Duplicates.xlsx")\
    
    dataframe = dataframe.drop_duplicates(subset=['ZACHBOOK', 
                                    'DISCNAME', 
                                    'STUDYYEAR', 
                                    'POLUGOD', 
                                    'TYPECONTROL', 
                                    'OCENKA'])
    dataframe = dataframe[dataframe['DISCNAME'] != "Элективные курсы по физической культуре и спорту"]
    return 