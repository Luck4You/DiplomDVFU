# # Логика обработки других типов дисциплин в зависимости от их индекса
#             # Проверяем, не изменился ли тип дисциплины с предыдущей итерации
#             if not DiscType_now == str(row['DISCINDEX'])[1]:
#                 # Проверяем, помещается ли новая информация в текущую таблицу или нужно переключиться на другую
#                 if index_Table + 1 > len(mainTable.rows) - 1:
#                     mainTable = rightTable
#                     index_Table = 0
                
#                 # Объединяем ячейки таблицы, если необходимо
#                 workwithdocs.MergeTable(mainTable, index_Table, 1)
                
#                 Sum_Of_Ze = 0
#                 # Определяем тип дисциплины и обновляем текущий тип
#                 match(str(row['DISCINDEX'])[1]):
#                     case "2":
#                         DiscType = "Практики\nв том числе:"
#                         DiscType_now = "2"
#                         Sum_Of_Ze = SumOfZe(df_student,"Б2")
#                     case "3":
#                         DiscType = "Государственная итоговая аттестация\nв том числе:"
#                         DiscType_now = "3"
#                         Sum_Of_Ze = SumOfZe(df_student,"Б3")
#                     case "Т":
#                         # Циклически обрабатываем данные из DataFrame KursJob для типа "Т"
#                         for jindex, row1 in KursJob.iterrows():
#                             discname = f"{row1['TYPECONTROL']}, {row1['DISCNAME']}"
#                             # Объединяем ячейки таблицы при необходимости
#                             if len(discname) >= 55 and index_Table + (len(discname) // 55) <= len(mainTable.rows) - 1:
#                                 workwithdocs.MergeTable(mainTable, index_Table, (len(discname) // 55))
#                                 index_Table = index_Table + (len(discname) // 55)
#                             elif index_Table + (len(discname) // 55) > len(mainTable.rows) - 1:
#                                 mainTable = rightTable
#                                 index_Table = 0
                            
#                             # Вставляем информацию в таблицу из DataFrame KursJob
#                             workwithdocs.Insert_Str_from(mainTable, index_Table, discname, "", str(row1['OCENKA']), WD_ALIGN_VERTICAL.BOTTOM)
#                             index_Table = index_Table + 1
                            
#                         DiscType = "Факультативные дисциплины\nв том числе:"
#                         DiscType_now = "Т"
                
#                 # Устанавливаем значение типа дисциплины в соответствующую ячейку таблицы
#                 workwithdocs.Insert_Str_from(mainTable, index_Table, DiscType, f"{Sum_Of_Ze} з.е.", "", WD_ALIGN_VERTICAL.BOTTOM)
                
#                 mainTable.cell(index_Table, 0).text = DiscType
#                 cell = mainTable.cell(index_Table, 0)
#                 # Применяем форматирование к ячейке
#                 workwithdocs.MakeFMT(cell, 11, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_VERTICAL.BOTTOM, 0, False)
                
#                 # Проверяем, нужно ли переключиться на другую таблицу или продолжить заполнение текущей
#                 if index_Table + 2 > len(mainTable.rows) - 1:
#                     mainTable = rightTable
#                     index_Table = 0
#                 else:
#                     index_Table = index_Table + 2
#             # Обработка информации о дисциплине и ее вставка в таблицу
#             # Преобразование названия дисциплины в строку
#             discname = str(row['DISCNAME'])

#             # Если длина названия дисциплины больше или равна 55 символам и добавление строки не выйдет за пределы таблицы
#             if len(discname) >= 55 and index_Table+(len(discname)//55) <= len(mainTable.rows) - 1:
#                 # Объединение ячеек таблицы в зависимости от длины названия дисциплины
#                 workwithdocs.MergeTable(mainTable,index_Table,(len(discname)//55))
#                 index_Table = index_Table + (len(discname)//55)
#             # Если добавление строки выйдет за пределы таблицы, переходим к правой таблице
#             elif index_Table+(len(discname)//55) > len(mainTable.rows) - 1:
#                 mainTable = rightTable
#                 index_Table = 0

#             # Вставка информации о дисциплине в таблицу
#             workwithdocs.Insert_Str_from(mainTable, 
#                             index_Table, 
#                             discname, 
#                             str(row['DISCHOURS']), 
#                             str(row['OCENKA']),
#                             WD_ALIGN_VERTICAL.BOTTOM)

#             # Проверка, не выйдем ли за пределы таблицы после вставки строки, если да, переходим к правой таблице
#             if index_Table + 1 > len(mainTable.rows) - 1:
#                 mainTable = rightTable
#                 index_Table = 0
#             else:
#                 index_Table = index_Table + 1