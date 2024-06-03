import workwithdocs
import workwithexcel
import pandas as pd
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx import Document


def MakeOne(df_student,file_save, shablon_path):
    doc = Document(shablon_path)

    index_Table = 0
    DiscType_now = "1"
    leftTable = None
    rightTable = None
    mainTable = None
    KursJob = pd.DataFrame(columns=['TYPECONTROL',
                                    'DISCNAME',
                                    'DISCHOURS',
                                    'OCENKA'])
    for table in doc.tables:
        if table.cell(0,0).text == "TABL1TABL":
            leftTable = table
        if table.cell(0,0).text == "TABL2TABL":
            rightTable = table
    mainTable = leftTable
    for index, row in df_student.iterrows():
        if row['TYPECONTROL'] == 'Курсовой проект' or row['TYPECONTROL'] == 'Курсовая работа':
            new_row = {'TYPECONTROL':row['TYPECONTROL'],
                       'DISCNAME':row['DISCNAME'],
                       'DISCHOURS':row['DISCHOURS'],
                       'OCENKA':row['OCENKA']}
            KursJob.loc[len(KursJob.index)] = new_row
        else:
            if not DiscType_now == str(row['DISCINDEX'])[1]:
                if index_Table+(len(discname)//54) > len(mainTable.rows) - 1:
                    mainTable = rightTable
                    index_Table = 0
                workwithdocs.MergeTable(mainTable,index_Table,1)
                match(str(row['DISCINDEX'])[1]):
                    case "2":
                        DiscType="Практики\nв том числе:"
                        DiscType_now = "2"
                    case "3":
                        DiscType="Государственная итоговая аттестация\nв том числе:"
                        DiscType_now = "3"
                    case "Т":
                        for jindex, row1 in KursJob.iterrows():
                            discname = f"{row1['TYPECONTROL']}, {row1['DISCNAME']}"
                            if len(discname) >= 54 and index_Table+(len(discname)//54) <= len(mainTable.rows) - 1:
                                workwithdocs.MergeTable(mainTable,index_Table,(len(discname)//54))
                                index_Table = index_Table + (len(discname)//54)
                            elif index_Table+(len(discname)//54) > len(mainTable.rows) - 1:
                                mainTable = rightTable
                                index_Table = 0
                            workwithdocs.Insert_Str_from(mainTable, 
                                            index_Table, 
                                            discname, 
                                            "", 
                                            str(row1['OCENKA']))
                            index_Table = index_Table + 1
                        #проверка
                        DiscType="Факультативные дисциплины\nв том числе:"
                        DiscType_now = "Т"
                mainTable.cell(index_Table, 0).text = DiscType
                cell = mainTable.cell(index_Table, 0)
                workwithdocs.MakeFMT(cell, 11, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_VERTICAL.BOTTOM, 0, False)
                if index_Table + 2 > len(mainTable.rows) - 1:
                    mainTable = rightTable
                    index_Table = 0
                else:
                    index_Table = index_Table + 2
            discname = str(row['DISCNAME'])
            if len(discname) >= 54 and index_Table+(len(discname)//54) <= len(mainTable.rows) - 1:
                workwithdocs.MergeTable(mainTable,index_Table,(len(discname)//54))
                index_Table = index_Table + (len(discname)//54)
            elif index_Table+(len(discname)//54) > len(mainTable.rows) - 1:
                mainTable = rightTable
                index_Table = 0
            workwithdocs.Insert_Str_from(mainTable, 
                            index_Table, 
                            discname, 
                            str(row['DISCHOURS']), 
                            str(row['OCENKA']))
            if index_Table + 1 > len(mainTable.rows) - 1:
                mainTable = rightTable
                index_Table = 0
            else:
                index_Table = index_Table + 1
    

    #зе проверка
    doc.save(f"C:/Users/pavlov.sa/Desktop/Software/DiplomDVFU-1/Test/Unmerged/{file_save}")

def MakeAll():
    shablon_path = "C:/Users/pavlov.sa/Desktop/Software/DiplomDVFU-1/Shablons/Diplom.docx"
    data_file_path = "C:/Users/pavlov.sa/Desktop/Software/DiplomDVFU-1/Test/DataSet.xlsx"
    df = pd.read_excel(data_file_path)
    df = workwithexcel.MakeDf(df)
    unique_zachbooks = df['ZACHBOOK'].unique()
    zk_dfs = {zachbook: df[df['ZACHBOOK'] == zachbook] for zachbook in unique_zachbooks}


    for zachbook, df_part in zk_dfs.items():
        filename = f"{zachbook}.docx"
        MakeOne(df_part, filename, shablon_path)
    df.to_excel("C:/Users/pavlov.sa/Desktop/Software/DiplomDVFU-1/Test/SortedDataSet.xlsx")