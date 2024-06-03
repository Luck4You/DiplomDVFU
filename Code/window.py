from tkinter import *
from tkinter.ttk import *
import mainalgorythm

def updateProgressBar(parameter=None):
    if progBar['value'] < 100:
        if parameter is not None:
            if progBar['value'] < 100 and parameter + progBar['value'] > 100:
                progBar['value']=100
            else:
                progBar['value'] += parameter
            progBar['value'] += parameter
        else:
            progBar['value'] += 1
        window.update_idletasks()

window = Tk()  
window.title("Генерация дипломов")
window.geometry('800x400')
window.resizable(False,False)
window.option_add("*tearOff", FALSE)

bar_Variable = IntVar(window, 0)
choosenSetting_Variable = StringVar(window, "Один студент")
saveNameDoc_Variable = StringVar(window, "“ФИО (З.К.)”, как имя документа")

shooseSt_But = Button(window, text="Выбрать студетов")
shooseSt_But.place(x = 30, y = 125)
oneSt_RadioBut = Radiobutton(window, text="Один студент", variable = choosenSetting_Variable , value="Один студент")
oneSt_RadioBut.place(x = 30, y = 150)
oneSt_TextBox = Text(window, height=1, width=20)
oneSt_TextBox.place(x = 30, y = 175)
oneSt_TextBox.insert("1.0", "Зачетная книга")
oneSt_TextBox.tag_add("gray", "1.0", "1.end")
oneSt_TextBox.tag_config("gray", foreground="gray")

group_RadioBut = Radiobutton(window, text="Группа", variable = choosenSetting_Variable , value="Группа")
group_RadioBut.place(x = 200, y = 125)
group_TextBox = Text(window, height=1, width=20)
group_TextBox.place(x = 200, y = 150)
group_TextBox.insert("1.0", "Группа")
group_TextBox.tag_add("gray", "1.0", "1.end")
group_TextBox.tag_config("gray", foreground="gray")

dataSet_RadioBut = Radiobutton(window, text="Весь Dataset", variable = choosenSetting_Variable , value="Весь Dataset")
dataSet_RadioBut.place(x = 365, y = 125)

copyInPDF_ChBox = Checkbutton(window, text = 'Создание копий документов в формате “.pdf”')
copyInPDF_ChBox.place(x = 30, y = 50)

makeFullDocs_ChBox = Checkbutton(window, text = 'Создание объединенного документа')
makeFullDocs_ChBox.place(x = 320, y = 50)
makeFullDocs_But = Button(window, text = 'Расположение', command=updateProgressBar)
makeFullDocs_But.place(x = 320, y = 75)

fioAsName_RadioBut = Radiobutton(window, text="“ФИО (З.К.)”, как имя документа", variable = saveNameDoc_Variable, value="“ФИО (З.К.)”, как имя документа")
fioAsName_RadioBut.place(x = 550, y = 50)
customName_RadioBut = Radiobutton(window, text="Ручное", variable = saveNameDoc_Variable, value="Ручное")
customName_RadioBut.place(x = 550, y = 75)
customName_TextBox = Text(window, height=1, width=20)
customName_TextBox.place(x = 550, y = 100)
customName_TextBox.insert("1.0", "Название документа")
customName_TextBox.tag_add("gray", "1.0", "1.end")
customName_TextBox.tag_config("gray", foreground="gray")

progBar = Progressbar(window, orient="horizontal", length=750, value=0, mode="determinate")
progBar.place(x=25, y=350)

Generate_But = Button(window, text="Сформировать")
Generate_But.place(x=550, y=200)

main_menu = Menu()
 
file_menu = Menu()
file_menu.add_command(label="New")
file_menu.add_command(label="Save", command=mainalgorythm.MakeAll)
file_menu.add_command(label="Open")
file_menu.add_separator()
file_menu.add_command(label="Exit")
 
main_menu.add_cascade(label="File", menu=file_menu)
main_menu.add_cascade(label="Edit")
main_menu.add_cascade(label="View")

window.config(menu=main_menu)
# df = pd.read_excel("C:/Users/pavlov.sa/Desktop/dfe.xlsx")
# df_check = pd.read_excel("C:/Users/pavlov.sa/Desktop/Контингент 17.05.xlsx")

# for index, row in df_check.iterrows():
#     df['Зачётные книги студентов (при необходимости)'] = df['Зачётные книги студентов (при необходимости)'].str.replace(str(row['Зачетная книга']),str(row['ФИО']))

# df.to_excel("C:/Users/pavlov.sa/Desktop/Ready.xlsx")

window.mainloop()