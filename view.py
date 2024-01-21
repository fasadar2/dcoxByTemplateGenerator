import tkinter  as tk
from tkinter import ttk
from tkinter import filedialog
from printFunctions import createDocx

templatePath = ""
dataPath = ""
def readMeta():
    meta = {"varsion":"","changes":"","author":""}
    with open("meta.meta","r") as file:
        text = file.read()
        meta["version"] = str(text[text.find("version:")+1:text.find("&&")])
        meta["changes"] = text[text.find("changes:") + 1:text.find("&&")]
        meta["author"] = text[text.find("author:") + 1:text.find("&&")]
    return meta
def openTemplate():
    filepath = filedialog.askopenfilename(filetypes=[("docx files", "*.docx")])
    global templatePath
    templatePath = filepath
    filepath = filepath.split("/")
    filepath = filepath[len(filepath)-1]
    templatePathLabel.configure(text=f"Выбранный шаблон: {filepath}")
def openData():
    filepath = filedialog.askopenfilename(filetypes=[("excel files",[ "*.xlsx", "*.xls"])])
    global dataPath
    dataPath = filepath
    filepath = filepath.split("/")
    filepath = filepath[len(filepath) - 1]
    dataPathLabel.configure(text=f"Выбранный excel файл с данными: {filepath}")
if __name__ == '__main__':
    meta = readMeta()
    mainWindow = tk.Tk()
    version = meta["version"]
    mainWindow.title(f"Dcox Generator {version}")
    mainWindow.geometry("800x600")
    mainWindow.iconbitmap("mainIco.ico")
    tk.Label(text="Instructions",font='Times 14',justify='left').pack()
    tk.Label(text="1. Select a template",font='Times 14',justify='left').pack()
    tk.Label(text="2. Select a data file",font='Times 14',justify='left').pack()
    tk.Label(text="3. Press create Button",font='Times 14',justify='left').pack()

    openButtonTemplate = ttk.Button(text="Выбрать шаблон",width=50, command=openTemplate)
    openButtonTemplate.pack()
    templatePathLabel = ttk.Label(text="Выбранный шаблон:", font='Times 12', justify='left')
    templatePathLabel.pack()

    openButtonData = ttk.Button(text="Выбрать excel файл с данными",width=50, command=openData)
    openButtonData.pack()
    dataPathLabel = ttk.Label(text="Выбранный excel файл с данными:", font='Times 12', justify='left')
    dataPathLabel.pack()

    CreateDocxButton = ttk.Button(text="Заполнить документ",width=50, command=lambda : createDocx(dataPath,templatePath))
    CreateDocxButton.pack()

    mainWindow.mainloop()