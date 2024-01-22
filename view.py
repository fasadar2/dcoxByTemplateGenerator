import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from printFunctions import CreateDocx
import tkinter.messagebox as message

templatePath = ""
dataPath = ""


def ReadMeta():
    meta = {"varsion": "", "changes": "", "author": ""}
    with open("_internal/meta.meta", "r") as file:
        text = file.read()
        meta["version"] = str(text[text.find("version:") + 1:text.find("&&")])
        meta["author"] = text[text.find("author:") + 1:text.find("&&")]
    return meta


def OpenTemplate():
    filepath = filedialog.askopenfilename(filetypes=[("docx files", "*.docx")])
    global templatePath
    templatePath = filepath
    filepath = filepath.split("/")
    filepath = filepath[len(filepath) - 1]
    templatePathLabel.configure(text=f"Выбранный шаблон: {filepath}")


def OpenData():
    filepath = filedialog.askopenfilename(filetypes=[("excel files", ["*.xlsx", "*.xls"])])
    global dataPath
    dataPath = filepath
    filepath = filepath.split("/")
    filepath = filepath[len(filepath) - 1]
    dataPathLabel.configure(text=f"Выбранный excel файл с данными: {filepath}")


if __name__ == '__main__':
    meta = ReadMeta()
    mainWindow = tk.Tk()
    version = meta["version"]
    author = meta["author"]
    mainWindow.title(f"Dcox Generator {version}")
    mainWindow.geometry("800x600")
    mainWindow.iconbitmap("_internal/mainIco.ico")
    ttk.Label(text="Инструкция:", font='Times 14', justify='left', style="BW.TLabel")\
        .grid(column=1, row=0,ipadx=6,columnspan=2, ipady=6, padx=4, pady=4,sticky=tk.W)
    ttk.Label(text="1. Выберите шаблон в формате docx", font='Times 14', justify='left', style="BW.TLabel")\
        .grid(column=1, row=1,ipadx=6,columnspan=2, ipady=6, padx=4, pady=4,sticky=tk.W)
    ttk.Label(text="2. Выберите файл с данными в формате xls или xlsx", font='Times 14', justify='left', style="BW.TLabel")\
        .grid(column=1, row=2,ipadx=6,columnspan=2, ipady=6, padx=4, pady=4,sticky=tk.W)
    ttk.Label(text="3. Нажмите кнопку Заполнить документы", font='Times 14', justify='left', style="BW.TLabel")\
        .grid(column=1, row=3,ipadx=6,columnspan=2, ipady=6, padx=4, pady=4,sticky=tk.W)

    openButtonTemplate = ttk.Button(text="Выбрать шаблон", width=50, command=OpenTemplate)
    openButtonTemplate.grid(column=1, row=4,ipadx=6, ipady=6, padx=4, pady=4,sticky=tk.W)
    templatePathLabel = ttk.Label(text="Выбранный шаблон:", font='Times 12', justify='left', style="BW.TLabel")
    templatePathLabel.grid(column=2, row=4,ipadx=6, ipady=6, padx=4, pady=4,sticky=tk.W)

    openButtonData = ttk.Button(text="Выбрать excel файл с данными", width=50, command=OpenData)
    openButtonData.grid(column=1, row=5,ipadx=6, ipady=6, padx=4, pady=4,sticky=tk.W)
    dataPathLabel = ttk.Label(text="Выбранный excel файл с данными:", font='Times 12', justify='left')
    dataPathLabel.grid(column=2, row=5,ipadx=6, ipady=6, padx=4, pady=4,sticky=tk.W)
    CreateDocxButton = ttk.Button(text="Заполнить документы", width=50,
                                  command=lambda: CreateDocx(dataPath, templatePath))
    CreateDocxButton.grid(column=1, row=6,ipadx=6, ipady=6, padx=4, pady=4,sticky=tk.W)

    mainWindow.mainloop()
