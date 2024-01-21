import docx
import xlrd
from threading import Thread

def CreateDocByTemplate(data: dict, templatePath: str) -> None:
    keys = list(data.keys())
    for i in range(len(data[keys[0]])):
        doc = docx.Document(templatePath)
        for key in keys:

            for parag in doc.paragraphs:
                if key in parag.text:
                    parag.text = parag.text.replace(key, data[key][i])

            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        if key in cell.text:
                            cell.text = cell.text.replace(key, data[key][i])

        doc.save(f"bin/{i}.docx")


def decompileExcel(dataPath:str) -> dict:
    # Open the Workbook
    workbook = xlrd.open_workbook(dataPath)
    rezult = {}
    # Выбор листа
    sheet = workbook.sheet_by_index(0)  # Или можно использовать sheet_by_name('имя_листа')

    # Получение первой строки
    first_row = sheet.row_values(0)

    # Формирование списка ключей
    keys_list = [str(key) for key in first_row]
    for key in keys_list:
        rezult[key] = []
    for i in range(1,sheet.nrows):
        row = sheet.row_values(i)
        for ii in range(len(keys_list)):
            rezult[keys_list[ii]].append(row[ii])


    return rezult

def createDocx(dataPath: str, templatepath: str):
    data = decompileExcel(dataPath)
    t1 = Thread(target=CreateDocByTemplate, args=(data,templatepath))
    t1.start()