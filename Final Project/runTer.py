from rich.console import Console
from rich.table import Table
import xlrd
import pandas as pd
import xlsxwriter


def showTag():
    table = Table(title="KEY TAG")
    for i in range(0,8):
        table.add_column(str(i))
    table.add_row("chán ghét", "thích thú", "giận dữ", "ngạc nhiên", "buồn bã", "sợ hãi", "khác", "LOẠI")
    
    console = Console()
    console.print(table)

def writeTagFile(file_path):
    workbook = xlsxwriter.Workbook(file_path)
    worksheet = workbook.add_worksheet()

    data = pd.read_excel(file_path, engine="openpyxl")
    df = pd.DataFrame(data, columns=['comment'])

    col = 0
    row = 0
    count = 0
    index = 0
    worksheet.write(0, 1, "comment")
    worksheet.write(0, 2, "tag")

    while (count < 101):
        showTag()
        print("\nCount: ", count,"\n")
        print(df.comment[index],"\n")
        tag = input("Tag: ")

        if (tag != "7"):
            worksheet.write(count+1, 0, count)
            worksheet.write(count+1, 1, df.comment[index])
            worksheet.write(count+1, 2, tag)
            count +=1
            index +=1
        else:
            count = count
            index+=1

    workbook.close()

if __name__ == "__main__":
    nameFile = input("Enter name file: ")
    writeTagFile(nameFile)