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

#Replace file
def replaceFile(file_replace_path, file_data_path, new_data_path):
    dataReplace = pd.read_excel(file_replace_path, engine = 'openpyxl')
    dataReplaceFrame = pd.DataFrame(dataReplace, columns=['word','word_replace'])

    data = pd.read_excel(file_data_path, engine='openpyxl')
    dataFrame = pd.DataFrame(data, columns=['Emotion','Sentence'])
    
    workbook = xlsxwriter.Workbook(new_data_path)
    worksheet = workbook.add_worksheet()

    worksheet.write(0, 1, "Emotion")
    worksheet.write(0, 2, "Sentence")

    for index in range(len(dataFrame)):
        tag = dataFrame.Emotion[index]
        beautiData = []
        for indexReplace in range(len(dataReplaceFrame)):
            beautiData.append(dataFrame.Sentence[index].replace(dataReplaceFrame.word[indexReplace], dataReplaceFrame.word_replace[indexReplace]))
        worksheet.write(index+1, 0, index)
        worksheet.write(index+1, 1, tag)    
        worksheet.write(index+1, 2 , beautiData[len(dataReplaceFrame)-1])
    workbook.close()
    print("DONE!")
        
    return 

def menu():
    print("1: Add tag\n")
    print("2: Smooth data\n")
    option = int(input("Option: "))
    if(option==1):
        nameFile = input("Enter name file: ")
        writeTagFile(nameFile)
    else:
        file_replace = input("Enter file replace path: ")
        file_data = input("Enter file data path: ")
        file_output = input("Enter file output path: ")
        replaceFile(file_replace, file_data, file_output)

if __name__ == "__main__":
    menu()