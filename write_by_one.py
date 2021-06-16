import openpyxl

file = openpyxl.Workbook()  #创建一个工作簿

print(file.sheetnames)     #新建文档默认存在一个Sheet

print(file.active)

sheet = file["Sheet"]

for i in range(1,100000):
    for k in range(1,256):
        sheet.cell(row=i, column=k, value=i)
    



file.save("3.xlsx")


