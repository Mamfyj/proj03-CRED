import openpyxl as xl

arquivo = xl.load_workbook(filename="C:/Users/mamfy/Desktop/prog/proj03-CRED/teste.xlsx")

ws = arquivo.active

for row in ws.iter_rows():
    if row[0].value.find("João") != -1:
        print([cell.value for cell in row])
