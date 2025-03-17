import openpyxl as xl
import re as re

arquivo = xl.load_workbook(filename="C:/Users/mamfy/Desktop/prog/proj03-CRED/exemplo_dados_com_cpf.xlsx")

ws = arquivo.active

def pesquisa(query):

    query = re.sub(r'\D', '', query)

    for row in ws.iter_rows(values_only=True): 
        if re.sub(r'\D', '', str(row[0])) == query:
            print(row)
            return
    print("404")

while True:
    try:
        query = input("Digite o nome para pesquisa: ")
        pesquisa(query)
        if query.lower() == 'sair':
            break
    except EOFError:
        break

