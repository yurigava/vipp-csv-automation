from openpyxl import Workbook
import csv
from collections import OrderedDict
import pickle
from unidecode import unidecode
pacName = 'pac'
sedexName = 'sedex'

columns = [
    'idPedido',
    'nome',
    'cep',
    'endereco',
    'numero',
    'complemento',
    'bairro',
    'cidade',
    'estado',
    'telefone',
    'celular',
    'email',
    'cpf',
    'conteudo',
    'valor'
]
data = {pacName: OrderedDict([(name, []) for name in columns]),
        sedexName: OrderedDict([(name, []) for name in columns])}


def getConteudo(pedidoId, csvFile, csvReader):
    found = ([linha[3], linha[10], linha[11]] for linha in csvReader if linha[0] == pedidoId)
    csvFile.seek(0)
    return list(found)


def getRowByShipping(rows, shippingType):
    return [row for row in rows if row[8].lower() == shippingType]


with open('pedidos.csv', newline='', encoding='iso-8859-1') as csvfile:
    reader = csv.reader(csvfile, delimiter=';')
    for shipType in data:
        currentData = data[shipType]
        filteredLines = getRowByShipping(reader, shipType)
        for rowNum, row in enumerate(filteredLines):
            currentData['idPedido'].append(row[0])
            currentData['nome'].append(row[27])
            currentData['endereco'].append(row[14])
            currentData['numero'].append(row[15])
            currentData['complemento'].append(row[16])
            currentData['bairro'].append(row[17])
            currentData['cidade'].append(row[18])
            currentData['estado'].append(row[19])
            currentData['cep'].append(row[20])
            currentData['email'].append(row[22])
            currentData['cpf'].append(row[23].split('"')[1])
            currentData['valor'].append(row[6])

with open('conteudo.csv', newline='', encoding='iso-8859-1') as csvfile:
    reader2 = csv.reader(csvfile, delimiter=';')
    for shipType in data:
        currentData = data[shipType]
        currentData['conteudo'] = [getConteudo(currentId, csvfile, reader2) for currentId in
                                   currentData['idPedido']]
        wb = Workbook()
        ws = wb.active
        for index in range(len(idPedido)):
            ws.cell(column=1, row=index+1, value=nome[index])
            ws.cell(column=2, row=index+1, value=cep[index])
            ws.cell(column=3, row=index+1, value=endereco[index])
            ws.cell(column=4, row=index+1, value=numero[index])
            ws.cell(column=5, row=index+1, value=complemento[index])
            ws.cell(column=6, row=index+1, value=bairro[index])
            ws.cell(column=7, row=index+1, value=cidade[index])
            ws.cell(column=8, row=index+1, value=estado[index])
            ws.cell(column=12, row=index+1, value=email[index])
            ws.cell(column=13, row=index+1, value=cpf[index])
            ws.cell(column=27, row=index+1, value=valor[index])
        wb.save(f'editado{shipType}.xlsx')

#print(f'Nome: {nome[0]}')
#print(f'CEP: {cep[0]}')
#print(f'Endereço: {endereco[0]}')
#print(f'Numero: {numero[0]}')
#print(f'Complemento: {complemento[0]}')
#print(f'Bairro: {bairro[0]}')
#print(f'Cidade: {cidade[0]}')
#print(f'Estado: {estado[0]}')
#print(f'email: {email[0]}')#11
#print(f'cpf: {cpf[0]}')#12
#valor 27
#conteudo 34
