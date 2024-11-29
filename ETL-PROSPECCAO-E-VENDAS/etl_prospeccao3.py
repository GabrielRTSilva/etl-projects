
# ○•••••• FUNCIONALIDADE DO CÓDIGO ••••••○

# Esse código tem como objetivo gerar a planilha 2 - Registro Contato

# ○••••••••○•••••••••○•••••••••○•••••••••○


import pandas as pd
import numpy as np
from datetime import datetime


index = 1
linha = 0

colunasRegistroContato = {'INDEX':[],
                          'PROSPECTOR':[],
                          'MEIO DE CONTATO': [],
                          'CONTATO':[],
                          'CARGO':[],
                          'EMPRESA':[],
                          'SETOR':[],
                          'CADASTRO':[],
                          'UF':[],
                          'MUNICÍPIO':[],
                          'INTERAÇÕES':[],
                          'DATA PRÓXIMA INTERAÇÃO': [],
                          'PRÓXIMA INTERAÇÃO':[]}

caminhoPlanilhaVFinal = 'C:/Users/gabriel.silva/Desktop/Python/Automate/ProjetosML/tarefaProspec.Vendas/1.planilhas/Planilha VFinal.xlsx'
planilhaPFV = pd.read_excel(caminhoPlanilhaVFinal)



for meio in planilhaPFV['MEIO DE CONTATO']:
    colunasRegistroContato['MEIO DE CONTATO'].append(meio.upper())


for contato in planilhaPFV['CONTATO']:
    colunasRegistroContato['INDEX'].append(index)
    colunasRegistroContato['PROSPECTOR'].append('Vazio')
    colunasRegistroContato['CONTATO'].append(contato.upper())
    index += 1
    linha += 1

for cargo in planilhaPFV['CARGO']:
    colunasRegistroContato['CARGO'].append(cargo.upper())

for empresa in planilhaPFV['EMPRESA']:
    colunasRegistroContato['EMPRESA'].append(empresa.upper())

for setor in planilhaPFV['SETOR DA EMPRESA']:
    colunasRegistroContato['SETOR'].append(setor.upper())

for cadastro in planilhaPFV['CADASTRO']:
    colunasRegistroContato['CADASTRO'].append(cadastro.upper())

for uf in planilhaPFV['ESTADO DE ATUAÇÃO']:
    colunasRegistroContato['UF'].append(uf)

for municipio in planilhaPFV['MUNICÍPIO']:
    colunasRegistroContato['MUNICÍPIO'].append(municipio)

for interacao in planilhaPFV['HISTORICO']:
    colunasRegistroContato['INTERAÇÕES'].append('Adicionar')
    colunasRegistroContato['DATA PRÓXIMA INTERAÇÃO'].append('Vazio')
    colunasRegistroContato['PRÓXIMA INTERAÇÃO'].append('Vazio')




caminhoPlanilhaFMO = 'C:/Users/gabriel.silva/Desktop/Python/Automate/ProjetosML/tarefaProspec.Vendas/1.planilhas/PLANILHA-FMO.xlsx'
planilhaFMO = pd.read_excel(caminhoPlanilhaFMO, sheet_name= 'Form1')

#print(planilhaFMO.shape[0])

for prospector in planilhaFMO['Name']:
    colunasRegistroContato['INDEX'].append(index)
    colunasRegistroContato['PROSPECTOR'].append(prospector.upper())
    index += 1

for meio in planilhaFMO['E-mail do Contato']:
    if str(meio) == 'nan':
        colunasRegistroContato['MEIO DE CONTATO'].append('Vazio')
    else:
        colunasRegistroContato['MEIO DE CONTATO'].append(meio.upper())

for contato in planilhaFMO['Nome do Contato']:
    colunasRegistroContato['CONTATO'].append(contato.upper())

for cargo in planilhaFMO['Cargo do Contato']:
    if str(cargo) == 'nan':
        colunasRegistroContato['CARGO'].append('Vazio')
    else:
        colunasRegistroContato['CARGO'].append(cargo.upper()) 

for empresa in planilhaFMO['Nome da Empresa']:
    colunasRegistroContato['EMPRESA'].append(empresa.upper())
    colunasRegistroContato['SETOR'].append('Vazio')
    colunasRegistroContato['CADASTRO'].append('Vazio')
    colunasRegistroContato['UF'].append('Vazio')
    colunasRegistroContato['MUNICÍPIO'].append('Vazio')
    colunasRegistroContato['INTERAÇÕES'].append('Adicionar2')

for next in planilhaFMO['Próxima Interação2']:
    colunasRegistroContato['PRÓXIMA INTERAÇÃO'].append(next)

for data in planilhaFMO['Data Próxima Ação']:
    colunasRegistroContato['DATA PRÓXIMA INTERAÇÃO'].append(data.strftime('%d/%m/%Y'))


planilhaRegistroContatos = pd.DataFrame(colunasRegistroContato)

planilhaRegistroContatos.to_excel('Planilha Registo de Contatos.xlsx', index = False)


