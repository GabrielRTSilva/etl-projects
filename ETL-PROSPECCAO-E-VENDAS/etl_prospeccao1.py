
# ○•••••• FUNCIONALIDADE DO CÓDIGO ••••••○

# Esse código tem como objetivo tratar a planilha de Funil e Prospecção e Vendas do Jader.
# A trativa é a seguinte:
    # - Ela preenche os nomes das empresas para os seus respectivos contatos
    # - Preenche os dados vazios como sendo um str de valor 'Vazio'

# ○••••••••○•••••••••○•••••••••○•••••••••○


# ----- BIBLIOTECAS -----

import pandas as pd
import numpy as np 

# ----- COLUNAS DO DF -----

colunaID = []
colunaEmpresa = []
colunaSetorEmpresa = []
colunaCadastro = []
colunaOrigemContato = []
colunaContato = []
colunaCargo = []
colunaEstadoAtuacao = []
colunaMunicipio = []
colunaMeioContato = []
colunaHistorico = []

# ----- VAR GLOBAIS -----



id = 1
caractere = '\n'
caractere2 =  '.'

# ----- CODIGO -----

arquivo = pd.read_excel('PLANILHA-PFV.xlsx')

for nome in arquivo['EMPRESA']:
    if nome == 'FIM':
        break
    elif str(nome) != 'nan':
        varNome = nome.split(caractere)[0]
        colunaEmpresa.append(varNome)
    elif nome != varNome and str(nome) != 'nan':
        varNome = nome.split(caractere)[0]
    elif str(nome) == 'nan':
        colunaEmpresa.append(varNome)
    colunaID.append(id)
    id += 1

 
# ----- TRATAMENTO CELULAS VAZIAS -----

for setor in arquivo['SETOR DA EMPRESA']:
    if setor == 'FIM':
        break
    if str(setor) != 'nan':
        colunaSetorEmpresa.append(setor)
    elif str(setor) == 'nan':
        colunaSetorEmpresa.append('Vazio')

for cadastro in arquivo['CADASTRO']:
    if cadastro == 'FIM':
        break
    elif cadastro == 'NÃO':
        colunaCadastro.append(cadastro)
    elif cadastro == 'SIM':
        colunaCadastro.append(cadastro)
    elif str(cadastro) == 'nan':
        colunaCadastro.append('NÃO')

for origem in arquivo['ORIGEM CONTATO']:
    if origem == 'FIM':
        break
    if str(origem) == 'nan':
        colunaOrigemContato.append('Vazio')
    elif str(origem) == 'Contato profissional ':
        print('Achei o contato profissional')
        colunaOrigemContato.append('Contato Profissional')
    elif str(origem) != 'nan':
        colunaOrigemContato.append(origem)
    else:
        colunaOrigemContato.append(origem)

for contato in arquivo['CONTATO']:
    if contato == 'FIM':
        break
    if str(contato) != 'nan':
        colunaContato.append(contato)
    elif str(contato) == '':
        colunaOrigemContato.append('Vazio')
    elif str(contato) == 'nan':
        colunaContato.append('Vazio')      

for cargo in arquivo['CARGO']:
    if cargo == 'FIM':
        break
    if str(cargo) != 'nan':
        colunaCargo.append(cargo)
    elif str(cargo) == 'nan':
        colunaCargo.append('Vazio')    

for estado in arquivo['ESTADO DE ATUAÇÃO']:
    if estado == 'FIM':
        break
    if str(estado) != 'nan':
        colunaEstadoAtuacao.append(estado)
    elif str(estado) == 'nan':
        colunaEstadoAtuacao.append('Vazio')  

for municipio in arquivo['MUNICÍPIO ']:
    if municipio == 'FIM':
        break
    elif str(municipio) == 'nan':
        colunaMunicipio.append('Vazio')      
    elif str(municipio) != 'nan':
        colunaMunicipio.append(municipio)

for meio in arquivo['MEIO DE CONTATO']:
    if meio == 'FIM':
        break
    elif str(meio) == 'nan':
        colunaMeioContato.append('Vazio')  
    elif str(meio) != 'nan':
        colunaMeioContato.append(meio)


# ----- MONTANDO PLANILHA -----

novoPFV = {'ID': colunaID ,
           'EMPRESA': colunaEmpresa,
           'SETOR DA EMPRESA': colunaSetorEmpresa,
           'CADASTRO' : colunaCadastro,
           'ORIGEM CONTATO' : colunaOrigemContato,
           'CONTATO' : colunaContato,
           'CARGO': colunaCargo,
           'ESTADO DE ATUAÇÃO': colunaEstadoAtuacao,
            'MUNICÍPIO' : colunaMunicipio,
            'MEIO DE CONTATO' : colunaMeioContato
           }
 

# ----- EXPORTANDO PARA EXCEL -----

novoArquivo = pd.DataFrame(novoPFV)

novoArquivo.to_excel('Planilha V02.xlsx',index= False)


# ----- IDENTIFICADOR DE ERROS ------

'''colunas = [colunaID, colunaEmpresa, colunaSetorEmpresa, colunaCadastro, colunaOrigemContato, colunaContato, colunaCargo, colunaEstadoAtuacao, colunaMunicipio, colunaMeioContato,colunaHistorico]

col = 1

for coluna in colunas:
    print(f'Quantidade de itens da coluna {col} é {len(coluna)}')
    col += 1

listaType = []
index = 1

for item in colunaOrigemContato:
    listaType.append(type(item))

print(listaType.count(str()))



contagem_str = sum(isinstance(item, str) for item in colunaEmpresa)
print(f"A lista contém {contagem_str} itens do tipo 'str'.")'''




