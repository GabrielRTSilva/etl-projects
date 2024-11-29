
# ○•••••• FUNCIONALIDADE DO CÓDIGO ••••••○

# Esse código tem como objetivo gerar a planilha 1 - Empresa
# ETAPAS:
    # 1- IDENTIFICAR AS COLUNAS A SEREM PUXADAS PARA A NOVA PLANILHA
    # 2- GERAR O DICIONÁRIO: colunasPlanilhaEmpresa  
    # 3- PADRONIZAR OS DADOS DAS COLUNAS PUXADAS ( Upper Case, com exceção do UF e Município)

# ○••••••••○•••••••••○•••••••••○•••••••••○

import pandas as pd
import numpy as np


colunasPEmpresas = {'INDEX' : [],
                    'CONTATO': [],
                    'EMPRESA': [],
                    'SETOR': [],
                    'CADASTRO': [],
                    'UF': [],
                    'MUNICÍPIO': []}

index = 1


caminhoPlanilhaVFinal = 'C:/Users/gabriel.silva/Desktop/Python/Automate/ProjetosML/tarefaProspec.Vendas/1.planilhas/Planilha VFinal.xlsx'
planilhaPFV = pd.read_excel(caminhoPlanilhaVFinal)

for nome in planilhaPFV['CONTATO']:
    colunasPEmpresas['INDEX'].append(index)
    colunasPEmpresas['CONTATO'].append(nome.upper())
    index += 1

for empresa in planilhaPFV['EMPRESA']:
    colunasPEmpresas['EMPRESA'].append(empresa.upper())

for setor in planilhaPFV['SETOR DA EMPRESA']:
    colunasPEmpresas['SETOR'].append(setor.upper())

for cadastro in planilhaPFV['CADASTRO']:
    colunasPEmpresas['CADASTRO'].append(cadastro.upper())

for uf in planilhaPFV['ESTADO DE ATUAÇÃO']:
    colunasPEmpresas['UF'].append(uf)

for municipio in planilhaPFV['MUNICÍPIO']:
    colunasPEmpresas['MUNICÍPIO'].append(municipio)




caminhoPlanilhaFMO = 'C:/Users/gabriel.silva/Desktop/Python/Automate/ProjetosML/tarefaProspec.Vendas/1.planilhas/PLANILHA-FMO.xlsx'
planilhaFMO = pd.read_excel(caminhoPlanilhaFMO, sheet_name= 'Form1')

for nome1 in planilhaFMO['Nome do Contato']:
    colunasPEmpresas['INDEX'].append(index)
    colunasPEmpresas['CONTATO'].append(nome1.upper())
    index += 1

for empresa1 in planilhaFMO['Nome da Empresa']:
    colunasPEmpresas['EMPRESA'].append(empresa1.upper())
    colunasPEmpresas['SETOR'].append('Vazio')
    colunasPEmpresas['CADASTRO'].append('Vazio')
    colunasPEmpresas['UF'].append('Vazio')
    colunasPEmpresas['MUNICÍPIO'].append('Vazio')



planilhaEmpresa = pd.DataFrame(colunasPEmpresas)

planilhaEmpresa.to_excel('Planilha das Empresas.xlsx', index = False)



