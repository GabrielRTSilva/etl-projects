
#  OBJETIVO DO ALGORITMO
    # Esse algoritmo cria uma PBI externa para cada projeto no caminhoPlanilha.
    # Utilizando a biblioteca Pandas.

# ---- BIBLIOTECAS ----
import pandas as pd
import numpy as np
import os
import datetime

# ---- VARIÁVEIS GLOBAIS ----
caractere = '.'
linha = 0


#ABAIXO A LINHA EM QUE AS PLANILHAS PARA EXTRAÇÃO E LIMPEZA DEVERÃO ESTAR
caminhoProjetos = 'digite_aqui_o_caminho_das_planilhas'


# ---- LOOP DE EXTRAÇÃO ----
for arquivo in os.listdir(caminhoProjetos):

    
    caminhoPlanilha = os.path.join(caminhoProjetos, arquivo)
    nomePBI = arquivo.split(caractere)[0]
    projeto = pd.read_excel(caminhoPlanilha, sheet_name= 'PBI')
    
    colunasPlanilha = {'DATA': [],
              'FAM': [],
              'ITEM': [],
              'DESCRIÇÃO': [],
              'CATEGORIA': [],
              'SUBCATEGORIA': [],
              'VALOR': []}
    linha = 0
    
    for item in projeto['VALOR']:
        if not isinstance(item, str) and str(item) != 'nan' and item != 0:
            
            data = projeto.iloc[linha,0]
            if isinstance(data, datetime.time) or isinstance(data, str):
                colunasPlanilha['DATA'].append(data)
            else:
                colunasPlanilha['DATA'].append(data.strftime('%d/%m/%Y'))

            colunasPlanilha['FAM'].append(projeto.iloc[linha,1])
            colunasPlanilha['ITEM'].append(projeto.iloc[linha,2])
            colunasPlanilha['DESCRIÇÃO'].append(projeto.iloc[linha,3])
            colunasPlanilha['CATEGORIA'].append(projeto.iloc[linha,4])
            colunasPlanilha['SUBCATEGORIA'].append(projeto.iloc[linha,5])
            colunasPlanilha['VALOR'].append(round(item,2))

        linha += 1
    
    
    newPBI = pd.DataFrame(colunasPlanilha)
    newPBI.to_excel(f'{nomePBI}-PBI.xlsx', index= False)
