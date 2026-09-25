import pandas as pd
import os
from datetime import datetime
from tkinter import Tk
from tkinter.filedialog import askopenfilename
contador = 1
data_atual = datetime.now().strftime("%d-%m-%y")
while os.path.exists(f"Relatório_Motivo Não Execução_{contador}.xlsx"):
    contador += 1
nome_arquivo_encontradas = f"Relatório_Motivo Não Execução_{contador}_{data_atual}.xlsx"
arquivo_postergadas = askopenfilename(filetypes=[("Arquivos Excel", "*.xlsx *.xls")])
base = pd.read_excel(arquivo_postergadas, engine='openpyxl')
resultado1 = base.loc[base['Motivo Não Execução'] == 'Inexistência de Rede - Postergar', 'Matrícula Unidade Comercial']
resultado1_1 = base.loc[base['Motivo Não Execução'] == 'Postergar - Necessita Escavação', 'Matrícula Unidade Comercial']
resultado1_2 = base.loc[base['Motivo Não Execução'] == 'Falta de Material', [
    'Recurso',
    'Ordem de Serviço',
    'Matrícula Unidade Comercial',
    'Parecer'
    ]
]

resultado1_3 = base.loc[base['Motivo Não Execução'] == 'Cliente não permitiu',[
    'Recurso',
    'Ordem de Serviço',
    'Matrícula Unidade Comercial',
    'Tipo de Atividade',
    'Parecer'
    ]
]

'''base_parecer = base.loc[base['Parecer'] == 'Postergar - Necessita Escavação', 'Matrícula Unidade Comercial']'''
base2 = pd.read_excel('os_selecionadas_20260923_0901.xlsx')


'''resultado2_1 = pd.merge(
    resultado1_1,
    base_parecer[['Parecer']],
    on='Matrícula Unidade Comercial',
    how='left'
)'''


'''for converter in ['NR_COORDENADA_X','NR_COORDENADA_Y']:
    base2[converter] = pd.to_numeric(
        base2[converter].astype(str).str.replace(',', '.'), errors="coerce"
    )'''

resultado2 = base2[
    base2['Matrícula Unidade Comercial'].isin(resultado1)]

resultado2_1 = base2[
    base2['Matrícula Unidade Comercial'].isin(resultado1_1)]


resultado_final = resultado2[
    [
        'Numero OS',
        'Matrícula Unidade Comercial',
        'Bucket Distancia',
        'NR_COORDENADA_X',
        'NR_COORDENADA_Y',
        'LOCAL',
        'BAIRRO_GEO',
        'Endereço_TOA',
        'DIST_CLIENTE_ATUAL_M',
        'DIST_REDE_AGUA',
        'DIST_INFRA_M',
        'DIST_VIA_M'
    ]
]



resultado_escavacao = resultado2_1[
    [
        'Numero OS',
        'Matrícula Unidade Comercial',
        'NR_COORDENADA_X',
        'NR_COORDENADA_Y',
        'LOCAL',
        'BAIRRO_GEO',
        'Endereço_TOA',
        'TIPO PAVIMENTO'
        
     
    ]

]

resultado_final = resultado_final.sort_values(
    by=['DIST_CLIENTE_ATUAL_M'], ascending=True)

nao_encontrados = resultado1[
    ~resultado1.isin(base2['Matrícula Unidade Comercial'])
]



with pd.ExcelWriter(nome_arquivo_encontradas) as writer: 
    resultado_final.to_excel(writer, sheet_name="Inexistência de Rede", index=False)
    '''nao_encontrados.to_excel(writer, sheet_name="Matriculas Não encontradas", index=False)'''
    resultado_escavacao.to_excel(writer,sheet_name="Necessita Escavação", index=False)
    resultado1_2.to_excel(writer,sheet_name="Falta de Material", index=False)
    resultado1_3.to_excel(writer,sheet_name='Cliente não permitiu', index=False)





print("Arquivo salvo")

