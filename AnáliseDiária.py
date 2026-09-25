import pandas as pd
import os
from tkinter import Tk
from tkinter.filedialog import askopenfilenames

Tk().withdraw()

arquivos = askopenfilenames(
    title="Selecione os arquivos Excel",
    filetypes=[("Arquivos Excel", "*.xlsx *.xls")]
)

if arquivos:

    for arquivo in arquivos:

        dados = pd.read_excel(arquivo, engine="openpyxl")

        unicos = dados["Recurso"].dropna().unique()

        total_programadas = dados['Recurso'].count()
        
        contagem_concluidos = (
            dados['Status da Atividade'] == 'concluído'
        ).sum()

        contagem_nao_concluidos = (
            dados['Status da Atividade'] == 'não concluído'
        ).sum()

        tipo_instalacao = (
            (dados['Descrição do Serviço'] ==
             '622299 - SE LIGA - INSTALAÇÃO DE LIGAÇÃO DE ÁGUA') &
            (dados['Status da Atividade'] == 'concluído')
        ).sum()

        tipo_religacao = (
            (dados['Descrição do Serviço'] ==
             '632221 - SE LIGA - RELIGAÇÃO SUPRESSÃO') &
            (dados['Status da Atividade'] == 'concluído')
        ).sum()

        tipo_padronizacao = (
            (dados['Descrição do Serviço'] ==
             '612713 - SE LIGA - PADRONIZAÇÃO DE LIGAÇÃO DE ÁGUA') &
            (dados['Status da Atividade'] == 'concluído')
        ).sum()

        nao_execucao_df = (
            dados['Motivo Não Execução']
            .value_counts()
            .reset_index()
        )

        nao_execucao_df.columns = [
            'Motivo Não Execução',
            'Quantidade'
        ]

        nao_execucao = nao_execucao_df.to_string(index=False)

        concluidos = dados[
            dados['Status da Atividade'] == 'concluído'
        ]

        cidades = pd.crosstab(
            concluidos['Cidade'],
            concluidos['Tipo de Atividade']
        )

        resumo = pd.DataFrame({
            'Indicador': [
                'Equipes',
                'Serviços Programados',
                'Serviços Concluídos',
                'Serviços Não Concluídos',
                'Instalação',
                'Religação',
                'Padronização',
                'Total'
            ],
            'Quantidade': [
                len(unicos),
                total_programadas,
                contagem_concluidos,
                contagem_nao_concluidos,
                tipo_instalacao,
                tipo_religacao,
                tipo_padronizacao,
                tipo_instalacao +
                tipo_religacao +
                tipo_padronizacao
            ]
        })
        nome_arquivo = os.path.basename(arquivo)
        nome_base = os.path.splitext(nome_arquivo)[0]

        arquivo_saida = os.path.join(
            os.path.dirname(arquivo),
            f"{nome_base}_Resumo.xlsx"
        )

        print(f"\nArquivo: {nome_arquivo}")
        print(f"Equipes: {len(unicos)}")
        print(f"Serviços Programados: {total_programadas}")
        print(f"Serviços Concluídos: {contagem_concluidos}")
        print(f"Serviços Não Concluídos: {contagem_nao_concluidos}")
        print(f"Instalação: {tipo_instalacao}")
        print(f"Religação: {tipo_religacao}")
        print(f"Padronização: {tipo_padronizacao}")
        print(
            f"Total: {tipo_instalacao + tipo_religacao + tipo_padronizacao}"
        )

        print("\nCidades x Tipo de Atividade")
        print(cidades)

        print("\nNão Executado")
        print(nao_execucao)

        print("-" * 60)

        with pd.ExcelWriter(
            arquivo_saida,
            engine="openpyxl"
        ) as writer:

            resumo.to_excel(
                writer,
                sheet_name="Resumo",
                index=False
            )

            cidades.to_excel(
                writer,
                sheet_name="Cidades"
            )

            nao_execucao_df.to_excel(
                writer,
                sheet_name="Nao Executado",
                index=False
            )

        print(f"Arquivo Excel gerado: {arquivo_saida}")
        print("=" * 60)

else:
    print("Nenhum arquivo selecionado.")



