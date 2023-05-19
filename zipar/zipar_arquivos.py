import pandas as pd
import os
import zipfile

# Função para compactar uma pasta em um arquivo zip
def compactar_pasta(pasta, destino):
    with zipfile.ZipFile(destino, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root, _, files in os.walk(pasta):
            for file in files:
                file_path = os.path.join(root, file)
                zipf.write(file_path, os.path.relpath(file_path, pasta))

# Ler os IDs do Excel usando pandas
caminho_planilha = r'C:'  # Substitua pelo caminho correto da sua planilha
dados = pd.read_excel(caminho_planilha)
lista_ids = dados[1].tolist()  # Supondo que o nome da coluna com os IDs seja 'ID'

pasta_principal = r'C:'  # Substitua pelo caminho correto da sua pasta principal

for id in lista_ids:
    pasta_id = os.path.join(pasta_principal, str(id))
    if os.path.exists(pasta_id):
        nome_arquivo_zip = f'{id}.zip'  # Nome do arquivo zip será o ID com a extensão .zip
        destino_zip = os.path.join(pasta_principal, nome_arquivo_zip)
        compactar_pasta(pasta_id, destino_zip)
        print(f'Pasta {pasta_id} compactada em {destino_zip}')
    else:
        print(f'Pasta {pasta_id} não encontrada.')
