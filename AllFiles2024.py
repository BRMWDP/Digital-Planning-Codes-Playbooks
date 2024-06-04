import os
import glob
import openpyxl
import pandas as pd
import re
import warnings
import time

# Inicia o cronômetro
start_time = time.time()

# Suprimir avisos sobre extensões desconhecidas
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

# Pasta principal onde estão as subpastas dos meses
main_folder_path = r'C:\Users\an770843\Mars Inc\Brazil_MW_PlanningComex - Documents\2 - DRP\2024\Arquivos Diários'

# Inicializa uma lista vazia para armazenar caminhos de arquivos
files = []

# Obtém todas as subpastas ordenadas por nome (que deve refletir a ordem cronológica)
subfolders = sorted([f.path for f in os.scandir(main_folder_path) if f.is_dir()])

# Seleciona apenas as últimas três pastas
subfolders = subfolders[-3:]

# Percorre as subpastas selecionadas e adiciona arquivos .xlsm à lista
for folder in subfolders:
    for filename in os.listdir(folder):
        if filename.endswith('.xlsm'):
            files.append(os.path.join(folder, filename))

# Função para extrair a data do nome do arquivo
def extract_date_from_filename(filename):
    match = re.search(r'\d{4}\.\d{2}\.\d{2}', filename)
    if match:
        return match.group(0)
    return None

# Inicializa uma lista vazia para armazenar DataFrames
df_list = []

# Cabeçalhos que serão usados em todos os DataFrames
headers = None

# Itera sobre cada arquivo na lista
for i, file in enumerate(files):
    # Extrai a data do nome do arquivo
    file_date = extract_date_from_filename(os.path.basename(file))
    
    # Carrega o arquivo .xlsm
    workbook = openpyxl.load_workbook(file, data_only=True)
    
    # Verifica se a aba "S&DCases" existe no arquivo
    if 'S&DCases' in workbook.sheetnames:
        # Carrega a aba "S&DCases"
        sheet = workbook['S&DCases']
        
        # Lê todas as informações da aba "S&DCases" e armazena em uma lista de listas
        data = list(sheet.values)
        
        if i == 0:
            # Para o primeiro arquivo, assume a 5ª linha como cabeçalho
            headers = data[4]
            data = data[5:]
        else:
            # Para os demais arquivos, desconsidera a 5ª linha
            data = data[5:]
        
        # Cria um DataFrame pandas com os dados
        df = pd.DataFrame(data)
        
        # Define as colunas do DataFrame usando os cabeçalhos do primeiro arquivo
        df.columns = headers
        
        # Remove as últimas 11 linhas do DataFrame
        if len(df) > 11:
            df = df[:-11]
        
        # Adiciona uma coluna com a data extraída do nome do arquivo
        df['Data'] = file_date
        
        # Adiciona o DataFrame à lista
        df_list.append(df)

# Verifica se a lista de DataFrames não está vazia antes de concatenar
if df_list:
    # Concatena todos os DataFrames em um único DataFrame
    final_df = pd.concat(df_list, ignore_index=True)
else:
    # Se não houver DataFrames na lista, cria um DataFrame vazio
    final_df = pd.DataFrame()

# Caminho para salvar o arquivo .xlsx
output_file = os.path.join(main_folder_path, 'dados_agrupados.xlsx')

# Salva o DataFrame final em um arquivo .xlsx, substituindo-o se já existir
final_df.to_excel(output_file, index=False)

# Para o cronômetro e calcula o tempo de execução
end_time = time.time()
execution_time = end_time - start_time

# Exibe o DataFrame final e o tempo de execução
print(f"Arquivo salvo em: {output_file}")
print(final_df)
print(f"Tempo de execução: {execution_time:.2f} segundos")