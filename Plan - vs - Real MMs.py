import pandas as pd
import os
import glob

# Caminho da pasta no SharePoint
caminho_pasta = r'C:\\Users\\an770843\\Mars Inc\\Brazil_MW_PlanningComex - Documents\\1 - PCP\\2 - Plano de produção\\2025\\M&Ms\\Compilado Automação BI'

# Caminho do arquivo Calendar.xlsx
caminho_calendar = r'C:\\Users\\an770843\\Mars Inc\\Brazil_MW_PlanningComex - Documents\\1 - PCP\\2 - Plano de produção\\2025\\M&Ms\\Compilado Automação BI\\Calendar.xlsx'

def reorganizar_dados(caminho_pasta):
    print("Iniciando a reorganização dos dados...")  # Mensagem de início
    
    # Criando uma lista para armazenar todos os dados reorganizados
    dados_reorganizados = []

    # Extraindo o ano do caminho da pasta
    partes_caminho = caminho_pasta.split('\\')  # Divide o caminho em partes
    if len(partes_caminho) > 14:  # Verifica se há pelo menos 15 partes
        ano = partes_caminho[14]  # Pega a parte que contém o ano (15ª parte, índice 14)
    else:
        print("Caminho da pasta não contém partes suficientes para extrair o ano.")
        return

    # Verifica se o ano tem 4 dígitos e é um número
    if len(ano) != 4 or not ano.isdigit():
        print("Ano extraído não é válido:", ano)
        return

    # Buscando todos os arquivos .xlsm na pasta
    arquivos = glob.glob(os.path.join(caminho_pasta, "*.xlsm"))
    
    for arquivo_excel in arquivos:
        # Extraindo o nome do arquivo para obter o período e a semana
        nome_arquivo = os.path.basename(arquivo_excel)
        
        # Extraindo Período e Week do nome do arquivo
        periodo = nome_arquivo[16:19]  # Pega os 3 caracteres a partir da posição 17 (índice 16)
        semana = nome_arquivo[nome_arquivo.find('W'):nome_arquivo.find('W') + 2]  # Wx

        try:
            # Lendo a aba do arquivo Excel, especificando que o cabeçalho está na linha 2 (índice 1)
            df = pd.read_excel(arquivo_excel, sheet_name="4_Plano (Nova)", header=1)
            print(f"Arquivo '{nome_arquivo}' lido com sucesso!")
        except Exception as e:
            print(f"Erro ao ler o arquivo '{nome_arquivo}': {e}")
            continue  # Pula para o próximo arquivo em caso de erro

        # Exibindo as primeiras linhas do DataFrame original
        print("Dados originais:")
        print(df.head())

        # Preenchendo as linhas vazias na coluna "Máquina"
        df['Máquina'] = df['Máquina'].fillna(method='ffill')

        # Ignorando a última linha que contém "Total GERAL"
        df = df[:-1]

        # Removendo linhas que contêm "Total GERAL" na coluna "Máquina"
        df = df[df['Máquina'] != 'Total GERAL']

        # Verificando se as colunas necessárias estão presentes
        colunas_necessarias = ['Máquina', 'SKU', 'Rate', 'Dom', 'Seg', 'Ter', 'Qua', 'Qui', 'Sex', 'Sáb']
        for coluna in colunas_necessarias:
            if coluna not in df.columns:
                print(f"A coluna '{coluna}' não foi encontrada no DataFrame.")
                continue

        # Mapeamento dos dias da semana para números
        dia_semana_mapping = {
            'Dom': 1,
            'Seg': 2,
            'Ter': 3,
            'Qua': 4,
            'Qui': 5,
            'Sex': 6,
            'Sáb': 7
        }

        # Iterando sobre as linhas do DataFrame
        for index, row in df.iterrows():
            # Ignorando linhas que contêm "Total GERAL" na coluna "SKU"
            if isinstance(row['SKU'], str) and 'Total GERAL' in row['SKU']:
                continue

            # Extraindo informações da linha
            maquina = row['Máquina']
            sku = row['SKU']
            rate = row['Rate']

            # Iterando sobre os dias da semana (incluindo Domingo)
            for dia in ['Dom', 'Seg', 'Ter', 'Qua', 'Qui', 'Sex', 'Sáb']:
                quantidade = row[dia]

                # Verificando se a quantidade é válida (não é um traço ou zero)
                if isinstance(quantidade, (int, float)) and quantidade > 0:
                    # Adicionando os dados reorganizados à lista
                    dados_reorganizados.append({
                        'Máquina': maquina,
                        'SKU': sku,
                        'Rate': rate,
                        'Data': dia,
                        'Quantidade': quantidade,
                        'Período': periodo,  # Adicionando o período correto
                        'Week': semana,      # Adicionando a semana correta
                        'Ano': ano,          # Adicionando o ano extraído
                        'Dia Semana': dia_semana_mapping[dia],  # Adicionando o dia da semana mapeado
                        'P/W/Y/D': periodo + semana + ano + str(dia_semana_mapping[dia])  # Concatenando para P/W/Y/D
                    })

    # Verificando se há dados reorganizados
    if not dados_reorganizados:
        print("Nenhum dado válido foi encontrado para reorganização.")
        return

    # Criando um DataFrame a partir dos dados reorganizados
    df_reorganizado = pd.DataFrame(dados_reorganizados)

    # Exibindo as primeiras linhas do DataFrame reorganizado
    print("Dados reorganizados:")
    print(df_reorganizado.head())

    # Lendo o arquivo Calendar.xlsx e a aba Original
    try:
        calendar_df = pd.read_excel(caminho_calendar, sheet_name="Original")
        print("Arquivo 'Calendar.xlsx' lido com sucesso!")
        print("Colunas disponíveis em 'Calendar.xlsx':", calendar_df.columns)  # Verifica as colunas
    except Exception as e:
        print(f"Erro ao ler o arquivo 'Calendar.xlsx': {e}")
        return

    # Realizando o PROCX para preencher a coluna 'DataFormat'
    df_reorganizado = df_reorganizado.merge(
        calendar_df[['P/W/Y/D', 'Date2']],  # Seleciona as colunas relevantes
        left_on='P/W/Y/D',                   # Coluna da tabela reorganizada
        right_on='P/W/Y/D',                  # Coluna da tabela Calendar
        how='left'                           # Faz um merge à esquerda
    )

    # Renomeia a coluna resultante para 'DataFormat'
    df_reorganizado.rename(columns={'Date2': 'DataFormat'}, inplace=True)

    # Exibindo as primeiras linhas do DataFrame após o merge
    print("Dados após o merge:")
    print(df_reorganizado.head())

    # Salvando o DataFrame reorganizado em um novo arquivo Excel
    try:
        df_reorganizado.to_excel(os.path.join(caminho_pasta, 'Dados_Organizados_MMs.xlsx'), index=False)
        print(f"Dados reorganizados salvos em: {os.path.join(caminho_pasta, 'Dados_Organizados_MMs.xlsx')}")
    except Exception as e:
        print(f"Erro ao salvar o arquivo: {e}")

# Chamando a função reorganizar_dados
reorganizar_dados(caminho_pasta)