import pandas as pd
import os
import glob

# Caminho da pasta onde estão os arquivos do Snickers
caminho_pasta = r'C:\\Users\\an770843\\Mars Inc\\Brazil_MW_PlanningComex - Documents\\1 - PCP\\2 - Plano de produção\\2025\\Snickers\\Compilado Automação BI'

def excel_date(num):
    """Converte um número de série do Excel para uma data do pandas."""
    return pd.to_datetime('1899-12-30') + pd.to_timedelta(num, unit='D')

def reorganizar_dados_snickers(caminho_pasta):
    print("Iniciando a reorganização dos dados do Snickers...")  # Mensagem de início
    
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
    
    # Criando uma lista para armazenar todos os dados reorganizados
    dados_reorganizados = []

    # Buscando todos os arquivos .xlsb na pasta
    arquivos = glob.glob(os.path.join(caminho_pasta, "*.xlsb"))
    
    dia_semana_mapping = {
        6: 'Dom',  # Domingo
        0: 'Seg',  # Segunda-feira
        1: 'Ter',  # Terça-feira
        2: 'Qua',  # Quarta-feira
        3: 'Qui',  # Quinta-feira
        4: 'Sex',  # Sexta-feira
        5: 'Sáb'   # Sábado
    }

    for arquivo_excel in arquivos:
        # Extraindo o nome do arquivo
        nome_arquivo = os.path.basename(arquivo_excel)
        
        try:
            # Lendo a aba '1.Heijunka' do arquivo do Snickers
            df = pd.read_excel(arquivo_excel, sheet_name='1.Heijunka', header=None)
            print(f"Arquivo '{nome_arquivo}' lido com sucesso!")
        except Exception as e:
            print(f"Erro ao ler o arquivo '{nome_arquivo}': {e}")
            continue  # Pula para o próximo arquivo em caso de erro

        # Exibindo as primeiras linhas do DataFrame original
        print("Dados originais:")
        print(df.head())

        # A primeira linha contém os períodos e semanas
        periodos_semanas = df.iloc[0, 6:].tolist()  # Pega os períodos e semanas a partir da 7ª coluna

        # A segunda linha contém as datas
        datas = df.iloc[1, 6:].tolist()  # Pega as datas a partir da 7ª coluna

        # Preenchendo datas mescladas
        ultima_data_valida = None
        for i in range(len(datas)):
            if pd.notna(datas[i]):
                ultima_data_valida = datas[i]
            else:
                datas[i] = ultima_data_valida

        # Convertendo datas de números de série para datetime
        datas = [excel_date(d) if isinstance(d, (int, float)) else d for d in datas]

        # Iterando sobre as linhas a partir da terceira linha
        for index in range(2, df.shape[0]):
            row = df.iloc[index]
            sku = row[0]  # SKU está na primeira coluna
            nome_material = row[1]  # Nome do material está na segunda coluna
            rate = row[3]  # TOC [CX/hora] está na quarta coluna

            # Ignorando linhas que contêm "Material" na coluna SKU ou que estão vazias
            if sku == "Material" or pd.isna(sku) or pd.isna(nome_material):
                continue

            # Verificando se o SKU é um número válido
            sku_valido = pd.to_numeric(sku, errors='coerce')
            if pd.isna(sku_valido):
                continue  # Pula linhas que não têm SKU numérico

            # Iterando sobre os períodos e semanas
            for i in range(len(periodos_semanas)):
                quantidade = row[i + 6]  # A quantidade está nas colunas a partir da 7ª
                quantidade = pd.to_numeric(quantidade, errors='coerce')  # Converte para numérico, substitui não numéricos por NaN
                
                if pd.notna(quantidade) and quantidade > 0:  # Verifica se a quantidade é válida
                    # Preenchendo a data corretamente
                    data_producao = datas[i]
                    
                    # Obtendo o dia da semana
                    dia_semana = dia_semana_mapping[data_producao.weekday()]
                    
                    # Obtendo o número do dia da semana (1 a 7)
                    numero_dia_semana = (data_producao.weekday() + 1) % 7 + 1

                    # Adiciona uma linha com a quantidade correta
                    dados_reorganizados.append({
                        'Máquina': 'Snickers',  # Nova coluna adicionada
                        'SKU': sku,
                        'Rate': rate,  # Nova coluna "Rate" adicionada
                        'Data': dia_semana,  # Nova coluna "Dia da Semana" adicionada
                        'Quantidade': quantidade,  # Usando a quantidade real
                        'Período': periodos_semanas[i][:3],  # Pega os 3 primeiros caracteres para o período
                        'Week': periodos_semanas[i][3:],      # Pega o restante para a semana
                        'Ano': ano,  # Nova coluna "Ano" adicionada
                        'Dia Semana': numero_dia_semana,  # Nova coluna "Número do Dia da Semana" adicionada
                        'P/W/Y/D': periodos_semanas[i][:3] + periodos_semanas[i][3:] + ano + str(numero_dia_semana),
                        #'Nome do material': nome_material,
                        'DataFormat': data_producao,  # Data correspondente

                   
                    })

        # Criando um DataFrame a partir dos dados reorganizados
        df_reorganizado = pd.DataFrame(dados_reorganizados)

        # Exibindo as primeiras linhas do DataFrame reorganizado
        print("Dados reorganizados:")
        print(df_reorganizado.head())

        # Verificando se há dados reorganizados
        if df_reorganizado.empty:
            print("Nenhum dado válido foi encontrado para reorganização.")
            return

        # Salvando o DataFrame reorganizado em um novo arquivo Excel
        try:
            df_reorganizado.to_excel(os.path.join(os.path.dirname(caminho_pasta), 'Dados_Organizados_Snickers.xlsx'), index=False)
            print(f"Dados reorganizados salvos em: {os.path.join(os.path.dirname(caminho_pasta), 'Dados_Organizados_Snickers.xlsx')}")
        except Exception as e:
            print(f"Erro ao salvar o arquivo: {e}")

# Chamando a função reorganizar_dados_snickers
reorganizar_dados_snickers(caminho_pasta)