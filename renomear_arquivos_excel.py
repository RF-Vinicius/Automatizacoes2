import pandas as pd
import os
import shutil

def renomear_arquivos(caminho_planilha, caminho_pasta_arquivos, coluna_chave, coluna_valor):
  

    caminho_pasta_nao_encontrados = os.path.join(caminho_pasta_arquivos, "NaoEncontrados")
    os.makedirs(caminho_pasta_nao_encontrados, exist_ok=True)

    # --- 2. LER A PLANILHA EXCEL ---
    try:
        df = pd.read_excel(caminho_planilha)
        print(f"Planilha '{os.path.basename(caminho_planilha)}' lida com sucesso.")
    except FileNotFoundError:
        print(f"ERRO: A planilha '{caminho_planilha}' não foi encontrada.")
        return
    except Exception as e:
        print(f"ERRO ao ler a planilha: {e}")
        return

    # Verificar se as colunas necessárias existem na planilha
    if coluna_chave not in df.columns or coluna_valor not in df.columns:
        print(f"ERRO: As colunas '{coluna_chave}' e/ou '{coluna_valor}' não foram encontradas na planilha.")
        return

    # O .astype(str) garante que a comparação seja feita com strings
    mapeamento_nomes = df.set_index(coluna_chave)[coluna_valor].astype(str).to_dict()


    # Percorrer todos os arquivos na pasta de origem
    for nome_arquivo in os.listdir(caminho_pasta_arquivos):
        if os.path.isdir(os.path.join(caminho_pasta_arquivos, nome_arquivo)):
            continue

        nome_base, extensao = os.path.splitext(nome_arquivo)

       
        if nome_base in mapeamento_nomes:
            novo_nome_base = mapeamento_nomes[nome_base]
            novo_nome_completo = novo_nome_base + extensao
            
            
            if nome_arquivo == novo_nome_completo:
                print(f"Arquivo '{nome_arquivo}' já tem o nome correto. Ignorando.")
                continue

            caminho_antigo = os.path.join(caminho_pasta_arquivos, nome_arquivo)
            caminho_novo = os.path.join(caminho_pasta_arquivos, novo_nome_completo)

            try:
                os.rename(caminho_antigo, caminho_novo)
                print(f"Renomeado: '{nome_arquivo}' -> '{novo_nome_completo}'")
            except FileExistsError:
                print(f"ALERTA: O arquivo '{novo_nome_completo}' já existe. Não foi possível renomear '{nome_arquivo}'.")
            except Exception as e:
                print(f"ERRO ao renomear o arquivo '{nome_arquivo}': {e}")
        else:
            # O arquivo não foi encontrado na planilha, movê-lo para a pasta de não encontrados
            caminho_origem = os.path.join(caminho_pasta_arquivos, nome_arquivo)
            caminho_destino = os.path.join(caminho_pasta_nao_encontrados, nome_arquivo)
            
            try:
                shutil.move(caminho_origem, caminho_destino)
                print(f"↔ Movido: '{nome_arquivo}' -> '{caminho_pasta_nao_encontrados}'")
            except Exception as e:
                print(f"ERRO ao mover o arquivo '{nome_arquivo}': {e}")


if __name__ == "__main__":
    # Configure os caminhos e nomes das colunas aqui
    CAMINHO_PLANILHA = "C:/Users/vini9/OneDrive/Área de Trabalho/ConstruCode/Automatizacoes/planilha_renomear.xlsx"
    CAMINHO_PASTA_ARQUIVOS = "C:/Users/vini9/OneDrive/Área de Trabalho/ConstruCode/Automatizacoes/Arquivos"
    
    # Nomes das colunas na sua planilha Excel
    COLUNA_CHAVE = "chave"
    COLUNA_VALOR = "valor"

    renomear_arquivos(CAMINHO_PLANILHA, CAMINHO_PASTA_ARQUIVOS, COLUNA_CHAVE, COLUNA_VALOR)