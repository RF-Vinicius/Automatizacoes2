import pandas as pd
import os
import shutil
import re

def renomear_revisao_docs(caminho_doc):
    
    caminho_pasta_nao_encontrados = os.path.join(caminho_doc, "NaoEncontradosRevisao")
    os.makedirs(caminho_pasta_nao_encontrados, exist_ok=True)

    padrao01 = r"(R)(\d{2})"
    padrao02 = r"(V)(\d{2})"
    padrao02_1 = r"(VR)(\d{2})"
    padrao03 = r"MN_COMMENTS" #Ajustei aqui para - por conta do replace

    # Percorrer todos os arquivos na pasta de origem
    for nome_arquivo in os.listdir(caminho_doc):
        if os.path.isdir(os.path.join(caminho_doc, nome_arquivo)):
            continue
        
        nome_base, extensao = os.path.splitext(nome_arquivo)
        #nome_base = str(nome_base).replace("_", "-")
        listNome = str(nome_base).split("-")
        
        
        if (re.search(padrao02, "".join(listNome[-2:]))) or (re.search(padrao02_1, "".join(listNome[-2:]))) : #Caso 02
            novo_nome = listNome
            novo_nome[-2] = novo_nome[-2].replace("_", "-")
            novo_nome_documento = "-".join(novo_nome[:-2]) + "-" + "".join(novo_nome[-2:]).replace("R", "") + extensao
            #print(novo_nome_documento)     
        elif (re.search(padrao03, nome_base)): #Caso 03
            novo_nome = nome_base.replace("_MN_COMMENTS", "")
            novo_nome = novo_nome.replace("_", "-").split('-')
            novo_nome.pop(8)
            novo_nome_documento = "-".join(novo_nome[:-2]) + "-" + "".join(novo_nome[-2:]).replace("VR", "X") + extensao
           #print (novo_nome_documento)
        elif (re.search(padrao01, listNome[-1])) and not (re.search(padrao03, nome_base)): #Caso 01
            novo_nome = listNome
            novo_nome[-1] = novo_nome[-1].replace("R","E")   
            novo_nome_documento = "-".join(novo_nome) + extensao
            #print(novo_nome_documento)
        else:
            # O arquivo não foi encontrado na planilha, movê-lo para a pasta de não encontrados
            caminho_origem = os.path.join(caminho_doc, nome_arquivo)
            caminho_destino = os.path.join(caminho_pasta_nao_encontrados, nome_arquivo)
            
            try:
                shutil.move(caminho_origem, caminho_destino)
                print(f"↔ Movido: '{nome_arquivo}' -> '{caminho_pasta_nao_encontrados}'")
            except Exception as e:
                print(f"ERRO ao mover o arquivo '{nome_arquivo}': {e}")
            continue

        caminho_antigo = os.path.join(caminho_doc, nome_arquivo)
        caminho_novo = os.path.join(caminho_doc, novo_nome_documento)

        try:
            os.rename(caminho_antigo, caminho_novo)
            print(f"'{nome_arquivo}' -> '{novo_nome_documento}'")
        except FileExistsError:
            print(f"ALERTA: O arquivo '{novo_nome_documento}' já existe. Não foi possível renomear '{nome_arquivo}'.")
        except Exception as e:
            print(f"ERRO ao renomear o arquivo '{nome_arquivo}': {e}")

if __name__ == "__main__":

    caminho_doc = r"C:\Users\vini9\OneDrive\Área de Trabalho\ConstruCode\Automatizacoes2\teste"

    renomear_revisao_docs(caminho_doc)
        