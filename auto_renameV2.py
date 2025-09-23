import os
import re
from pathlib import Path

try:
    import PyPDF2
    import pdfplumber
except ImportError:
    print("ERRO: Instale as dependências primeiro!")
    print("Execute: pip install PyPDF2 pdfplumber")
    exit()

def extrair_nomes(texto_pdf):
    
    # Nomes
    padrao_locador = r'LOCADOR\s*\(\w*\):\s*(.*?),'
    padrao_locatario = r'LOCATÁRIO\s*\(\w*\):\s*(.*?),'
    
    locador_match = re.search(padrao_locador, texto_pdf, re.IGNORECASE)
    locatario_match = re.search(padrao_locatario, texto_pdf, re.IGNORECASE)

    locador = locador_match.group(1).strip() if locador_match else "Nao encontrado"
    locatario = locatario_match.group(1).strip() if locatario_match else "Nao encontrado"
    
    # Data
    padrao_data = r"Can\w+\s*,\s*(\d+\s+de\s+\w+\s+de\s+\d{4})"
    data_match = re.search(padrao_data, texto_pdf)
    data = data_match.group(1).strip() if data_match else "Nao encontrada"
    
    nome_locador_formatado = " ".join(locador.split()[:2])
    nome_locatario_formatado = " ".join(locatario.split()[:2])
    
    return nome_locador_formatado, nome_locatario_formatado, data

def main():
    """Função principal que orquestra o programa."""
    print("🚀 RENOMEADOR AUTOMÁTICO DE ARQUIVOS")
    print("=" * 40)
    
    pasta_origem = Path.cwd() # MUDAR CASO PRECISE
    
    pasta_destino = Path(r"\\C:") # EDITAR DEPENDENDO DO AMBIENTE
    
    # TRY/EXCEPT PRA VER SE PASTA EXISTE
    try:
        pasta_destino.mkdir(parents=True, exist_ok=True)
    except Exception as e:
        print(f"❌ Erro ao criar/acessar a pasta de destino: {e}")
        return

    # PROCESSA TODOS ARQUIVOS NA PASTA
    for nome_arquivo in os.listdir(pasta_origem):
        if nome_arquivo.endswith(('.pdf', '.docx')):
            caminho_arquivo_original = pasta_origem / nome_arquivo
            
            print(f"\nProcessando arquivo: {nome_arquivo}")

            texto_completo = ""
            
            if nome_arquivo.endswith('.pdf'): # EXTRAIR TEXTO DO PDF
                try:
                    with pdfplumber.open(caminho_arquivo_original) as pdf:
                        for page in pdf.pages:
                            texto_completo += page.extract_text() or ""
                except Exception as e:
                    print(f"❌ Erro ao extrair texto do PDF: {e}")
                    continue
            
            elif nome_arquivo.endswith('.docx'): # ADD FUTURAMENTE TEXTO DE DOCX
                print("⚠️ A extração de texto de arquivos DOCX não foi implementada neste script.")
                continue

            locador, locatario, data = extrair_nomes(texto_completo) # EXTRAI NOME E DATA

            novo_nome = f"cl {locador} X {locatario}.pdf"
            caminho_destino = pasta_destino / novo_nome
            
            try:
                os.rename(caminho_arquivo_original, caminho_destino)
                print(f"✅ Renomeado e movido para: {caminho_destino}")
            except Exception as e:
                print(f"❌ Erro ao renomear/mover o arquivo: {e}")
    
    print("\n✅ Processamento concluído!")

if __name__ == "__main__":

    main()

