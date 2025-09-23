import os
import re
from pathlib import Path

# Importações para PDF
try:
    import PyPDF2
    import pdfplumber
except ImportError:
    print("ERRO: Instale as dependências primeiro!")
    print("Execute: pip install PyPDF2 pdfplumber")
    exit()

def extrair_nomes(texto_pdf):
    """Extrai os nomes do locador e locatário e a data de um texto de contrato."""
    
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
    
    # Formata os nomes para o nome do arquivo
    nome_locador_formatado = " ".join(locador.split()[:2])
    nome_locatario_formatado = " ".join(locatario.split()[:2])
    
    return nome_locador_formatado, nome_locatario_formatado, data

def main():
    """Função principal que orquestra o programa."""
    print("🚀 RENOMEADOR AUTOMÁTICO DE ARQUIVOS")
    print("=" * 40)
    
    # Define a pasta de origem dos arquivos
    # Mude este caminho se os arquivos de origem estiverem em outra pasta
    pasta_origem = Path.cwd()
    
    # Define a pasta de destino para os arquivos renomeados
    # ESTE É O NOVO CAMINHO QUE VOCÊ PEDIU
    pasta_destino = Path(r"\\SRV_IMOBILIARIA\imobiliaria\0_Estagiario\V1 python\Testes CLs")
    
    # Certifica-se de que a pasta de destino existe
    try:
        pasta_destino.mkdir(parents=True, exist_ok=True)
    except Exception as e:
        print(f"❌ Erro ao criar/acessar a pasta de destino: {e}")
        return

    # Processa cada arquivo na pasta de origem
    for nome_arquivo in os.listdir(pasta_origem):
        if nome_arquivo.endswith(('.pdf', '.docx')):
            caminho_arquivo_original = pasta_origem / nome_arquivo
            
            print(f"\nProcessando arquivo: {nome_arquivo}")

            texto_completo = ""
            
            # Extrai texto do PDF
            if nome_arquivo.endswith('.pdf'):
                try:
                    with pdfplumber.open(caminho_arquivo_original) as pdf:
                        for page in pdf.pages:
                            texto_completo += page.extract_text() or ""
                except Exception as e:
                    print(f"❌ Erro ao extrair texto do PDF: {e}")
                    continue
            
            # Extrai texto do DOCX (aqui você precisaria de uma biblioteca para .docx)
            elif nome_arquivo.endswith('.docx'):
                print("⚠️ A extração de texto de arquivos DOCX não foi implementada neste script.")
                continue

            # Extrai os nomes e a data do texto
            locador, locatario, data = extrair_nomes(texto_completo)

            # Cria o novo nome do arquivo
            novo_nome = f"cl {locador} X {locatario}.pdf"
            caminho_destino = pasta_destino / novo_nome
            
            # Renomeia o arquivo
            try:
                os.rename(caminho_arquivo_original, caminho_destino)
                print(f"✅ Renomeado e movido para: {caminho_destino}")
            except Exception as e:
                print(f"❌ Erro ao renomear/mover o arquivo: {e}")
    
    print("\n✅ Processamento concluído!")

if __name__ == "__main__":
    main()