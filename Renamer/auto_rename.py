import os
import re
from pathlib import Path
import shutil

# Importações para PDF e DOCX
try:
    import PyPDF2
    import pdfplumber
    import docx
except ImportError:
    print("ERRO: Instale as dependências primeiro!")
    print("Execute: pip install PyPDF2 pdfplumber python-docx")
    exit()

def extrair_nomes_e_data(texto_documento):
    """Extrai os nomes do locador e locatário e a data de um texto de contrato."""
    
    # Padrões para extração de nomes - AGORA ACEITA SINGULAR E PLURAL
    padrao_locador = r'LOCADOR(?:ES)?(?:\s*\(A\))?:\s*(.*?)(?:,|;|\n)'
    padrao_locatario = r'LOCATÁRIO(?:S)?(?:\s*\(A\))?:\s*(.*?)(?:,|;|\n)'
    
    locador_match = re.search(padrao_locador, texto_documento, re.IGNORECASE)
    locatario_match = re.search(padrao_locatario, texto_documento, re.IGNORECASE)

    locador = locador_match.group(1).strip().title() if locador_match else "Nao encontrado"
    locatario = locatario_match.group(1).strip().title() if locatario_match else "Nao encontrado"
    
    # Padrão para extração de data
    padrao_data = r"(\d+\s+de\s+\w+\s+de\s+\d{4})"
    data_match = re.search(padrao_data, texto_documento)
    data = data_match.group(1).strip() if data_match else "Nao encontrada"
    
    return locador, locatario, data

def extrair_texto_docx(caminho_arquivo):
    """Extrai o texto completo de um documento DOCX."""
    try:
        doc = docx.Document(caminho_arquivo)
        texto_completo = ""
        for paragraph in doc.paragraphs:
            texto_completo += paragraph.text + " "
        return texto_completo
    except Exception as e:
        print(f"❌ Erro ao extrair texto de DOCX: {e}")
        return None

def main():
    """Função principal que orquestra o programa."""
    print("🚀 RENOMEADOR AUTOMÁTICO DE ARQUIVOS")
    print("=" * 40)
    
    # Define a pasta onde estão as subpastas com os contratos.
    # O programa irá entrar em cada subpasta para processar os arquivos.
    pasta_principal = Path(r"\\SRV_IMOBILIARIA\imobiliaria\0_Estagiario\V1python\TestesCLs")

    # Processa cada item na pasta principal
    for nome_item in os.listdir(pasta_principal):
        caminho_item = pasta_principal / nome_item
        
        # Se for uma pasta, entra nela para processar os arquivos
        if caminho_item.is_dir():
            print(f"\n📁 Entrando na pasta: {caminho_item}")
            
            for nome_arquivo in os.listdir(caminho_item):
                if nome_arquivo.endswith(('.pdf', '.docx')):
                    caminho_arquivo_original = caminho_item / nome_arquivo
                    
                    print(f"\nProcessando arquivo: {caminho_arquivo_original}")

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
                    
                    # Extrai texto do DOCX
                    elif nome_arquivo.endswith('.docx'):
                        texto_completo = extrair_texto_docx(caminho_arquivo_original)
                        if texto_completo is None:
                            continue
                    
                    # Extrai os nomes e a data do texto
                    locador, locatario, data = extrair_nomes_e_data(texto_completo)

                    # Cria o nome base do novo arquivo com "CL" em maiúsculas
                    novo_nome_base = f"CL {locador} X {locatario} {data.replace(' ', '-')}.pdf"
                    caminho_destino_final = caminho_item / novo_nome_base
                    
                    # Verifica se o arquivo já existe e adiciona um número se necessário
                    contador = 1
                    while caminho_destino_final.exists():
                        novo_nome_com_contador = f"CL {locador} X {locatario} ({contador}).pdf"
                        caminho_destino_final = caminho_item / novo_nome_com_contador
                        contador += 1
                    
                    # Renomeia o arquivo
                    try:
                        os.rename(caminho_arquivo_original, caminho_destino_final)
                        print(f"✅ Renomeado e movido para: {caminho_destino_final}")
                    except Exception as e:
                        print(f"❌ Erro ao renomear/mover o arquivo: {e}")
        
    print("\n✅ Processamento concluído!")

if __name__ == "__main__":
    main()
