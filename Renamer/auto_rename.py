import os
import re
from pathlib import Path

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
    
    # Padrões para extração de nomes
    padrao_locador = r'LOCADOR(?:\s*\(A\))?:\s*(.*?)(?:,|;|\n)'
    padrao_locatario = r'LOCATÁRIO(?:\s*\(A\))?:\s*(.*?)(?:,|;|\n)'
    
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
    
    # Define a pasta de origem dos arquivos
    # O programa irá vasculhar todas as subpastas dentro desta pasta
    pasta_origem = Path.cwd()
    
    # Define a pasta de destino para os arquivos renomeados
    pasta_destino = Path(r"\\SRV_IMOBILIARIA\imobiliaria\0_Estagiario\V1 python\Testes CLs")
    
    # Certifica-se de que a pasta de destino existe
    try:
        pasta_destino.mkdir(parents=True, exist_ok=True)
    except Exception as e:
        print(f"❌ Erro ao criar/acessar a pasta de destino: {e}")
        return

    # Percorre a pasta de origem e todas as subpastas
    for root, dirs, files in os.walk(pasta_origem):
        for nome_arquivo in files:
            if nome_arquivo.endswith(('.pdf', '.docx')):
                caminho_arquivo_original = Path(root) / nome_arquivo
                
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

                # Cria o novo nome do arquivo
                novo_nome = f"cl {locador} X {locatario}.pdf"
                caminho_destino_final = pasta_destino / novo_nome
                
                # Renomeia e move o arquivo
                try:
                    os.rename(caminho_arquivo_original, caminho_destino_final)
                    print(f"✅ Renomeado e movido para: {caminho_destino_final}")
                except Exception as e:
                    print(f"❌ Erro ao renomear/mover o arquivo: {e}")
    
    print("\n✅ Processamento concluído!")

if __name__ == "__main__":
    main()
