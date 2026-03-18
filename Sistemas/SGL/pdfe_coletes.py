import pdfplumber
import pandas as pd
import re
import os

def limpar_caminho(caminho):
    return caminho.strip().replace('"', '').replace("'", "")

def limpar_numero_espacado(texto):
    if not texto: return ""
    return texto.replace(" ", "").replace(".", "").strip()

def solicitar_dados():
    while True:
        pdf_input = input("\n📂 Caminho do PDF: ")
        pdf_input = limpar_caminho(pdf_input)
        if os.path.exists(pdf_input): break
        print("❌ Arquivo não encontrado.")

    out_input = input("\n💾 Nome do Excel (Ex: coletes.xlsx): ")
    out_input = limpar_caminho(out_input)
    if not out_input.lower().endswith('.xlsx'): out_input += '.xlsx'

    while True:
        try:
            page_input = int(input("\n📄 A partir de qual página começar? (Ex: 114): "))
            if page_input > 0: break
        except: pass
    
    return pdf_input, out_input, page_input

# REGEX 1: Busca o Valor Numérico (ex: 1.200,00)
REGEX_VALOR = re.compile(r"(\d{1,3}(?:\.\d{3})*,\d{2})")

# REGEX 2: Lado Esquerdo (CORRIGIDO PARA IGNORAR "R$" NO FINAL)
# Grupo 1 (Cod): Digitos/Espaços
# Grupo 2 (Pat): Digitos/Espaços
# Grupo 3 (Serie): Texto sem espaço
# Grupo 4 (Fab): Fabricante (Meio)
# Grupo 5 (Sexo): M ou U
# Grupo 6 (Valid): Data
# FINAL (.*$): Ignora o "R$" ou espaços que sobrem no final
REGEX_ESQUERDA = re.compile(r"^([\d\s\.]+)\s+([\d\s\.]+)\s+([^\s]+)\s+(.+?)\s+([MU])\s+(\d{1,2}-[A-Za-zç]{3}-\d{4}).*$", re.IGNORECASE)

def processar_pdf():
    pdf_path, excel_path, pagina_inicial = solicitar_dados()
    print(f"\n🚀 Iniciando extração...")

    data_rows = []
    lines_debug = []

    with pdfplumber.open(pdf_path) as pdf:
        total = len(pdf.pages)
        if pagina_inicial > total: return

        for i in range(pagina_inicial - 1, total):
            page = pdf.pages[i]
            text = page.extract_text()
            if not text: continue

            lines = text.split('\n')
            
            for line in lines:
                # Ignora cabeçalhos
                if "Listagem de" in line or "Cod Mat" in line or "Valor Total" in line:
                    continue

                # 1. Acha o VALOR NUMÉRICO para dividir a linha
                match_valor = REGEX_VALOR.search(line)
                
                if match_valor:
                    valor = match_valor.group(1)
                    try:
                        partes = line.split(valor)
                        lado_esquerdo = partes[0].strip() # Cod ... Data R$
                        lado_direito = partes[1].strip()  # NE ...

                        # 2. Processa Lado Esquerdo (Cod até Data)
                        match_esq = REGEX_ESQUERDA.match(lado_esquerdo)

                        if match_esq:
                            cod_raw = match_esq.group(1)
                            pat_raw = match_esq.group(2)
                            serie = match_esq.group(3)
                            fabricante = match_esq.group(4).strip()
                            sexo = match_esq.group(5).upper()
                            validade = match_esq.group(6)
                        else:
                            lines_debug.append(f"FALHA REGEX: {lado_esquerdo}")
                            continue

                        # 3. Processa Lado Direito (NE, NL, Conta)
                        tokens_dir = lado_direito.split()
                        ne = tokens_dir[0] if len(tokens_dir) > 0 else ""
                        nl = tokens_dir[1] if len(tokens_dir) > 1 else ""
                        conta = tokens_dir[2] if len(tokens_dir) > 2 else ""

                        data_rows.append({
                            "Cod Mat": limpar_numero_espacado(cod_raw),
                            "Patrimônio": limpar_numero_espacado(pat_raw),
                            "N. Série": serie,
                            "Fabricante": fabricante,
                            "Sexo": sexo,
                            "Validade": validade,
                            "Valor": valor,
                            "NE": ne,
                            "NL": nl,
                            "Conta Pat": conta
                        })

                    except:
                        lines_debug.append(f"ERRO SPLIT: {line}")
                else:
                    if len(lines_debug) < 5 and len(line) > 20: 
                        lines_debug.append(f"NO VAL: {line}")

            print(f"Processando Pág {i+1}/{total} | Registros: {len(data_rows)}", end='\r')

    print(f"\n\n✅ Finalizado! {len(data_rows)} registros encontrados.")
    
    if data_rows:
        df = pd.DataFrame(data_rows)
        df.to_excel(excel_path, index=False)
        print(f"💾 Salvo em: {os.path.abspath(excel_path)}")
    else:
        print("\n⚠️ Erros remanescentes (copie e me mande se falhar):")
        for l in lines_debug[:5]:
            print(l)

if __name__ == "__main__":
    processar_pdf()