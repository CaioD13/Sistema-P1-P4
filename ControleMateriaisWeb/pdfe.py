import pdfplumber
import pandas as pd
import re
import os

def limpar_caminho(caminho):
    return caminho.strip().replace('"', '').replace("'", "")

def limpar_numero_espacado(texto):
    """Transforma '2 4 5 2 20' em '245220'"""
    if not texto: return ""
    # Remove espaços apenas se for composto majoritariamente por dígitos
    clean = texto.replace(" ", "").replace(".", "")
    return clean

def solicitar_dados():
    while True:
        pdf_input = input("\n📂 Caminho do PDF: ")
        pdf_input = limpar_caminho(pdf_input)
        if os.path.exists(pdf_input): break
        print("❌ Arquivo não encontrado.")

    out_input = input("\n💾 Nome do Excel (Ex: armas.xlsx): ")
    out_input = limpar_caminho(out_input)
    if not out_input.lower().endswith('.xlsx'): out_input += '.xlsx'

    while True:
        try:
            page_input = int(input("\n📄 A partir de qual página começar? (Ex: 114): "))
            if page_input > 0: break
        except: pass
    
    return pdf_input, out_input, page_input

# Regex que busca apenas o VALOR como âncora (R$ 0.000,00)
REGEX_VALOR = re.compile(r"(\d{1,3}(?:\.\d{3})*,\d{2})")

def processar_pdf():
    pdf_path, excel_path, pagina_inicial = solicitar_dados()
    print(f"\n🚀 Iniciando extração (Estratégia: Ancoragem por Valor)...")

    data_rows = []
    lines_debug = [] # Para mostrar ao usuário caso falhe

    with pdfplumber.open(pdf_path) as pdf:
        total = len(pdf.pages)
        if pagina_inicial > total: return

        for i in range(pagina_inicial - 1, total):
            page = pdf.pages[i]
            text = page.extract_text()
            if not text: continue

            lines = text.split('\n')
            
            for line in lines:
                # Pula linhas irrelevantes
                if "Listagem de" in line or "Página" in line or "Valor Total" in line:
                    continue

                # 1. Tenta achar o valor monetário na linha
                match_valor = REGEX_VALOR.search(line)
                
                if match_valor:
                    valor = match_valor.group(1)
                    # Divide a linha em ANTES e DEPOIS do valor
                    partes = line.split(valor)
                    
                    lado_esquerdo = partes[0].strip() # Cod, Pat, Serie, Desc
                    lado_direito = partes[1].strip()  # UGE, NE, NL, Conta, Fab

                    try:
                        # --- PROCESSAMENTO DO LADO ESQUERDO ---
                        # Espera-se: Codigo(num/space)  Patrimonio(num/space)  Serie  Descricao
                        # Vamos dividir por "pelo menos 2 espaços" para separar colunas grudadas
                        # Se não funcionar, divide por 1 espaço e tenta remontar
                        
                        # Estratégia: Pegar as 3 primeiras "palavras" como Cod, Pat, Serie
                        tokens_esq = lado_esquerdo.split()
                        
                        # A descrição é tudo que sobra depois dos 3 primeiros tokens (ou mais, dependendo da quebra)
                        # O problema é o Codigo vir como "2 4 5". 
                        # Vamos usar regex no inicio da linha para pegar numeros espaçados
                        
                        match_inicio = re.match(r"^([\d\s\.]+)\s+([\d\s\.]+)\s+([^\s]+)\s+(.*)$", lado_esquerdo)
                        
                        if match_inicio:
                            cod_raw = match_inicio.group(1)
                            pat_raw = match_inicio.group(2)
                            serie = match_inicio.group(3)
                            desc = match_inicio.group(4)
                        else:
                            # Fallback simples: Pega tokens
                            # Assume que Cod e Pat podem estar fragmentados, mas Serie é unida
                            # Isso é arriscado, vamos tentar limpar o codigo depois
                            desc = lado_esquerdo # Se falhar, joga tudo na descrição para ver no excel
                            cod_raw = "ERRO"
                            pat_raw = "ERRO"
                            serie = "ERRO"

                        # --- PROCESSAMENTO DO LADO DIREITO ---
                        # Espera-se: UGE  NE  NL  Conta  Fabricante
                        tokens_dir = lado_direito.split()
                        
                        if len(tokens_dir) >= 4:
                            uge = tokens_dir[0]
                            ne = tokens_dir[1]
                            nl = tokens_dir[2]
                            conta = tokens_dir[3]
                            # Fabricante é tudo que sobrar (pode ter espaços, ex: TAURUS ARMAS)
                            fab = " ".join(tokens_dir[4:])
                        else:
                            # Caso falte campo, preenche o que der
                            uge = tokens_dir[0] if len(tokens_dir) > 0 else ""
                            ne = tokens_dir[1] if len(tokens_dir) > 1 else ""
                            nl = ""
                            conta = ""
                            fab = " ".join(tokens_dir[2:]) if len(tokens_dir) > 2 else ""

                        # Adiciona na lista final com limpeza
                        data_rows.append({
                            "Código Material": limpar_numero_espacado(cod_raw),
                            "Nº Patrimônio": limpar_numero_espacado(pat_raw),
                            "Nº Série": serie,
                            "Descrição": desc,
                            "Valor (R$)": valor,
                            "UGE": uge,
                            "NE": ne,
                            "NL": nl,
                            "Conta Pat": conta,
                            "Fabricante": fab,
                            "_DEBUG_RAW": line[:50] # Coluna oculta para conferência
                        })
                    except Exception as e:
                        continue # Pula linha com erro de parse
                else:
                    # Guarda as primeiras 5 linhas rejeitadas para mostrar ao usuario
                    if len(lines_debug) < 5 and len(line) > 20:
                        lines_debug.append(line)

            print(f"Lendo Pág {i+1}/{total} | Itens: {len(data_rows)}", end='\r')

    print(f"\n\n✅ Finalizado! {len(data_rows)} registros.")
    
    if len(data_rows) == 0:
        print("\n⚠️ DIAGNÓSTICO DE ERRO:")
        print("O script não identificou o padrão de valor (0,00) ou a estrutura da linha.")
        print("Abaixo estão exemplos de linhas que o PDF leu (copie e me mande):")
        print("-" * 50)
        for l in lines_debug:
            print(f"LINHA LIDA: {l}")
        print("-" * 50)
    
    if data_rows:
        df = pd.DataFrame(data_rows)
        # Remove a coluna de debug antes de salvar
        if "_DEBUG_RAW" in df.columns: del df["_DEBUG_RAW"]
        df.to_excel(excel_path, index=False)
        print(f"💾 Salvo em: {os.path.abspath(excel_path)}")

if __name__ == "__main__":
    processar_pdf()