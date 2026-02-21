import xlsxwriter
import os

nome_arquivo = 'Modelo_Escala_Visual.xlsx'

workbook = xlsxwriter.Workbook(nome_arquivo)
worksheet = workbook.add_worksheet('ESCALA')

# --- FORMATOS ---
titulo_format = workbook.add_format({
    'bold': True, 'align': 'center', 'valign': 'vcenter', 'font_size': 14, 'font_name': 'Arial'
})
subtitulo_format = workbook.add_format({
    'bold': True, 'align': 'center', 'valign': 'vcenter', 'font_size': 11, 'font_name': 'Arial'
})
header_format = workbook.add_format({
    'bold': True, 'align': 'center', 'valign': 'vcenter',
    'fg_color': '#000000', 'font_color': 'white', 'border': 1, 'text_wrap': True
})

# Cabeçalho Institucional
largura_total = 15 
worksheet.merge_range(0, 0, 0, largura_total, "POLÍCIA MILITAR DO ESTADO DE SÃO PAULO", titulo_format)
worksheet.merge_range(1, 0, 1, largura_total, "COMANDO DE POLICIAMENTO DE ÁREA METROPOLITANA DEZ", subtitulo_format)
worksheet.merge_range(2, 0, 2, largura_total, "TRIGÉSIMO SÉTIMO BATALHÃO DE POLICIA MILITAR METROPOLITANO", subtitulo_format)
worksheet.merge_range(3, 0, 3, largura_total, "ESCALA DE SERVIÇO Nº 37BPMM", subtitulo_format)
worksheet.merge_range(4, 0, 4, largura_total, "PERÍODO: {PERIODO_DATA}", subtitulo_format)

linha_inicio = 6
colunas = [
    ('Equipe', 10), ('Posto/Grad', 15), ('RE', 12), ('QRA', 30), ('US', 15),
    ('CPP', 10), ('Modalidade', 15), ('Função', 20), ('Seção', 10),
    ('Seg', 8), ('Ter', 8), ('Qua', 8), ('Qui', 8), ('Sex', 8), ('Sab', 8), ('Dom', 8)
]

for i, (nome, largura) in enumerate(colunas):
    worksheet.write(linha_inicio, i, nome, header_format)
    worksheet.set_column(i, i, largura)

# --- REGRAS DE CORES (CONTRASTE ALTO) ---
cinza_claro = workbook.add_format({'bg_color': '#E0E0E0', 'border': 1})
cinza_escuro = workbook.add_format({'bg_color': '#A6A6A6', 'border': 1})

primeira_linha_dados = linha_inicio + 1 
area = f'A{primeira_linha_dados+1}:P500' 

# 1. BLOCO COMANDO (Comando, Sub, Coord Op) -> TUDO CINZA CLARO
worksheet.conditional_format(area, {'type': 'formula', 'criteria': '=$I8="Comando"', 'format': cinza_claro})
worksheet.conditional_format(area, {'type': 'formula', 'criteria': '=$I8="Subcomando"', 'format': cinza_claro})
worksheet.conditional_format(area, {'type': 'formula', 'criteria': '=$I8="Coord Op"', 'format': cinza_claro})

# 2. BLOCO P1 (P1, P1/P5) -> ESCURO
worksheet.conditional_format(area, {'type': 'formula', 'criteria': '=$I8="P1"', 'format': cinza_escuro})
worksheet.conditional_format(area, {'type': 'formula', 'criteria': '=$I8="P1/P5"', 'format': cinza_escuro})

# 3. BLOCO P2 -> CLARO
worksheet.conditional_format(area, {'type': 'formula', 'criteria': '=$I8="P2"', 'format': cinza_claro})

# 4. BLOCO P3 -> ESCURO
worksheet.conditional_format(area, {'type': 'formula', 'criteria': '=$I8="P3"', 'format': cinza_escuro})

# 5. BLOCO P4 -> CLARO
worksheet.conditional_format(area, {'type': 'formula', 'criteria': '=$I8="P4"', 'format': cinza_claro})

# 6. BLOCO SPJMD -> ESCURO
worksheet.conditional_format(area, {'type': 'formula', 'criteria': '=$I8="SPJMD"', 'format': cinza_escuro})

# 7. BLOCO CFP -> CLARO
worksheet.conditional_format(area, {'type': 'formula', 'criteria': '=$I8="CFP"', 'format': cinza_claro})

# --- EQUIPES (OPERACIONAL) ---
# Equipes B e D -> Escuro
worksheet.conditional_format(area, {'type': 'formula', 'criteria': '=$A8="B"', 'format': cinza_escuro})
worksheet.conditional_format(area, {'type': 'formula', 'criteria': '=$A8="D"', 'format': cinza_escuro})

# Equipes A e C -> Claro
worksheet.conditional_format(area, {'type': 'formula', 'criteria': '=$A8="A"', 'format': cinza_claro})
worksheet.conditional_format(area, {'type': 'formula', 'criteria': '=$A8="C"', 'format': cinza_claro})

# Borda padrão para o resto
worksheet.conditional_format(area, {'type': 'no_blanks', 'format': cinza_claro})

worksheet_dados = workbook.add_worksheet('DADOS')
worksheet_dados.hide()

workbook.close()
print(f"✅ Template atualizado! Comando unificado no Cinza Claro.")