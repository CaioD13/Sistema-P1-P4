import calendar
from datetime import datetime, date
from io import BytesIO

import xlsxwriter
from flask import Blueprint, request, redirect, url_for, flash, send_file
from flask_login import login_required, current_user

from extensions import db
from models import Policial, AlteracaoEscala
from utils.datas import formatar_data_militar, get_semana_mes
from utils.textos import obter_texto_exibicao, get_habilitacoes_texto, get_sat_texto
from utils.ordenacao import ordenar_em_agrupado, ordenar_cia_agrupado
from utils.validacoes import traduzir_para_frequencia, is_afastamento
from services.escala_service import calcular_escala_dia


frequencia_bp = Blueprint("frequencia", __name__)


@frequencia_bp.route('/exportar_excel')
@login_required
def exportar_excel():
    try:
        filtro_posto = request.args.get('filtro_posto', 'Todos')
        mes = request.args.get('mes', datetime.now().month, type=int)
        ano = request.args.get('ano', datetime.now().year, type=int)
        cia = request.args.get('cia_filtro', 'TODAS')

        u_dia = calendar.monthrange(ano, mes)[1]

        output = BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        ws = workbook.add_worksheet("Frequencia")

        ws.set_paper(9)
        ws.set_landscape()
        ws.set_margins(0.3, 0.3, 0.3, 0.3)
        ws.fit_to_pages(1, 0)

        fmt_titulo = workbook.add_format({
            'bold': True, 'font_name': 'Arial', 'font_size': 12,
            'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#D9D9D9'
        })
        fmt_header = workbook.add_format({
            'bold': True, 'font_name': 'Arial', 'font_size': 8,
            'align': 'center', 'valign': 'vcenter', 'border': 1,
            'bg_color': '#EFEFEF', 'text_wrap': True
        })
        fmt_dia_header = workbook.add_format({
            'bold': True, 'font_name': 'Arial', 'font_size': 8,
            'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#FFFFFF'
        })
        fmt_texto_centro = workbook.add_format({
            'font_name': 'Arial', 'font_size': 8, 'align': 'center',
            'valign': 'vcenter', 'border': 1
        })
        fmt_texto_esq = workbook.add_format({
            'font_name': 'Arial', 'font_size': 8, 'align': 'left',
            'valign': 'vcenter', 'border': 1, 'indent': 1
        })
        fmt_total = workbook.add_format({
            'bold': True, 'font_name': 'Arial', 'font_size': 8,
            'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#F2F2F2'
        })
        fmt_bonus = workbook.add_format({
            'bold': True, 'font_name': 'Arial', 'font_size': 8,
            'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#E6E6FF'
        })

        ws.set_column('A:A', 8)
        ws.set_column('B:B', 10)
        ws.set_column('C:C', 35)
        ws.set_column(3, 33, 2.8)
        ws.set_column(34, 35, 6)
        ws.set_column(36, 36, 7)

        titulo_texto = f'CONTROLE DE FREQUÊNCIA - {cia if cia != "TODAS" else "37º BPM/M"}'
        subtitulo = f'REFERÊNCIA: {mes:02d}/{ano}'

        ws.merge_range('A1:AK1', titulo_texto, fmt_titulo)
        ws.merge_range('A2:AK2', subtitulo, fmt_titulo)

        ws.write(2, 0, "POSTO", fmt_header)
        ws.write(2, 1, "RE", fmt_header)
        ws.write(2, 2, "NOME DE GUERRA", fmt_header)

        col_idx = 3
        for d in range(1, 32):
            if d &lt;= u_dia:
                ws.write(2, col_idx, str(d), fmt_dia_header)
            else:
                ws.write(2, col_idx, "X", fmt_header)
            col_idx += 1

        ws.write(2, col_idx, "DIAS", fmt_header)
        ws.write(2, col_idx + 1, "VALOR", fmt_header)
        ws.write(2, col_idx + 2, "BÔNUS", fmt_header)

        lista_oficiais = [
            'Coronel PM', 'Cel PM',
            'Tenente Coronel PM', 'Ten Cel PM',
            'Major PM', 'Maj PM',
            'Capitão PM', 'Cap PM',
            '1º Tenente PM', '1º Ten PM',
            '2º Tenente PM', '2º Ten PM',
            'Aspirante a Oficial', 'Asp Of PM'
        ]

        q = Policial.query.filter_by(ativo=True).order_by(Policial.nome)

        if cia != 'TODAS':
            q = q.filter_by(cia=cia)

        if filtro_posto == 'Oficiais':
            q = q.filter(Policial.posto.in_(lista_oficiais))
        elif filtro_posto == 'Pracas':
            q = q.filter(Policial.posto.notin_(lista_oficiais))

        pms = q.all()

        linha = 3
        for pm in pms:
            ws.write(linha, 0, pm.posto, fmt_texto_centro)
            ws.write(linha, 1, pm.re, fmt_texto_centro)
            ws.write(linha, 2, pm.nome_guerra or pm.nome, fmt_texto_esq)

            c = 3
            dias_trabalhados = 0
            soma_valor = 0
            bonus_count = 0

            for d in range(1, 32):
                val_cell = ""

                if d &lt;= u_dia:
                    data_dia = date(ano, mes, d)
                    info_dia = calcular_escala_dia(pm, data_dia)

                    sigla = info_dia.get('sigla')
                    duracao = info_dia.get('duracao_horas_escala')

                    freq = traduzir_para_frequencia(sigla, duracao)

                    val_cell = freq["codigo"]

                    if freq["conta_dias"]:
                        dias_trabalhados += 1
                        if freq["diarias"] is not None:
                            soma_valor += freq["diarias"]

                    if freq["codigo"] in ['0', '1', '2', '3']:
                        bonus_count += 1

                ws.write(linha, c, val_cell, fmt_texto_centro)
                c += 1

            ws.write(linha, c, dias_trabalhados, fmt_total)
            ws.write(linha, c + 1, soma_valor, fmt_total)
            ws.write(linha, c + 2, bonus_count, fmt_bonus)

            linha += 1

        linha += 3
        ws.merge_range(f'C{linha}:J{linha}', '_' * 30, fmt_texto_centro)
        ws.merge_range(f'C{linha+1}:J{linha+1}', 'ENCARREGADO DE SECRETARIA', fmt_texto_centro)
        ws.merge_range(f'R{linha}:Y{linha}', '_' * 30, fmt_texto_centro)
        ws.merge_range(f'R{linha+1}:Y{linha+1}', f'CMT DA {cia if cia != "TODAS" else "UNIDADE"}', fmt_texto_centro)

        workbook.close()
        output.seek(0)

        nome_arquivo = f"Frequencia_{cia}_{filtro_posto}_{mes}_{ano}.xlsx"
        return send_file(output, download_name=nome_arquivo, as_attachment=True)

    except Exception as e:
        flash(f"Erro ao gerar Excel: {e}", 'danger')
        return redirect(url_for('index'))