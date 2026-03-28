from datetime import datetime, timedelta

from flask import Blueprint, render_template, request, redirect, url_for, flash
from flask_login import login_required, current_user
from sqlalchemy import desc

from models import Policial, AlteracaoEscala, LogAuditoria
from extensions import db
from services.escala_service import (
    calcular_escala_dia,
    filtrar_por_tipo_hierarquico,
    agrupar_por_secao,
    ordem_blocos_para_tipo,
    formatar_horario_turno,
    get_proxima_versao
)
from utils.ordenacao import ordenar_por_patente, ordenar_em_agrupado, ordenar_cia_agrupado
from utils.textos import obter_texto_exibicao
from utils.validacoes import is_afastamento

escala_bp = Blueprint('escala', __name__, url_prefix='/escala')


@escala_bp.route('/visualizar')
@login_required
def visualizar_escala():
    cia = request.args.get('cia', 'EM')
    tipo_filtro = request.args.get('tipo_filtro', 'TODOS')
    data_str = request.args.get('data_inicio')

    if data_str:
        inicio = datetime.strptime(data_str, '%Y-%m-%d').date()
    else:
        hoje = datetime.now().date()
        inicio = hoje - timedelta(days=hoje.weekday())

    return visualizar_escala_semanal(cia, inicio, tipo_filtro)


@escala_bp.route('/visualizar/semanal')
@login_required
def visualizar_escala_semanal(cia=None, inicio=None, tipo_filtro=None):
    if cia is None:
        cia = request.args.get('cia', 'EM')
    if tipo_filtro is None:
        tipo_filtro = request.args.get('tipo_filtro', 'TODOS')
    if inicio is None:
        data_str = request.args.get('data_inicio')
        if data_str:
            inicio = datetime.strptime(data_str, '%Y-%m-%d').date()
        else:
            hoje = datetime.now().date()
            inicio = hoje - timedelta(days=hoje.weekday())

    try:
        fim = inicio + timedelta(days=6)
        versao_txt = get_proxima_versao(cia, inicio, tipo_filtro)

        tipo_txt = (
            "OFICIAIS" if tipo_filtro == 'OFICIAIS'
            else "PRAÇAS" if tipo_filtro == 'PRACAS'
            else "CFP" if tipo_filtro == 'CFP'
            else "POLICIAIS MILITARES"
        )

        numerador = f"{(inicio.day - 1) // 7 + 1:02d}.{inicio.month}/{str(inicio.year)[2:]}"
        periodo_txt = f"DE {inicio.strftime('%d%b%y').upper()} A {fim.strftime('%d%b%y').upper()}"

        pms_raw = Policial.query.filter_by(ativo=True).all()
        if cia != 'TODAS':
            pms_raw = [pm for pm in pms_raw if pm.cia == cia]

        pms = filtrar_por_tipo_hierarquico(pms_raw, tipo_filtro)

        if tipo_filtro in ('OFICIAIS', 'CFP'):
            pms = ordenar_por_patente(pms)
        elif cia == 'EM':
            pms = ordenar_em_agrupado(pms)
        else:
            pms = ordenar_cia_agrupado(pms)

        mapa = agrupar_por_secao(pms, tipo_filtro)
        ordem = ordem_blocos_para_tipo(tipo_filtro)

        dias_semana = []
        nomes_dias = ['SEG', 'TER', 'QUA', 'QUI', 'SEX', 'SAB', 'DOM']
        for i in range(7):
            d = inicio + timedelta(days=i)
            dias_semana.append({
                'dia_semana': nomes_dias[i],
                'data_str': f"{d.day:02d}/{d.month:02d}",
                'data': d
            })

        blocos = []
        for nome_bloco in ordem:
            lista_pms = mapa.get(nome_bloco, [])
            if not lista_pms:
                continue

            horario_turno = 'Horário a definir'
            if lista_pms:
                horario_turno = formatar_horario_turno(
                    getattr(lista_pms[0], 'hora_inicio_escala', None),
                    getattr(lista_pms[0], 'hora_fim_escala', None)
                )

            pms_bloco = []
            for pm in lista_pms:
                info_dia = calcular_escala_dia(pm, inicio)
                sigla = (info_dia.get('sigla') or '').strip().upper()
                descricao = info_dia.get('descricao', sigla)

                if sigla == 'TRABALHO':
                    status = 'SERVIÇO'
                elif sigla == 'FOLGA':
                    status = 'FOLGA'
                else:
                    status = descricao if descricao else sigla

                pm.status = status
                pm.dias = []

                for i in range(7):
                    data_dia = inicio + timedelta(days=i)
                    info = calcular_escala_dia(pm, data_dia)
                    sigla_dia = (info.get('sigla') or '').strip().upper()
                    descricao_dia = info.get('descricao', sigla_dia)

                    if sigla_dia == 'TRABALHO':
                        valor = obter_texto_exibicao('Trabalho', pm)
                        is_afast = False
                    elif sigla_dia == 'FOLGA':
                        valor = 'FOLGA'
                        is_afast = False
                    else:
                        valor = descricao_dia if descricao_dia else sigla_dia
                        is_afast = is_afastamento(sigla_dia)

                    pm.dias.append({
                        'valor': valor,
                        'is_afastamento': is_afast
                    })

                pms_bloco.append(pm)

            blocos.append({
                'nome_bloco': nome_bloco,
                'horario_turno': horario_turno,
                'pms': pms_bloco
            })

        return render_template(
            'visualizar_escala_semanal.html',
            cia=cia,
            data_inicio=inicio,
            data_fim=fim,
            tipo_filtro=tipo_filtro,
            tipo_txt=tipo_txt,
            numerador=numerador,
            periodo_txt=periodo_txt,
            dias_semana=dias_semana,
            blocos=blocos,
            versao_txt=versao_txt,
            print_mode=False
        )

    except Exception as e:
        flash(f'Erro ao carregar visualização semanal: {e}', 'danger')
        return redirect(url_for('index.index'))


@escala_bp.route('/visualizar/semanal/imprimir')
@login_required
def imprimir_escala():
    cia = request.args.get('cia', 'EM')
    tipo_filtro = request.args.get('tipo_filtro', 'TODOS')
    data_str = request.args.get('data_inicio')

    if data_str:
        inicio = datetime.strptime(data_str, '%Y-%m-%d').date()
    else:
        hoje = datetime.now().date()
        inicio = hoje - timedelta(days=hoje.weekday())

    return visualizar_escala_semanal(cia, inicio, tipo_filtro)


@escala_bp.route('/visualizar/diaria')
@login_required
def visualizar_escala_diaria():
    cia = request.args.get('cia', 'EM')
    data_str = request.args.get('data')

    if data_str:
        data = datetime.strptime(data_str, '%Y-%m-%d').date()
    else:
        data = datetime.now().date()

    pms = Policial.query.filter_by(ativo=True).all()
    if cia != 'TODAS':
        pms = [pm for pm in pms if pm.cia == cia]

    retorno = []
    for pm in pms:
        info = calcular_escala_dia(pm, data)
        retorno.append({
            'pm': pm,
            'info': info
        })

    return render_template(
        'visualizar_escala_diaria.html',
        cia=cia,
        data=data,
        resultados=retorno
    )


@escala_bp.route('/alterar', methods=['POST'])
@login_required
def alterar_escala():
    try:
        policial_id = request.form.get('policial_id', type=int)
        data_str = request.form.get('data')
        valor = (request.form.get('valor') or '').strip()
        motivo = (request.form.get('motivo') or '').strip()
        hora_inicio_ee = request.form.get('hora_inicio_ee') or None
        hora_fim_ee = request.form.get('hora_fim_ee') or None

        if not policial_id or not data_str:
            flash('Erro: dados incompletos.', 'danger')
            return redirect(url_for('index.index'))

        data = datetime.strptime(data_str, '%Y-%m-%d').date()

        if valor == 'LIMPAR':
            AlteracaoEscala.query.filter_by(policial_id=policial_id, data=data).delete()
        else:
            alt = AlteracaoEscala.query.filter_by(policial_id=policial_id, data=data).first()
            if not alt:
                alt = AlteracaoEscala(policial_id=policial_id, data=data)

            alt.valor = valor
            alt.motivo = motivo
            alt.hora_inicio_ee = hora_inicio_ee
            alt.hora_fim_ee = hora_fim_ee

            db.session.add(alt)

        db.session.add(LogAuditoria(
            usuario_id=getattr(current_user, 'id', None),
            usuario_nome=getattr(current_user, 'nome', ''),
            usuario_re=getattr(current_user, 're', ''),
            acao='ALTERAR_ESCALA',
            detalhes=f'PM {policial_id} em {data_str} => {valor}'
        ))

        db.session.commit()
        flash('✅ Escala atualizada com sucesso!', 'success')
        return redirect(url_for('escala.visualizar_escala_semanal', cia=request.form.get('cia', 'EM')))

    except Exception as e:
        db.session.rollback()
        flash(f'Erro: {e}', 'danger')
        return redirect(url_for('index.index'))