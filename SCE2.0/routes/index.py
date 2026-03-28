from flask import Blueprint, render_template, request
from flask_login import login_required, current_user
from sqlalchemy import or_

from models import Policial
from utils.textos import get_habilitacoes_texto, get_sat_texto
from utils.ordenacao import ordenar_por_antiguidade, ordenar_em_agrupado

index_bp = Blueprint('index', __name__)


@index_bp.route('/')
@login_required
def index():
    pms_raw = Policial.query.filter_by(ativo=True).all()

    nome_filtro = (request.args.get('nome') or '').strip()
    cia_filtro = (request.args.get('cia') or '').strip()
    posto_filtro = (request.args.get('posto') or '').strip()

    query = Policial.query.filter_by(ativo=True)

    if nome_filtro:
        query = query.filter(
            or_(
                Policial.nome.ilike(f'%{nome_filtro}%'),
                Policial.nome_guerra.ilike(f'%{nome_filtro}%'),
                Policial.re.ilike(f'%{nome_filtro}%')
            )
        )

    if cia_filtro and cia_filtro != 'TODAS':
        query = query.filter_by(cia=cia_filtro)

    if posto_filtro and posto_filtro != 'TODOS':
        query = query.filter_by(posto=posto_filtro)

    pms_raw = query.all()

    tem_em = any(pm.cia == 'EM' for pm in pms_raw)
    if tem_em:
        pms_em = [p for p in pms_raw if p.cia == 'EM']
        pms_outros = [p for p in pms_raw if p.cia != 'EM']
        pms_ordenados = ordenar_em_agrupado(pms_em) + ordenar_por_antiguidade(pms_outros)
    else:
        pms_ordenados = ordenar_por_antiguidade(pms_raw)

    e_admin = (getattr(current_user, 'perfil', 'usuario') == 'admin')

    return render_template(
        'index.html',
        pms=pms_ordenados,
        user=current_user,
        e_admin=e_admin,
        get_habilitacoes_texto=get_habilitacoes_texto,
        get_sat_texto=get_sat_texto
    )