from datetime import datetime, date, timedelta

from models import Policial, AlteracaoEscala, EscalaVersao
from utils.textos import normalizar_texto, obter_posto_pm, obter_secao_pm, obter_texto_exibicao
from utils.ordenacao import (
    ordenar_por_patente,
    ordenar_em_agrupado,
    ordenar_cia_agrupado,
    ORDEM_SECOES
)
from utils.validacoes import is_afastamento


def calcular_duracao_hhmm(hora_inicio, hora_fim):
    if not hora_inicio or not hora_fim:
        return None

    inicio = datetime.strptime(hora_inicio, '%H:%M')
    fim = datetime.strptime(hora_fim, '%H:%M')

    if fim <= inicio:
        fim += timedelta(days=1)

    return (fim - inicio).total_seconds() / 3600


def formatar_horario_turno(hora_inicio_escala, hora_fim_escala):
    def _fmt(valor):
        if not valor:
            return ''
        if isinstance(valor, str):
            v = valor.strip()
            if len(v) >= 5 and v[2] == ':':
                return v[:2] + 'h'
            if len(v) >= 2 and v[:2].isdigit():
                return v[:2] + 'h'
            return v
        try:
            return valor.strftime('%Hh')
        except Exception:
            return str(valor)

    ini = _fmt(hora_inicio_escala)
    fim = _fmt(hora_fim_escala)

    if ini and fim:
        return f'{ini}/{fim}'

    return ini or fim or ''


def calcular_escala_dia(pm, data, registros_semana=None):
    if isinstance(data, datetime):
        data = data.date()

    alt = AlteracaoEscala.query.filter_by(policial_id=pm.id, data=data).first()
    if alt and alt.valor and alt.valor.strip().upper() not in ['LIMPAR']:
        valor_manual = alt.valor.strip().upper()

        afastamentos = ['JM', 'M1', 'M2', 'FM', 'DS', 'DR', 'FR', 'PF', 'F*', 'LSV', 'AD', 'LTS', 'CV', 'FN', 'FS', 'LG', 'LA', 'PT', 'NP', 'LP', 'EAP', 'LT', 'Férias']

        if valor_manual in afastamentos:
            return {
                "sigla": valor_manual,
                "classe": "bg-warning text-dark",
                "descricao": valor_manual,
                "trabalha": False,
                "ciclo_dia": None,
                "regime": getattr(pm, "regime_escala", ""),
                "duracao_horas_escala": 0.0,
            }

        elif valor_manual == 'EE':
            hora_inicio = alt.hora_inicio_ee or ''
            hora_fim = alt.hora_fim_ee or ''
            duracao = calcular_duracao_hhmm(hora_inicio, hora_fim) if hora_inicio and hora_fim else 0.0
            descricao = f"EE {hora_inicio}–{hora_fim}" if hora_inicio and hora_fim else 'EE'
            return {
                "sigla": "EE",
                "classe": "bg-warning text-dark",
                "descricao": descricao,
                "trabalha": True,
                "ciclo_dia": None,
                "regime": getattr(pm, "regime_escala", ""),
                "duracao_horas_escala": duracao,
            }

        elif valor_manual == 'DES':
            info_base = _calcular_escala_base(pm, data)
            return {
                "sigla": info_base["sigla"],
                "classe": f"{info_base['classe']} des-marcado",
                "descricao": f"{info_base['descricao']} DES",
                "trabalha": info_base["trabalha"],
                "ciclo_dia": info_base["ciclo_dia"],
                "regime": getattr(pm, "regime_escala", ""),
                "duracao_horas_escala": info_base["duracao_horas_escala"],
            }
        else:
            return {
                "sigla": valor_manual,
                "classe": "bg-warning text-dark",
                "descricao": f"Alteração manual: {valor_manual}",
                "trabalha": False,
                "ciclo_dia": None,
                "regime": getattr(pm, "regime_escala", ""),
                "duracao_horas_escala": 0.0,
            }

    return _calcular_escala_base(pm, data)


def _calcular_escala_base(pm, data):
    data_base = getattr(pm, "data_base_escala", None)
    if isinstance(data_base, datetime):
        data_base = data_base.date()

    regime_raw = getattr(pm, "regime_escala", "") or ""
    regime = regime_raw.strip().lower()

    offset = None
    if isinstance(data_base, date) and isinstance(data, date):
        offset = (data - data_base).days

    duracao_horas_escala = getattr(pm, "duracao_horas_escala", None)
    if duracao_horas_escala is not None:
        try:
            duracao_horas_escala = float(duracao_horas_escala)
        except (TypeError, ValueError):
            duracao_horas_escala = None

    default = {
        "sigla": "Folga",
        "classe": "bg-light text-muted",
        "descricao": "Folga",
        "trabalha": False,
        "ciclo_dia": None,
        "regime": regime_raw,
        "duracao_horas_escala": duracao_horas_escala,
    }

    if not regime:
        return default

    if regime == "administrativo":
        if data.weekday() >= 5:
            return {
                "sigla": "Folga",
                "classe": "bg-light text-muted",
                "descricao": "Folga",
                "trabalha": False,
                "ciclo_dia": data.weekday(),
                "regime": regime_raw,
                "duracao_horas_escala": duracao_horas_escala,
            }

        return {
            "sigla": "Trabalho",
            "classe": "bg-success text-white",
            "descricao": "Trabalho",
            "trabalha": True,
            "ciclo_dia": data.weekday(),
            "regime": regime_raw,
            "duracao_horas_escala": duracao_horas_escala,
        }

    if offset is None:
        return default

    if regime == "12x36":
        ciclo_dia = offset % 2
        trabalha = (ciclo_dia == 0)
    elif regime == "12x24_12x48":
        ciclo_dia = offset % 4
        trabalha = ciclo_dia in (0, 1)
    elif regime == "12x60":
        ciclo_dia = offset % 6
        trabalha = (ciclo_dia == 0)
    elif regime == "12x36_12x60":
        ciclo_dia = offset % 6
        trabalha = ciclo_dia in (0, 3)
    elif regime == "24x48":
        ciclo_dia = offset % 3
        trabalha = (ciclo_dia == 0)
    else:
        return default

    if trabalha:
        return {
            "sigla": "Trabalho",
            "classe": "bg-success text-white",
            "descricao": "Trabalho",
            "trabalha": True,
            "ciclo_dia": ciclo_dia,
            "regime": regime_raw,
            "duracao_horas_escala": duracao_horas_escala,
        }

    return {
        "sigla": "Folga",
        "classe": "bg-light text-muted",
        "descricao": "Folga",
        "trabalha": False,
        "ciclo_dia": ciclo_dia,
        "regime": regime_raw,
        "duracao_horas_escala": duracao_horas_escala,
    }


def filtrar_por_tipo_hierarquico(pms, tipo_filtro):
    oficiais = {
        'CORONEL PM',
        'TENENTE CORONEL PM',
        'MAJOR PM',
        'CAPITÃO PM',
        '1º TENENTE PM',
        '2º TENENTE PM',
        'ASPIRANTE A OFICIAL'
    }

    pracas = {
        'SUBTENENTE PM',
        '1º SARGENTO PM',
        '2º SARGENTO PM',
        '3º SARGENTO PM',
        'CABO PM',
        'SOLDADO PM'
    }

    def posto_normalizado(pm):
        posto = (
            getattr(pm, 'posto_grad', None)
            or getattr(pm, 'posto', None)
            or ''
        )
        return normalizar_texto(posto).upper()

    def secao_normalizada(pm):
        return normalizar_texto(getattr(pm, 'secao_em', None)).upper()

    if tipo_filtro == 'OFICIAIS':
        return [pm for pm in pms if posto_normalizado(pm) in oficiais]

    if tipo_filtro == 'PRACAS':
        return [pm for pm in pms if posto_normalizado(pm) in pracas]

    if tipo_filtro == 'CFP':
        return [pm for pm in pms if secao_normalizada(pm) == 'CFP']

    return pms


def cfp_turno(pm):
    hora_ini = getattr(pm, 'hora_inicio_escala', None)

    if not hora_ini:
        return 'CFP Diurno'

    try:
        if isinstance(hora_ini, datetime):
            hh = hora_ini.hour
            mm = hora_ini.minute
        else:
            texto = str(hora_ini).strip()
            if not texto:
                return 'CFP Diurno'
            partes = texto.split(':')
            hh = int(partes[0])
            mm = int(partes[1]) if len(partes) > 1 else 0

        if hh > 16 or (hh == 16 and mm > 59):
            return 'CFP Noturno'
        return 'CFP Diurno'
    except Exception:
        return 'CFP Diurno'


def agrupar_por_secao(pms, tipo_filtro='TODOS'):
    if tipo_filtro == 'CFP':
        mapa_cfp = {'CFP Diurno': [], 'CFP Noturno': []}

        for pm in pms:
            if normalizar_texto(getattr(pm, 'secao_em', None)).upper() != 'CFP':
                continue
            bloco = cfp_turno(pm)
            mapa_cfp[bloco].append(pm)

        mapa_cfp['CFP Diurno'] = ordenar_por_patente(mapa_cfp['CFP Diurno'])
        mapa_cfp['CFP Noturno'] = ordenar_por_patente(mapa_cfp['CFP Noturno'])
        return {k: v for k, v in mapa_cfp.items() if v}

    mapa = {}
    for pm in pms:
        secao = obter_secao_pm(pm)
        mapa.setdefault(secao, []).append(pm)

    for secao in mapa:
        mapa[secao] = ordenar_por_patente(mapa[secao])

    if tipo_filtro == 'OFICIAIS':
        ordem_preferencial = ['Comando', 'SubComando', 'CoordOP']
        mapa_ordenado = {}

        for secao in ordem_preferencial:
            if secao in mapa and mapa[secao]:
                mapa_ordenado[secao] = mapa[secao]

        for secao, lista in mapa.items():
            if secao not in mapa_ordenado and lista:
                mapa_ordenado[secao] = lista

        return mapa_ordenado

    if tipo_filtro == 'PRACAS':
        return {k: v for k, v in mapa.items() if v}

    mapa_ordenado = {}
    for secao in ORDEM_SECOES:
        if secao in mapa and mapa[secao]:
            mapa_ordenado[secao] = mapa[secao]

    for secao, lista in mapa.items():
        if secao not in mapa_ordenado and lista:
            mapa_ordenado[secao] = lista

    return mapa_ordenado


def ordem_blocos_para_tipo(tipo_filtro='TODOS'):
    if tipo_filtro == 'CFP':
        return ['CFP Diurno', 'CFP Noturno']

    if tipo_filtro == 'OFICIAIS':
        ordem = ['Comando', 'SubComando', 'CoordOP']
        for secao in ORDEM_SECOES:
            if secao not in ordem:
                ordem.append(secao)
        return ordem

    return ORDEM_SECOES


def get_proxima_versao(cia, data_inicio, tipo_filtro):
    registro = EscalaVersao.query.filter_by(
        cia=cia,
        data_inicio=data_inicio,
        tipo_filtro=tipo_filtro
    ).first()

    if registro:
        registro.versao_atual += 1
        versao = registro.versao_atual
    else:
        registro = EscalaVersao(cia=cia, data_inicio=data_inicio, tipo_filtro=tipo_filtro, versao_atual=1)
        from extensions import db
        db.session.add(registro)
        versao = 1

    from extensions import db
    db.session.commit()
    return f"V{versao}.0"