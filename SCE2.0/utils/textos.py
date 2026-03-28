from models import Policial


def normalizar_texto(valor):
    return (valor or '').strip()


def obter_posto_pm(pm):
    return normalizar_texto(getattr(pm, 'posto', None))


def obter_secao_pm(pm):
    secao = normalizar_texto(getattr(pm, 'secao_em', None))
    return secao if secao else 'Outros / A Definir'


def obter_texto_exibicao(sigla, pm=None):
    if sigla == 'Trabalho' and pm:
        hora_inicio = getattr(pm, 'hora_inicio_escala', None)
        hora_fim = getattr(pm, 'hora_fim_escala', None)

        if hora_inicio and hora_fim:
            return f"{hora_inicio.strftime('%H')}h/{hora_fim.strftime('%H')}h"

    return sigla


def get_emojis(selections, emoji_dict):
    return ' '.join(emoji_dict.get(sel, '') for sel in selections if sel in emoji_dict).strip()


def deserialize_selections(serialized):
    return serialized.split(',') if serialized else ''


def get_habilitacoes_texto(pm):
    HABILITACOES_EMOJIS = {
        'Taser / AIN': 'AIN-',
        'Fuzil 556': 'F556-',
        'Fuzil 762': 'F762-',
        'Escopeta Benelli cal 12': 'EB12-',
        'Calibre 12': 'E12'
    }
    return get_emojis(deserialize_selections(getattr(pm, 'habilitacoes', None)), HABILITACOES_EMOJIS)


def get_sat_texto(pm):
    SAT_EMOJIS = {
        'A (Motos)': 'A',
        'B (Carros)': 'B',
        'C (Carga acima de 3.500kg)': 'C',
        'D (Transporte de passageiros)': 'D',
        'E (Veículo Combinado)': 'E'
    }
    return get_emojis(deserialize_selections(getattr(pm, 'sat', None)), SAT_EMOJIS)