from utils.textos import obter_posto_pm, obter_secao_pm, normalizar_texto

HIERARQUIA = {
    'Coronel': 1, 'Coronel PM': 1,
    'Ten Coronel': 2, 'Tenente Coronel PM': 2,
    'Major': 3, 'Major PM': 3,
    'Capitão': 4, 'Capitão PM': 4,
    '1º Tenente': 5, '1º Tenente PM': 5,
    '2º Tenente': 6, '2º Tenente PM': 6,
    'Asp Oficial': 7, 'Aspirante a Oficial': 7,
    'Subtenente': 8, 'Subtenente PM': 8,
    '1º Sgt': 9, '1º Sargento PM': 9,
    '2º Sgt': 10, '2º Sargento PM': 10,
    '3º Sgt': 11, '3º Sargento PM': 11,
    'Cabo': 12, 'Cabo PM': 12,
    'Soldado': 13, 'Soldado PM': 13,
    'Sd 2ª Cl PM': 14
}

ORDEM_SECOES = [
    "Comando",
    "Subcomando",
    "Coord Op",
    "P1",
    "P2",
    "P3",
    "P4",
    "P5",
    "UGE",
    "UIS",
    "DivOP",
    "Ordenança",
    "GT",
    "Sub Frota",
    "DivADM",
    "SPJMD",
    "Secretaria",
    "Motomec"
]

ORDEM_SECOES_EM = {
    'Comando': 1, 'Subcomando': 2, 'Coord Op': 3,
    'P1': 4, 'P2': 5, 'P3': 6, 'P4': 7, 'P5': 8,
    'UGE': 9, 'UIS': 10, 'DivOP': 11, 'Ordenança': 12, 'GT': 13,
    'Sub Frota': 14, 'DivADM': 15, 'SPJMD': 16, 'Secretaria': 17,
    'Motomec': 18, 'CFP': 19
}


def ordenar_por_patente(pm_list):
    def chave(pm):
        posto = obter_posto_pm(pm)
        nome = normalizar_texto(getattr(pm, 'nome_guerra', None)) or normalizar_texto(getattr(pm, 'nome', None))
        return (
            HIERARQUIA.get(posto, 999),
            posto,
            nome
        )
    return sorted(pm_list, key=chave)


def ordenar_por_antiguidade(lista_pms):
    return sorted(lista_pms, key=lambda pm: (HIERARQUIA.get(pm.posto, 99), pm.re))


def ordenar_em_agrupado(lista_pms):
    return sorted(lista_pms, key=lambda pm: (
        ORDEM_SECOES_EM.get(pm.secao_em, 99),
        HIERARQUIA.get(pm.posto, 99),
        pm.re
    ))


def ordenar_cia_agrupado(lista_pms):
    return sorted(lista_pms, key=lambda pm: (
        getattr(pm, 'equipe', 'Z') or 'Z',
        HIERARQUIA.get(pm.posto, 99),
        pm.re
    ))