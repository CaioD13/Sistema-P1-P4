def is_afastamento(sigla):
    afastamentos = {
        'JM', 'M1', 'M2', 'FM', 'DS', 'DR', 'FR', 'PF', 'F*', 'LSV', 'AD',
        'LTS', 'CV', 'FN', 'FS', 'LG', 'LA', 'PT', 'NP', 'LP', 'EAP', 'LT',
        'FOLGA', 'F', 'A', 'SS', 'SR', 'EE'
    }
    return (sigla or '').strip().upper() in afastamentos


def traduzir_para_frequencia(sigla, duracao_horas_escala=None):
    if not sigla:
        return {"codigo": "?", "conta_dias": False, "diarias": None, "motivo": "Sigla ausente"}

    sigla = str(sigla).strip().upper()

    afastamentos = {
        'A', 'FOLGA', 'F', 'F*', 'JM', 'FM', 'FN', 'F.N.', 'FÉRIAS', 'FERIAS',
        'FS', 'DS', 'DR', 'FR', 'PF', 'CV', 'LTS', 'LT', 'LG', 'PT', 'LA',
        'NP', 'LICENÇA', 'AFASTADO', 'DISPENSA', 'LICENCA', 'LSV', 'AD',
        'L.P.', 'LP', 'LICENÇA PRÊMIO', 'LICENCA PREMIO'
    }

    if sigla in afastamentos:
        return {"codigo": "A", "conta_dias": False, "diarias": 0, "motivo": "Afastamento/Folga"}

    if sigla in {"SR", "PPJMD"}:
        return {"codigo": "3", "conta_dias": True, "diarias": 3, "motivo": "Sigla especial 3 diárias"}

    if sigla in {"SS", "EE"}:
        return {"codigo": "?", "conta_dias": False, "diarias": None, "motivo": "Sigla especial sem regra"}

    try:
        duracao = float(duracao_horas_escala) if duracao_horas_escala is not None else None
    except (TypeError, ValueError):
        duracao = None

    if duracao is None:
        return {"codigo": "?", "conta_dias": False, "diarias": None, "motivo": "Duração ausente ou inválida"}

    if duracao < 8:
        return {"codigo": "0", "conta_dias": True, "diarias": 0, "motivo": "Turno inferior a 8h"}
    elif duracao < 12:
        return {"codigo": "1", "conta_dias": True, "diarias": 1, "motivo": "Turno entre 8h e 12h"}
    elif duracao < 18:
        return {"codigo": "2", "conta_dias": True, "diarias": 2, "motivo": "Turno entre 12h e 18h"}
    else:
        return {"codigo": "3", "conta_dias": True, "diarias": 3, "motivo": "Turno igual ou superior a 18h"}