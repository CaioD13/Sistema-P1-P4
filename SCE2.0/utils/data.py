from datetime import datetime


def formatar_data_militar(data):
    meses = {
        1: 'JAN', 2: 'FEV', 3: 'MAR', 4: 'ABR', 5: 'MAI', 6: 'JUN',
        7: 'JUL', 8: 'AGO', 9: 'SET', 10: 'OUT', 11: 'NOV', 12: 'DEZ'
    }
    ano_curto = data.year % 100
    return f"{data.day:02d}{meses[data.month]}{ano_curto}"


def get_semana_mes(data):
    semana = (data.day - 1) // 7 + 1
    return f"{data.month:02d}.{semana}"