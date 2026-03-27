import os
import calendar
import xlsxwriter
from datetime import datetime, timedelta, date, time
from io import BytesIO
from flask import Flask, render_template, request, redirect, url_for, flash, send_file
from werkzeug.security import generate_password_hash, check_password_hash
from flask_login import LoginManager, login_user, login_required, logout_user, current_user, UserMixin
from sqlalchemy import or_, desc
from flask_sqlalchemy import SQLAlchemy
from sqlalchemy.orm import DeclarativeBase

INFO_SECOES = {
    "Comando": "Horário a verificar",
    "Subcomando": "Horário a verificar",
    "Coord Op": "Horário a verificar",
    "P1": "Horário a verificar",
    "P2": "Horário a verificar",
    "P3": "Horário a verificar",
    "P4": "Horário a verificar",
    "P5": "Horário a verificar",
    "UGE": "Horário a verificar",
    "UIS": "Horário a verificar",
    "DivOP": "Horário a verificar",
    "Ordenança": "Horário a verificar",
    "GT": "Horário a verificar",
    "Sub Frota": "Horário a verificar",
    "DivADM": "Horário a verificar",
    "SPJMD": "Horário a verificar",
    "Secretaria": "Horário a verificar",
    "Motomec": "Horário a verificar",
    "CFP Diurno": "05h00 - 16h59 / instrução diurna",
    "CFP Noturno": "17h00 - 05h00 / instrução noturna",
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

SIGLAS_AFASTAMENTOS = {
    'JM', 'M1', 'M2', 'FM', 'DS', 'DR', 'FR', 'PF', 'F*', 'LSV', 'AD',
    'LTS', 'CV', 'FN', 'FS', 'LG', 'LA', 'PT', 'NP', 'LP', 'EAP', 'LT'
}

ORDEM_SECOES_EM = {
    'Comando': 1, 'Subcomando': 2, 'Coord Op': 3,
    'P1': 4, 'P2': 5, 'P3': 6, 'P4': 7, 'P5': 8,
    'UGE': 9, 'UIS': 10, 'DivOP': 11, 'Ordenança': 12, 'GT': 13,
    'Sub Frota': 14, 'DivADM': 15, 'SPJMD': 16, 'Secretaria': 17,
    'Motomec': 18,  'CFP': 19
}

INFO_BLOCOS = {
    "Administração": "08h00 - 18h00 / jornada administrativa",
    "Manutenção/Motomec": "07h00 - 19h00 / apoio operacional",
    "Guarda do Quartel Diurno": "06h00 - 18h00 / período diurno",
    "Guarda do Quartel Noturno": "18h00 - 06h00 / período noturno",
    "CFP Diurno": "05h00 - 17h00 / instrução diurna",
    "CFP Noturno": "17h00 - 05h00 / instrução noturna",
    "Outros / A Definir": "Horário a verificar"
}

ORDEM_BLOCOS = [
    "Administração",
    "Manutenção/Motomec",
    "Guarda do Quartel Diurno",
    "Guarda do Quartel Noturno",
    "CFP Diurno",
    "CFP Noturno",
    "Outros / A Definir"
]

STATUS_INFO = {
    # --- NOVO MODELO POR HORÁRIO ---
    'Trabalho': {
        'classe': 'bg-success text-white',
        'descricao': None  # Será substituído por horário
    },
    
    'F': {
        'classe': 'bg-light text-muted border-light',
        'descricao': 'Folga'
    },
    'A': {
        'classe': 'bg-light text-muted border-light',
        'descricao': 'Afastamento'
    },

    # --- CÓDIGOS ESPECIAIS QUE CONTINUAM VÁLIDOS ---
    'FN': {
        'classe': 'bg-warning text-dark',
        'descricao': 'Férias Normais'
    },
    'FS': {
        'classe': 'bg-warning text-dark',
        'descricao': 'Férias Sobrestadas'
    },
    'LTS': {
        'classe': 'bg-warning text-dark',
        'descricao': 'Licença Saúde'
    },
    'CV': {
        'classe': 'bg-warning text-dark',
        'descricao': 'Convalescença'
    },
    'LG': {
        'classe': 'bg-warning text-dark',
        'descricao': 'Licença Gestante'
    },
    'PT': {
        'classe': 'bg-warning text-dark',
        'descricao': 'Licença Paternidade'
    },
    'LA': {
        'classe': 'bg-warning text-dark',
        'descricao': 'Licença Adoção'
    },
    'NP': {
        'classe': 'bg-warning text-dark',
        'descricao': 'Núpcias'
    },
    'LT': {
        'classe': 'bg-warning text-dark',
        'descricao': 'Luto'
    },
    'LP': {
        'classe': 'bg-info text-white',
        'descricao': 'Licença Prêmio'
    },
    'LSV': {
        'classe': 'bg-secondary text-white',
        'descricao': 'Licença sem Vencimentos'
    },
    'AD': {
        'classe': 'bg-primary text-white',
        'descricao': 'Adido a outra OPM'
    },
    'SS': {
        'classe': 'bg-info text-dark',
        'descricao': 'Superior Sobreaviso'
    },
    'SR': {
        'classe': 'bg-info text-dark',
        'descricao': 'Supervisor Regional'
    },
    'EE': {
        'classe': 'bg-info text-dark',
        'descricao': 'Escala Extra'
    },
    'EAP': {
        'classe': 'bg-warning text-dark',
        'descricao': 'Curso Presencial'
    },
}

# ==============================================================================
# CONFIGURAÇÕES INICIAIS
# ==============================================================================

basedir = os.path.abspath(os.path.dirname(__file__))
db_path = os.path.join(basedir, 'frequencia.db')
print("📌 CURRENT WORKING DIR:", os.getcwd())
print("📌 FILE LOCATION:", __file__)
print("📌 BASEDIR:", os.path.abspath(os.path.dirname(__file__)))
print("📌 BANCO CARREGADO:", db_path)

app = Flask(__name__)
app.config['SQLALCHEMY_DATABASE_URI'] = f'sqlite:///{db_path}'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
app.config['SECRET_KEY'] = 'chave_super_secreta_37bpmm_2026_final'
app.config['SESSION_PERMANENT'] = False
app.config['PERMANENT_SESSION_LIFETIME'] = timedelta(minutes=30)

db = SQLAlchemy(app)
login_manager = LoginManager()
login_manager.init_app(app)
login_manager.login_view = 'login'

@login_manager.user_loader
def load_user(user_id):
    return db.session.get(Usuario, int(user_id))

# ==============================================================================
# MODELOS DE BANCO DE DADOS
# ==============================================================================

class HistoricoEscala(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    policial_id = db.Column(db.Integer, db.ForeignKey('policial.id'), nullable=False)
    mes = db.Column(db.Integer, nullable=False)
    ano = db.Column(db.Integer, nullable=False)
    tipo_escala_antigo = db.Column(db.String(20), nullable=False)

class Usuario(UserMixin, db.Model):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(20), unique=True, nullable=False)
    re = db.Column(db.String(20), unique=True, nullable=False)
    nome = db.Column(db.String(100), nullable=False)
    password = db.Column(db.String(100), nullable=False)
    perfil = db.Column(db.String(20), default='comum') # 'admin', 'operador', 'comum'

class Policial(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    posto = db.Column(db.String(20))
    re = db.Column(db.String(20), unique=True)
    nome = db.Column(db.String(100))
    nome_guerra = db.Column(db.String(50))
    cia = db.Column(db.String(20))
    secao_em = db.Column(db.String(50))
    tipo_escala = db.Column(db.String(20))
    data_base_escala = db.Column(db.Date, nullable=True)
    funcao = db.Column(db.String(100))
    equipe = db.Column(db.String(50))
    viatura = db.Column(db.String(50))
    ativo = db.Column(db.Boolean, default=True)
    hora_inicio_escala = db.Column(db.Time, nullable=True)
    hora_fim_escala = db.Column(db.Time, nullable=True)
    duracao_horas_escala = db.Column(db.Float, nullable=True)
    regime_escala = db.Column(db.String(50), nullable=True)
    habilitacoes = db.Column(db.String(500), nullable=True)
    sat = db.Column(db.String(500), nullable=True)

    # Campos de Saúde
    status_saude = db.Column(db.String(50), default='Apto Sem Restrições')
    restricao_tipo = db.Column(db.String(100))
    data_fim_restricao = db.Column(db.Date)

class AlteracaoEscala(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    policial_id = db.Column(db.Integer, db.ForeignKey('policial.id'))
    data = db.Column(db.Date)
    valor = db.Column(db.String(10))
    motivo = db.Column(db.String(200))
    hora_inicio_ee = db.Column(db.String(5))
    hora_fim_ee = db.Column(db.String(5)) 

class LogAuditoria(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    data_hora = db.Column(db.DateTime, default=datetime.now)
    usuario_id = db.Column(db.Integer)
    usuario_nome = db.Column(db.String(100))
    usuario_re = db.Column(db.String(20))
    acao = db.Column(db.String(50))
    detalhes = db.Column(db.String(500))

class HistoricoMovimentacao(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    policial_id = db.Column(db.Integer, db.ForeignKey('policial.id'))
    data_mudanca = db.Column(db.DateTime, default=datetime.now)
    tipo_mudanca = db.Column(db.String(50))
    anterior = db.Column(db.String(100))
    novo = db.Column(db.String(100))
    usuario_responsavel = db.Column(db.String(100))

class EscalaVersao(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    cia = db.Column(db.String(50))
    data_inicio = db.Column(db.Date)
    tipo_filtro = db.Column(db.String(20))
    versao_atual = db.Column(db.Integer, default=1)

class ValidacaoMensal(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    policial_id = db.Column(db.Integer, db.ForeignKey('policial.id'))
    mes = db.Column(db.Integer)
    ano = db.Column(db.Integer)
    admin_id = db.Column(db.Integer)
    admin_info = db.Column(db.String(200)) # Vai guardar: "1º Ten PM Silva (RE 123456)"
    data_validacao = db.Column(db.DateTime, default=datetime.now)


# ==============================================================================
# FUNÇÕES AUXILIARES E LÓGICA DE NEGÓCIO
# ==============================================================================

def normalizar_texto(valor):
    return (valor or '').strip()

def obter_posto_pm(pm):
    return normalizar_texto(getattr(pm, 'posto', None))

def obter_secao_pm(pm):
    secao = normalizar_texto(getattr(pm, 'secao_em', None))
    return secao if secao else 'Outros / A Definir'

def ordenar_por_patente(pm_list):
    def chave(pm):
        posto = obter_posto_pm(pm)
        return (
            HIERARQUIA.get(posto, 999),
            posto,
            normalizar_texto(getattr(pm, 'nome_guerra', None)) or normalizar_texto(getattr(pm, 'nome', None))
        )
    return sorted(pm_list, key=chave)

def cfp_turno(pm):
    """
    horário início após às 16h59 = CFP Noturno
    """
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
            if obter_secao_pm(pm) != 'CFP':
                continue
            bloco = cfp_turno(pm)
            mapa_cfp[bloco].append(pm)

        mapa_cfp['CFP Diurno'] = ordenar_por_patente(mapa_cfp['CFP Diurno'])
        mapa_cfp['CFP Noturno'] = ordenar_por_patente(mapa_cfp['CFP Noturno'])
        return mapa_cfp

    mapa = {secao: [] for secao in ORDEM_SECOES}

    for pm in pms:
        secao = obter_secao_pm(pm)
        if secao in mapa:
            mapa[secao].append(pm)

    for secao in mapa:
        mapa[secao] = ordenar_por_patente(mapa[secao])

    return {k: v for k, v in mapa.items() if v}

def ordem_blocos_para_tipo(tipo_filtro='TODOS'):
    if tipo_filtro == 'CFP':
        return ['CFP Diurno', 'CFP Noturno']
    return ORDEM_SECOES

def is_afastamento(sigla):
    return (sigla or '').strip().upper() in SIGLAS_AFASTAMENTOS

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

def filtrar_por_tipo_hierarquico(pms, tipo_filtro):
    oficiais = [
        'Coronel', 'Coronel PM', 'Ten Coronel', 'Tenente Coronel PM',
        'Major', 'Major PM', 'Capitão', 'Capitão PM',
        '1º Tenente', '1º Tenente PM', '2º Tenente', '2º Tenente PM',
        'Asp Oficial', 'Aspirante a Oficial'
    ]
    if tipo_filtro == 'OFICIAIS':
        return [pm for pm in pms if pm.posto in oficiais]
    elif tipo_filtro == 'PRACAS':
        return [pm for pm in pms if pm.posto not in oficiais]
    return pms

def registrar_log(acao, detalhes):
    if current_user.is_authenticated:
        try:
            novo = LogAuditoria(
                usuario_id=current_user.id, usuario_nome=current_user.nome,
                usuario_re=getattr(current_user, 're', 'N/A'), acao=acao, detalhes=detalhes
            )
            db.session.add(novo)
            db.session.commit()
        except: pass

def registrar_movimentacao(pm_id, tipo, anterior, novo):
    if anterior != novo:
        try:
            hist = HistoricoMovimentacao(
                policial_id=pm_id,
                tipo_mudanca=tipo,
                anterior=str(anterior or 'N/A'),
                novo=str(novo or 'N/A'),
                usuario_responsavel=current_user.nome
            )
            db.session.add(hist)
        except: pass

def e_feriado(data_obj):
    if isinstance(data_obj, datetime): data = data_obj.date()
    else: data = data_obj
    feriados = [(1, 1), (21, 4), (1, 5), (9, 7), (12, 10), (2, 11), (15, 11), (25, 12)]
    return (data.day, data.month) in feriados

def _normalizar_data_base(data_base):
    if isinstance(data_base, date):
        return data_base
    return None


def _calcular_offset_dias(data_base, data_alvo):
    if data_base is None or data_alvo is None:
        return None
    return (data_alvo - data_base).days


def _inicio_da_semana(data_alvo):
    return data_alvo - timedelta(days=data_alvo.weekday())


def _obter_sigla_administrativo(pm, data):
    """
    Regime Administrativo:
    - Segunda a sexta: trabalha
    - Sábado e domingo: folga

    As siglas JM, M1 e M2 devem ser controladas fora desta função,
    em uma estrutura específica de marcação diária/semana.
    """
    # 5 = sábado, 6 = domingo
    if data.weekday() >= 5:
        return {
            "sigla": "Folga",
            "classe": "bg-light text-muted",
            "descricao": "Folga",
            "trabalha": False,
            "ciclo_dia": data.weekday(),
            "regime": "Administrativo",
        }

    return {
        "sigla": "Trabalho",
        "classe": "bg-success text-white",
        "descricao": "Trabalho",
        "trabalha": True,
        "ciclo_dia": data.weekday(),
        "regime": "Administrativo",
    }


def validar_meios_expedientes_semanais(id_pm, inicio, fim, valor):
    """
    Valida meios expedientes conforme regras:
    - Permitido: 2 'M1', ou 2 'M2', ou 1 'M1' + 1 'M2' (máximo 2 meios na semana)
    - Proibido: 'M1' ou 'M2' + 'JM' (JM não pode coexistir com M1/M2)
    - JM: requer ≥ 4 dias trabalhados na semana (excluindo JM e período atual)
    - M1/M2: requer ≥ 3 dias trabalhados na semana (excluindo M1/M2 e período atual)
    
    Args:
        id_pm: ID do policial
        inicio: Data inicial do período
        fim: Data final do período
        valor: Valor a ser marcado ('JM', 'M1' ou 'M2')
    
    Returns:
        tuple: (bool sucesso, str mensagem)
    """
    if valor not in {'JM', 'M1', 'M2'}:
        return True, ""  # Não é meio expediente, valida automaticamente

    pm = db.session.get(Policial, id_pm)
    if not pm:
        return False, "Policial não encontrado."

    # Calcular semana da data inicial (segunda a domingo)
    inicio_semana = inicio - timedelta(days=inicio.weekday())  # Segunda-feira
    fim_semana = inicio_semana + timedelta(days=6)  # Domingo
    
    print(f"DEBUG: Semana calculada: {inicio_semana} a {fim_semana}")  # Debug - comente após testar
    
    # Buscar alterações existentes na semana (todas as de JM/M1/M2)
    alteracoes_semana = AlteracaoEscala.query.filter(
        AlteracaoEscala.policial_id == id_pm,
        AlteracaoEscala.data >= inicio_semana,
        AlteracaoEscala.data <= fim_semana,
        AlteracaoEscala.valor.in_(['JM', 'M1', 'M2'])
    ).all()
    
    print(f"DEBUG: Alterações JM/M1/M2 encontradas: {len(alteracoes_semana)}")  # Debug
    
    # Contar meios expedientes existentes (excluindo o período atual)
    siglas_existentes = []
    for alt in alteracoes_semana:
        if alt.data < inicio or alt.data > fim:  # Fora do período atual
            siglas_existentes.append(alt.valor)
    
    # Adicionar os que estamos salvando (mesmo valor para todo o período)
    dias_periodo = (fim - inicio).days + 1
    siglas_periodo = [valor] * dias_periodo
    siglas_totais = siglas_existentes + siglas_periodo
    
    print(f"DEBUG: Siglas totais: {siglas_totais}")  # Debug
    
    # Contar ocorrências
    count_m1 = siglas_totais.count("M1")
    count_m2 = siglas_totais.count("M2")
    count_jm = siglas_totais.count("JM")
    total_meios = count_m1 + count_m2 + count_jm

    # Regra 1: Máximo 2 meios expedientes na semana
    if total_meios > 2:
        return False, "Não é permitido mais de 2 meios expedientes na mesma semana."

    # Regra 2: JM não pode coexistir com M1/M2
    if count_jm > 0 and (count_m1 > 0 or count_m2 > 0):
        return False, "JM não pode ser combinado com M1 ou M2 na mesma semana."

    # Regra 3: JM requer ≥ 4 dias trabalhados na semana
    if count_jm > 0:
        dias_trabalhados = contar_dias_trabalhados_semana(id_pm, inicio_semana, fim_semana, pm, excluir_periodo=inicio, excluir_fim=fim)
        print(f"DEBUG: Dias trabalhados para JM: {dias_trabalhados}")  # Debug
        if dias_trabalhados < 4:
            return False, "Para ter direito a JM, o policial deve ter trabalhado ao menos 4 dias na semana."

    # Regra 4: M1/M2 requer ≥ 3 dias trabalhados na semana
    if count_m1 > 0 or count_m2 > 0:
        dias_trabalhados = contar_dias_trabalhados_semana(id_pm, inicio_semana, fim_semana, pm, excluir_periodo=inicio, excluir_fim=fim)
        print(f"DEBUG: Dias trabalhados para M1/M2: {dias_trabalhados}")  # Debug
        if dias_trabalhados < 3:
            return False, "Para ter direito a M1 ou M2, o policial deve ter trabalhado ao menos 3 dias na semana."

    # Regra 5: Combinações permitidas para M1/M2
    if count_m1 > 2 or count_m2 > 2:
        return False, "M1 ou M2 não podem se repetir mais de 2 vezes na semana."

    return True, ""


def contar_dias_trabalhados_semana(id_pm, inicio_semana, fim_semana, pm, excluir_periodo=None, excluir_fim=None):
    """
    Conta dias trabalhados normais na semana, incluindo dias sem alteração.
    
    Args:
        id_pm: ID do policial (para query)
        inicio_semana: Início da semana
        fim_semana: Fim da semana
        pm: Objeto Policial (para calcular_escala_dia)
        excluir_periodo: Data inicial a excluir (opcional)
        excluir_fim: Data final a excluir (opcional)
    
    Returns:
        int: Número de dias trabalhados
    """
    dias_trabalhados = 0
    current_date = inicio_semana
    
    while current_date <= fim_semana:
        # Verificar se é o período a excluir
        if excluir_periodo and excluir_fim:
            if current_date < excluir_periodo or current_date > excluir_fim:
                # Buscar alteração para o dia
                alt = AlteracaoEscala.query.filter_by(policial_id=id_pm, data=current_date).first()
                sigla = alt.valor if alt and alt.valor not in ['JM', 'M1', 'M2', 'LIMPAR'] else None
                
                # Se não há alteração válida, usar escala base
                if not sigla:
                    info = calcular_escala_dia(pm, current_date)
                    sigla = info.get('sigla', '')
                
                # Verificar se é trabalhado
                freq = traduzir_para_frequencia(sigla, info.get('duracao_horas_escala') if 'info' in locals() else None)
                if freq.get('conta_dias', False):
                    dias_trabalhados += 1
            else:
                # Período atual: não contar (evita dupla contagem de meios)
                pass
        else:
            # Se não excluir, contar todos os dias
            alt = AlteracaoEscala.query.filter_by(policial_id=id_pm, data=current_date).first()
            sigla = alt.valor if alt and alt.valor not in ['JM', 'M1', 'M2', 'LIMPAR'] else None
            
            if not sigla:
                info = calcular_escala_dia(pm, current_date)
                sigla = info.get('sigla', '')
            
            freq = traduzir_para_frequencia(sigla, info.get('duracao_horas_escala') if 'info' in locals() else None)
            if freq.get('conta_dias', False):
                dias_trabalhados += 1
        
        current_date += timedelta(days=1)
    
    print(f"DEBUG: Total dias trabalhados na semana: {dias_trabalhados}")  # Debug - comente após testar
    return dias_trabalhados


def calcular_escala_dia(pm, data, registros_semana=None):
    if isinstance(data, datetime):
        data = data.date()

    # Verificar AlteracaoEscala primeiro (prioridade para afastamentos)
    alt = AlteracaoEscala.query.filter_by(policial_id=pm.id, data=data).first()
    if alt and alt.valor and alt.valor.strip().upper() not in ['LIMPAR']:
        valor_manual = alt.valor.strip().upper()
        
        # Siglas de afastamento (lista completa - adicione mais se necessário)
        afastamentos = ['JM', 'M1', 'M2', 'FM', 'DS', 'DR', 'FR', 'PF', 'F*', 'LSV', 'AD', 'LTS', 'CV', 'FN', 'FS', 'LG', 'LA', 'PT', 'NP', 'LP', 'EAP', 'LT', 'Férias']
        
        if valor_manual in afastamentos:
            # Para afastamentos: sigla pura, duração 0, cor warning
            return {
                "sigla": valor_manual,
                "classe": "bg-warning text-dark",
                "descricao": valor_manual,  # Sigla pura (ex.: 'JM')
                "trabalha": False,
                "ciclo_dia": None,
                "regime": getattr(pm, "regime_escala", ""),
                "duracao_horas_escala": 0.0,
            }
        
        elif valor_manual == 'EE':
            # Para EE: calcular duração e mostrar horário
            hora_inicio = alt.hora_inicio_ee or ''
            hora_fim = alt.hora_fim_ee or ''
            duracao = calcular_duracao_hhmm(hora_inicio, hora_fim) if hora_inicio and hora_fim else 0.0
            descricao = f"EE {hora_inicio}–{hora_fim}" if hora_inicio and hora_fim else 'EE'
            return {
                "sigla": "EE",
                "classe": "bg-warning text-dark",
                "descricao": descricao,
                "trabalha": True,  # EE conta como trabalhado
                "ciclo_dia": None,
                "regime": getattr(pm, "regime_escala", ""),
                "duracao_horas_escala": duracao,
            }
        
        elif valor_manual == 'DES':
            # Para DES: adicionar à descrição base, com classe adicional
            info_base = _calcular_escala_base(pm, data)  # Função auxiliar abaixo
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
            # Outras alterações manuais: mostrar valor com prefixo
            return {
                "sigla": valor_manual,
                "classe": "bg-warning text-dark",
                "descricao": f"Alteração manual: {valor_manual}",
                "trabalha": False,
                "ciclo_dia": None,
                "regime": getattr(pm, "regime_escala", ""),
                "duracao_horas_escala": 0.0,
            }

    # Fallback para escala base (sua lógica original, em função auxiliar para clareza)
    return _calcular_escala_base(pm, data)


def _calcular_escala_base(pm, data):
    """
    Lógica original da escala base (sem alterações).
    """
    data_base = getattr(pm, "data_base_escala", None)
    if isinstance(data_base, datetime):
        data_base = data_base.date()

    regime_raw = getattr(pm, "regime_escala", "") or ""
    regime = regime_raw.strip().lower()

    offset = None
    if isinstance(data_base, date) and isinstance(data, date):
        offset = (data - data_base).days

    # Campo de duração padronizado para exportação/frequência
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

def _validar_meios_expedientes_semanais(id_pm, data_inicio, data_fim, novo_valor):
    """
    Valida as regras de meios expedientes por semana.

    Regras:
    - no máximo 2 meios expedientes na mesma semana
    - apenas 1 JM por semana
    - JM não pode coexistir com M1/M2
    - M1 + M2 é permitido no máximo uma vez cada na semana
    - se houver JM, não pode haver M1/M2
    - se houver M1 ou M2, não pode haver JM

    Observação:
    - não conta duas vezes os dias que estão sendo editados
    - substitui corretamente o valor existente pelo novo valor na validação
    """

    novo_valor = (novo_valor or "").strip().upper()

    if novo_valor == "LIMPAR":
        return True, ""

    valores_validos = {"JM", "M1", "M2"}

    # Só valida regras de meio expediente
    if novo_valor not in valores_validos:
        return True, ""

    datas_afetadas = {
        data_inicio + timedelta(days=i)
        for i in range((data_fim - data_inicio).days + 1)
    }

    semanas_verificadas = set()

    for data_alvo in datas_afetadas:
        inicio_semana = data_alvo - timedelta(days=data_alvo.weekday())
        fim_semana = inicio_semana + timedelta(days=6)

        chave_semana = (inicio_semana.year, inicio_semana.month, inicio_semana.day)
        if chave_semana in semanas_verificadas:
            continue
        semanas_verificadas.add(chave_semana)

        alteracoes_semana = AlteracaoEscala.query.filter(
            AlteracaoEscala.policial_id == id_pm,
            AlteracaoEscala.data >= inicio_semana,
            AlteracaoEscala.data <= fim_semana
        ).all()

        meios_encontrados = []

        for alt in alteracoes_semana:
            valor_alt = (alt.valor or "").strip().upper()

            # Se a alteração atual está dentro do período que está sendo editado,
            # ela não deve ser contada como valor antigo
            if alt.data in datas_afetadas:
                continue

            if valor_alt in valores_validos:
                meios_encontrados.append(valor_alt)

        # Adiciona o novo valor apenas uma vez para a validação
        meios_encontrados.append(novo_valor)

        total_meios = len(meios_encontrados)
        quantidade_jm = meios_encontrados.count("JM")
        quantidade_m1 = meios_encontrados.count("M1")
        quantidade_m2 = meios_encontrados.count("M2")

        if total_meios > 2:
            return False, "Não é permitido mais de 2 meios expedientes na mesma semana."

        if quantidade_jm > 1:
            return False, "Só é permitida 1 JM por semana."

        if quantidade_jm > 0 and (quantidade_m1 > 0 or quantidade_m2 > 0):
            return False, "JM não pode coexistir com M1 ou M2 na mesma semana."

        if quantidade_m1 > 1 or quantidade_m2 > 1:
            return False, "M1 e M2 não podem se repetir na mesma semana."

    return True, ""
    
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

def obter_texto_exibicao(sigla, pm=None):
    """
    Retorna o texto correto para exibição na tela/Excel
    """
    if sigla == 'Trabalho' and pm:
        return formatar_horario_turno(
            getattr(pm, 'hora_inicio_escala', None),
            getattr(pm, 'hora_fim_escala', None)
        )
    
    # Para outros casos, usa o STATUS_INFO ou a sigla
    info = STATUS_INFO.get(sigla, {'descricao': sigla})
    return info.get('descricao', sigla)

def verificar_validade_restricoes():
    hoje = datetime.now().date()
    pms_restritos = Policial.query.filter(
        Policial.status_saude == 'Apto Com Restrições',
        Policial.data_fim_restricao != None
    ).all()
    alterados = 0
    for pm in pms_restritos:
        if pm.data_fim_restricao < hoje:
            pm.status_saude = 'Apto Sem Restrições'
            pm.restricao_tipo = None
            pm.data_fim_restricao = None
            alterados += 1
    if alterados > 0: db.session.commit()

def descobrir_afastamento_atual(policial_id):
    """
    Verifica se o PM está com alguma alteração de escala regular (afastamento) HOJE.
    Se estiver, procura a data em que essa alteração termina para retornar à escala normal.
    """
    hoje = datetime.now().date()
    
    # Siglas que consideramos como Afastamento Regular a serem mostrados na tela inicial
    siglas_afastamento = ['DS', 'F*', 'FN', 'FS', 'LTS', 'CV', 'LG', 'PT', 'LA', 'NP', 'LT', 'LP', 'LSV', 'AD'] 
    
    # 1. Verifica se HOJE existe uma alteração cadastrada para ele
    alt_hoje = AlteracaoEscala.query.filter(
        AlteracaoEscala.policial_id == policial_id,
        AlteracaoEscala.data == hoje,
        AlteracaoEscala.valor.in_(siglas_afastamento)
    ).first()
    
    if not alt_hoje:
        return None # Trabalhando normalmente ou de folga padrão
        
    sigla_atual = alt_hoje.valor
    
    # 2. Descobre qual é o ÚLTIMO DIA desse afastamento consecutivo
    # Ele busca alterações futuras do mesmo tipo e ordena pela data mais distante
    ultimo_dia = AlteracaoEscala.query.filter(
        AlteracaoEscala.policial_id == policial_id,
        AlteracaoEscala.data >= hoje,
        AlteracaoEscala.valor == sigla_atual
    ).order_by(desc(AlteracaoEscala.data)).first()
    
    # Formata a data de retorno (Dia D + 1) ou a data fim
    data_fim_fmt = ultimo_dia.data.strftime('%d/%m/%Y') if ultimo_dia else hoje.strftime('%d/%m/%Y')
    
    # Dicionário visual para a tela
    titulos = {
        'F*': 'Folga com Observação', 
        'DS': 'Dispensa de Serviço',
        'FN': 'Férias',
        'FS': 'Falta ao Serv.',
        'CV': 'Convalescença',
        'LG': 'Licença Gestante',
        'PT': 'Licença Paternidade',
        'LA': 'Licença Adoção',
        'NP': 'Núpcias',
        'LT': 'Luto',
        'LP': 'Licença Prêmio',
        'LSV': 'Licença S/ Venc.',
        'LTS': 'Licença Saúde',
        'AD': 'Adido (Outra OPM)'
    }
    
    return {
        'sigla': sigla_atual,
        'titulo': titulos.get(sigla_atual, sigla_atual),
        'data_fim': data_fim_fmt
    }

def serialize_selections(selections):
    return ','.join(selections) if selections else ''

def deserialize_selections(serialized):
    return serialized.split(',') if serialized else []

HABILITACOES_EMOJIS = {
    'Taser / AIN': 'AIN-',
    'Fuzil 556': 'F556-',
    'Fuzil 762': 'F762-',
    'Escopeta Benelli cal 12': 'EB12-',
    'Calibre 12': 'E12'
}

SAT_EMOJIS = {
    'A (Motos)': 'A',
    'B (Carros)': 'B',
    'C (Carga acima de 3.500kg)': 'C',
    'D (Transporte de passageiros)': 'D',
    'E (Veículo Combinado)': 'E'
}

def get_emojis(selections, emoji_dict):
    return ' '.join(emoji_dict.get(sel, '') for sel in selections if sel in emoji_dict).strip()

def get_habilitacoes_texto(pm):
    return get_emojis(deserialize_selections(getattr(pm, 'habilitacoes', None)), HABILITACOES_EMOJIS)

def get_sat_texto(pm):
    return get_emojis(deserialize_selections(getattr(pm, 'sat', None)), SAT_EMOJIS)

def calcular_duracao_hhmm(hora_inicio, hora_fim):
    if not hora_inicio or not hora_fim:
        return None

    inicio = datetime.strptime(hora_inicio, '%H:%M')
    fim = datetime.strptime(hora_fim, '%H:%M')

    if fim <= inicio:
        fim += timedelta(days=1)

    return (fim - inicio).total_seconds() / 3600

# --- NOVAS FUNÇÕES: BLOCOS DE HORÁRIO E CABEÇALHO ---

def formatar_data_militar(data):
    meses = {1:'JAN', 2:'FEV', 3:'MAR', 4:'ABR', 5:'MAI', 6:'JUN', 
             7:'JUL', 8:'AGO', 9:'SET', 10:'OUT', 11:'NOV', 12:'DEZ'}
    ano_curto = data.year % 100
    return f"{data.day:02d}{meses[data.month]}{ano_curto}"

def get_semana_mes(data):
    semana = (data.day - 1) // 7 + 1
    return f"{data.month:02d}.{semana}"

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
        db.session.add(registro)
        versao = 1
    
    db.session.commit()
    return f"V{versao}.0"

def identificar_bloco_horario(pm):
    """
    Define em qual bloco visual o PM vai aparecer na ESCALA SEMANAL.
    Versão preparada para o novo modelo por horário livre.
    """

    funcao = (pm.funcao or "").upper()

    # Se houver horário livre cadastrado, o bloco pode ser inferido pela jornada
    hora_inicio = getattr(pm, 'hora_inicio_escala', None)
    hora_fim = getattr(pm, 'hora_fim_escala', None)

    if hora_inicio and hora_fim:
        inicio_min = hora_inicio.hour * 60 + hora_inicio.minute
        fim_min = hora_fim.hour * 60 + hora_fim.minute

        if fim_min <= inicio_min:
            fim_min += 24 * 60

        duracao = (fim_min - inicio_min) / 60

        # CFP
        if 'CFP' in funcao or 'AUX' in funcao or 'MOTORISTA' in funcao:
            if hora_inicio >= time(16, 45) or hora_inicio >= time(18, 0):
                return "CFP Noturno"
            return "CFP Diurno"

        # Guarda / Sentinela / Permanência
        if 'GUARDA' in funcao or 'SENTINELA' in funcao or 'PERMANÊNCIA' in funcao or 'PERMANENCIA' in funcao:
            if hora_inicio >= time(18, 0):
                return "Guarda do Quartel Noturno"
            return "Guarda do Quartel Diurno"

        # Manutenção / Motomec
        if 'MOTOMEC' in funcao or 'MANUTEN' in funcao or 'VTR' in funcao:
            return "Manutenção/Motomec"

        # Administração por jornada típica
        if duracao < 8:
            return "Administração"
        elif duracao < 12:
            return "Administração"
        elif duracao < 18:
            return "Guarda do Quartel Diurno"
        else:
            return "Outros / A Definir"

    # Fallback por função, caso não haja horários cadastrados
    if 'CFP' in funcao or 'AUX' in funcao or 'MOTORISTA' in funcao:
        return "CFP Diurno"

    if 'S.I.R' in funcao or 'SIR' in funcao:
        return "Guarda do Quartel Diurno"

    if 'GUARDA' in funcao or 'SENTINELA' in funcao or 'PERMANÊNCIA' in funcao or 'PERMANENCIA' in funcao:
        return "Guarda do Quartel Diurno"

    if 'MOTOMEC' in funcao or 'MANUTEN' in funcao or 'VTR' in funcao:
        return "Manutenção/Motomec"

    termos_adm = ['P1', 'P2', 'P3', 'P4', 'P5', 'TESOURARIA', 'SECRETARIA', 'JUSTIÇA', 'CMT', 'SUB', 'CMD']
    if any(t in funcao for t in termos_adm):
        return "Administração"

    return "Outros / A Definir"

def normalizar_secao(valor):
    valor = (valor or '').strip()
    if valor.upper() == 'COORD OP':
        return 'Coord Op'
    if valor.upper() == 'SUB FROTA':
        return 'Sub Frota'
    if valor.upper() == 'DIVADM':
        return 'DivADM'
    if valor.upper() == 'DIVOP':
        return 'DivOP'
    return valor


# ==============================================================================
# ROTAS E VIEWS
# ==============================================================================

@app.route('/')
@login_required
def index():
    verificar_validade_restricoes()
    pms_raw = Policial.query.filter_by(ativo=True).all()
    
    hoje = datetime.now().date()
    
    siglas_afastamento = ['DS', 'F*', 'F', 'FN', 'FS', 'CV', 'LP', 'LSV', 'LTS', 'AD', 'LG', 'LA', 'PT', 'NP', 'EAP', 'LT']
    
    # Nomes bonitos para aparecer na tela inicial em vez da sigla pura
    nomes_afastamento = {
        'DS': 'Dispensa de Serviço',
        'F*': 'Folga com observação', 
        'F': 'Férias',
        'FN': 'Férias', 
        'FS': 'Falta ao Serv.',
        'CV': 'Convalescença',
        'LP': 'Licença Prêmio', 
        'LSV': 'Lic. Sem Venc.',
        'LTS': 'Trat. Saúde', 
        'AD': 'Adido (Outra OPM)',
        'LG': 'Licença Gestante',
        'PT': 'Lic. Paternidade',
        'LA': 'Licença Adoção',
        'NP': 'Núpcias',
        'LT': 'Luto'
    }

    for pm in pms_raw:
        pm.afastamento_atual = None
        pm.data_fim_afastamento = None
        
        # 1. Verifica o que está lançado na escala para HOJE
        alt_hoje = AlteracaoEscala.query.filter_by(policial_id=pm.id, data=hoje).first()
        
        if alt_hoje and alt_hoje.valor in siglas_afastamento:
            sigla = alt_hoje.valor
            pm.afastamento_atual = nomes_afastamento.get(sigla, sigla)
            
            # 2. Descobre quando acaba (lê os dias seguintes no banco)
            alts_futuras = AlteracaoEscala.query.filter(
                AlteracaoEscala.policial_id == pm.id,
                AlteracaoEscala.data >= hoje
            ).order_by(AlteracaoEscala.data).all()
            
            data_fim = hoje
            for alt in alts_futuras:
                # Se a data for consecutiva e a sigla for a mesma, estica a data fim
                if alt.valor == sigla and (alt.data - data_fim).days <= 1:
                    data_fim = alt.data
                elif alt.valor != sigla:
                    break # Mudou de código (ex: de Férias para Trabalho), encerra a busca
                    
            pm.data_fim_afastamento = data_fim

    # Agrupamento e Ordenação
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
        hoje=hoje,
        get_habilitacoes_texto=get_habilitacoes_texto,
        get_sat_texto=get_sat_texto
    )

@app.route('/login', methods=['GET', 'POST'])
def login():
    if current_user.is_authenticated: return redirect(url_for('index'))
    if request.method == 'POST':
        re_login = request.form.get('username')
        senha = request.form.get('password')
        user = Usuario.query.filter_by(username=re_login).first()
        if user and check_password_hash(user.password, senha):
            login_user(user)
            registrar_log('Login', 'Entrou no sistema')
            return redirect(url_for('index'))
        else: flash('RE ou Senha incorretos.', 'error')
    return render_template('login.html')

@app.route('/logout')
@login_required
def logout():
    registrar_log('Logout', 'Saiu')
    logout_user()
    return redirect(url_for('login'))

@app.route('/salvar_pm', methods=['POST'])
def salvar_pm():
    try:
        id_pm = request.form.get('id_pm')
        re = request.form.get('re')
        posto = request.form.get('posto')
        nome = request.form.get('nome')
        nome_guerra = request.form.get('nome_guerra') or None
        cia = request.form.get('cia')
        secao_em = request.form.get('secao_em') or None
        funcao = request.form.get('funcao') or None
        viatura = request.form.get('viatura') or None
        equipe = request.form.get('equipe') or None

        # Novos campos
        habilitacoes = request.form.getlist('habilitacoes')
        sat = request.form.getlist('sat')

        regime_escala = (request.form.get('regime_escala') or '').strip()
        if regime_escala:
            mapa_regimes = {
                'administrativo': 'Administrativo',
                '12x36': '12x36',
                '12x24_12x48': '12x24_12x48',
                '12x60': '12x60',
                '12x36_12x60': '12x36_12x60',
                '24x48': '24x48'
            }
            regime_escala = mapa_regimes.get(regime_escala.lower(), regime_escala)
        else:
            regime_escala = None

        data_base_escala_str = request.form.get('data_base_escala') or None

        hora_inicio_escala = request.form.get('hora_inicio_escala') or None
        hora_fim_escala = request.form.get('hora_fim_escala') or None
        duracao_horas_escala = request.form.get('duracao_horas_escala') or None
        status_saude = request.form.get('status_saude') or 'Apto Sem Restrições'
        restricao_tipo = request.form.get('restricao_tipo') or None
        data_fim_restricao = request.form.get('data_fim_restricao') or None

        # Converte a data-base da escala para date real
        data_base_escala = None
        if data_base_escala_str:
            data_base_escala = datetime.strptime(data_base_escala_str, '%Y-%m-%d').date()

        # Converte a data final da restrição
        if data_fim_restricao:
            data_fim_restricao = datetime.strptime(data_fim_restricao, '%Y-%m-%d').date()

        # Calcula duração automaticamente, se necessário
        if hora_inicio_escala and hora_fim_escala:
            hi = datetime.strptime(hora_inicio_escala, '%H:%M').time()
            hf = datetime.strptime(hora_fim_escala, '%H:%M').time()

            if not duracao_horas_escala:
                inicio_min = hi.hour * 60 + hi.minute
                fim_min = hf.hour * 60 + hf.minute

                if fim_min <= inicio_min:
                    fim_min += 24 * 60

                duracao_horas_escala = round((fim_min - inicio_min) / 60, 2)
        else:
            hora_inicio_escala = None
            hora_fim_escala = None
            duracao_horas_escala = None

        if id_pm:
            pm = db.session.get(Policial, int(id_pm))
            if not pm:
                flash('PM não encontrado.', 'danger')
                return redirect(url_for('index'))

            pm.re = re
            pm.posto = posto
            pm.nome = nome
            pm.nome_guerra = nome_guerra
            pm.cia = cia
            pm.secao_em = secao_em
            pm.funcao = funcao
            pm.viatura = viatura
            pm.equipe = equipe
            pm.regime_escala = regime_escala
            pm.data_base_escala = data_base_escala
            pm.hora_inicio_escala = datetime.strptime(hora_inicio_escala, '%H:%M').time() if hora_inicio_escala else None
            pm.hora_fim_escala = datetime.strptime(hora_fim_escala, '%H:%M').time() if hora_fim_escala else None
            pm.duracao_horas_escala = float(duracao_horas_escala) if duracao_horas_escala not in [None, ''] else None
            pm.status_saude = status_saude
            pm.restricao_tipo = restricao_tipo
            pm.data_fim_restricao = data_fim_restricao

            # Salva habilitações e SAT como string separada por vírgulas
            pm.habilitacoes = ','.join(habilitacoes) if habilitacoes else ''
            pm.sat = ','.join(sat) if sat else ''

            db.session.commit()
            flash('✅ PM atualizado com sucesso!', 'success')

        else:
            novo_pm = Policial(
                posto=posto,
                re=re,
                nome=nome,
                nome_guerra=nome_guerra,
                cia=cia,
                secao_em=secao_em,
                tipo_escala=None,
                data_base_escala=data_base_escala,
                regime_escala=regime_escala,
                funcao=funcao,
                equipe=equipe,
                viatura=viatura,
                ativo=True,
                hora_inicio_escala=datetime.strptime(hora_inicio_escala, '%H:%M').time() if hora_inicio_escala else None,
                hora_fim_escala=datetime.strptime(hora_fim_escala, '%H:%M').time() if hora_fim_escala else None,
                duracao_horas_escala=float(duracao_horas_escala) if duracao_horas_escala not in [None, ''] else None,
                status_saude=status_saude,
                restricao_tipo=restricao_tipo,
                data_fim_restricao=data_fim_restricao,
                habilitacoes=','.join(habilitacoes) if habilitacoes else '',
                sat=','.join(sat) if sat else ''
            )
            db.session.add(novo_pm)
            db.session.commit()
            flash('✅ PM cadastrado com sucesso!', 'success')

        return redirect(url_for('index'))

    except Exception as e:
        db.session.rollback()
        flash(f'Erro ao salvar PM: {e}', 'danger')
        return redirect(url_for('index'))

@app.route('/consultar_pm', methods=['GET', 'POST'])
@login_required
def consultar_pm():
    pm = None; historico_mov = []; historico_afast = []
    if request.method == 'POST':
        re_busca = request.form.get('re_busca', '').strip()
        pm = Policial.query.filter_by(re=re_busca).first()
        if pm:
            historico_mov = HistoricoMovimentacao.query.filter_by(policial_id=pm.id).order_by(desc(HistoricoMovimentacao.data_mudanca)).all()
            siglas_comuns = ['S', 'S/1', 'S/2', 'S/3', 'S/4', 'S/5', 'S/6', 'M1', 'M2', 'SS', 'SR', 'PPJMD']
            historico_afast = AlteracaoEscala.query.filter(
                AlteracaoEscala.policial_id == pm.id,
                AlteracaoEscala.valor.not_in(siglas_comuns)
            ).order_by(desc(AlteracaoEscala.data)).all()
        else:
            flash('Policial não encontrado com este RE.', 'warning')
    return render_template('consultar_pm.html', pm=pm, historico_mov=historico_mov, historico_afast=historico_afast)

@app.route('/atualizar_saude', methods=['POST'])
@login_required
def atualizar_saude():
    id_pm = request.form.get('id_pm')
    status = request.form.get('status_saude')
    tipo = request.form.get('restricao_tipo')
    data_fim = request.form.get('data_fim_restricao')
    
    pm = db.session.get(Policial, id_pm)
    if pm:
        pm.status_saude = status
        if status == 'Apto Com Restrições':
            pm.restricao_tipo = tipo
            if data_fim: pm.data_fim_restricao = datetime.strptime(data_fim, '%Y-%m-%d').date()
        else:
            pm.restricao_tipo = None; pm.data_fim_restricao = None
        db.session.commit()
        flash('Situação de saúde atualizada!', 'success')
    return redirect(url_for('index'))

@app.route('/escala/<int:id_pm>')
@login_required
def escala(id_pm):
    try:
        pm = db.session.get(Policial, id_pm)
        if not pm:
            flash('Policial não encontrado.', 'danger')
            return redirect(url_for('index'))

        mes = request.args.get('mes', date.today().month, type=int)
        ano = request.args.get('ano', date.today().year, type=int)

        if mes < 1 or mes > 12 or ano < 2000 or ano > 2100:
            flash('Mês ou ano inválido.', 'danger')
            return redirect(url_for('index'))

        data_atual = date(ano, mes, 1)
        prox_mes = data_atual.replace(day=28) + timedelta(days=4)
        prox_mes = prox_mes.replace(day=1)
        ant_mes = data_atual.replace(day=1) - timedelta(days=1)

        nav = type('Nav', (), {
            'ant_m': ant_mes.month,
            'ant_a': ant_mes.year,
            'prox_m': prox_mes.month,
            'prox_a': prox_mes.year
        })()

        dias_vazios = (date(ano, mes, 1).weekday() + 1) % 7
        ultimo_dia = calendar.monthrange(ano, mes)[1]

        alteracoes = AlteracaoEscala.query.filter_by(policial_id=id_pm).all()
        alt_map = {
            alt.data: alt for alt in alteracoes
            if alt.data.year == ano and alt.data.month == mes
        }

        escala_dias = []
        totais = type('Totais', (), {'trabalhado': 0, 'soma_valor': 0})()

        for d in range(1, ultimo_dia + 1):
            data_dia = date(ano, mes, d)
            dia_obj = type('Dia', (), {})()
            dia_obj.dia = d
            dia_obj.data_iso = data_dia.isoformat()

            info = calcular_escala_dia(pm, data_dia)
            freq = traduzir_para_frequencia(info.get('sigla'), info.get('duracao_horas_escala'))

            classe = info.get('classe')
            descricao = info.get('descricao')
            duracao_horas_escala = info.get('duracao_horas_escala')
            texto_exibicao = None

            if data_dia in alt_map:
                alt_obj = alt_map[data_dia]
                valor_manual = (alt_obj.valor or '').strip().upper()

                if valor_manual == 'LIMPAR':
                    texto_exibicao = descricao

                elif valor_manual == 'TRABALHO':
                    classe = 'bg-success text-white'
                    texto_exibicao = formatar_horario_turno(
                        getattr(pm, 'hora_inicio_escala', None),
                        getattr(pm, 'hora_fim_escala', None)
                    )
                    descricao = texto_exibicao
                    duracao_horas_escala = getattr(pm, 'duracao_horas_escala', None)
                    freq = traduzir_para_frequencia('TRABALHO', duracao_horas_escala)

                elif valor_manual == 'EE':
                    classe = 'bg-warning text-dark'
                    duracao_horas_escala = calcular_duracao_hhmm(
                        alt_obj.hora_inicio_ee,
                        alt_obj.hora_fim_ee
                    )
                    texto_exibicao = f"EE {alt_obj.hora_inicio_ee}–{alt_obj.hora_fim_ee}"
                    descricao = texto_exibicao
                    freq = traduzir_para_frequencia('TRABALHO', duracao_horas_escala)

                elif valor_manual == 'DES':
                    classe = f"{classe} des-marcado" if classe else 'des-marcado'
                    texto_exibicao = f"{descricao} DES" if descricao else 'DES'
                    descricao = texto_exibicao

                else:
                    # Para todos os afastamentos e outros manuais
                    if valor_manual in ['JM', 'M1', 'M2', 'FM', 'DS', 'DR', 'FR', 'PF', 'F*', 'LSV', 'AD', 'LTS', 'CV', 'FN', 'FS', 'LG', 'LA', 'PT', 'NP', 'LP', 'EAP', 'LT', 'Férias']:
                        classe = 'bg-warning text-dark'
                        descricao = valor_manual  # Sigla pura
                        texto_exibicao = valor_manual
                        duracao_horas_escala = 0
                        freq = traduzir_para_frequencia(valor_manual, 0)
                    else:
                        classe = 'bg-warning text-dark'
                        descricao = f'Alteração manual: {valor_manual}'
                        texto_exibicao = descricao

            dia_obj.valor = texto_exibicao if texto_exibicao else freq['codigo']
            dia_obj.classe = classe
            dia_obj.descricao = descricao
            dia_obj.duracao_horas_escala = duracao_horas_escala
            dia_obj.trabalha = freq['conta_dias']

            if freq['conta_dias']:
                totais.trabalhado += 1
                if freq['diarias'] is not None:
                    totais.soma_valor += freq['diarias']

            escala_dias.append(dia_obj)

        validacao = ValidacaoMensal.query.filter_by(policial_id=id_pm, mes=mes, ano=ano).first()

        # Agrupamento por seção para visualização geral (se aplicável)
        secao_pm = pm.cia or pm.secao_em or 'Outros'
        blocos_secao = agrupar_por_secao([pm], ORDEM_SECOES = ['Comando', 'Subcomando', 'Coord Op','P1', 'P2', 'P3', 'P4', 'P5','UGE', 'UIS', 'DivOP', 'Ordenança', 'GT','Sub Frota', 'DivADM', 'SPJMD', 'Secretaria', 'Motomec'])

        return render_template(
            'editar.html',
            pm=pm,
            escala=escala_dias,
            dias_vazios=dias_vazios,
            mes=f'{mes:02d}',
            ano=ano,
            nav=nav,
            validacao=validacao,
            totais=totais,
            secao=secao_pm,
            blocos_secao=blocos_secao  # Passa para template se necessário
        )

    except Exception as e:
        flash(f'Erro ao carregar escala: {e}', 'danger')
        return redirect(url_for('index'))


@app.route('/salvar_alteracao', methods=['POST'])
@login_required
def salvar_alteracao():
    try:
        id_pm = int(request.form.get('id_pm'))
        inicio = datetime.strptime(request.form.get('data_inicio'), '%Y-%m-%d').date()
        fim_str = request.form.get('data_fim')
        fim = datetime.strptime(fim_str, '%Y-%m-%d').date() if fim_str else inicio
        valor = (request.form.get('valor') or '').strip().upper()

        hora_inicio_ee = (request.form.get('hora_inicio_ee') or '').strip()
        hora_fim_ee = (request.form.get('hora_fim_ee') or '').strip()

        pm = db.session.get(Policial, id_pm)
        if not pm:
            flash('Erro: Policial não encontrado.', 'error')
            return redirect(request.referrer)

        mes_inicio = inicio.month
        ano_inicio = inicio.year
        validacao = ValidacaoMensal.query.filter_by(policial_id=id_pm, mes=mes_inicio, ano=ano_inicio).first()
        e_admin = getattr(current_user, 'perfil', 'usuario') == 'admin'

        if validacao and not e_admin:
            flash(
                f'A escala deste mês já foi validada por {validacao.admin_info}. Apenas o administrador pode alterá-la.',
                'danger'
            )
            return redirect(request.referrer)

        if valor in {'JM', 'M1', 'M2'}:
            # Chama a função corrigida (calcula tudo internamente, passando pm)
            ok, mensagem = validar_meios_expedientes_semanais(id_pm, inicio, fim, valor)
            if not ok:
                flash(mensagem, 'danger')
                return redirect(request.referrer)

        if valor == 'EE':
            if not hora_inicio_ee or not hora_fim_ee:
                flash('Erro: informe hora de início e fim para EE.', 'danger')
                return redirect(request.referrer)

        for i in range((fim - inicio).days + 1):
            dia = inicio + timedelta(days=i)
            alt = AlteracaoEscala.query.filter_by(policial_id=id_pm, data=dia).first()

            if valor == 'LIMPAR':
                if alt:
                    db.session.delete(alt)
            else:
                if not alt:
                    alt = AlteracaoEscala(policial_id=id_pm, data=dia, valor=valor)
                    db.session.add(alt)
                else:
                    alt.valor = valor

                if valor == 'EE':
                    alt.hora_inicio_ee = hora_inicio_ee
                    alt.hora_fim_ee = hora_fim_ee
                else:
                    alt.hora_inicio_ee = None
                    alt.hora_fim_ee = None

        db.session.commit()
        flash('Escala atualizada com sucesso!', 'success')

    except Exception as e:
        db.session.rollback()
        flash(f'Erro interno: {e}', 'danger')

    return redirect(request.referrer)

@app.route('/validar_escala/<int:id_pm>/<int:mes>/<int:ano>', methods=['POST'])
@login_required
def validar_escala(id_pm, mes, ano):
    if getattr(current_user, 'perfil', 'usuario') != 'admin':
        flash('Acesso negado. Apenas o Chefe P1 (Administrador) pode validar escalas.', 'danger')
        return redirect(request.referrer)
        
    # Busca o Posto e Nome de Guerra do Admin logado cruzando o RE
    admin_pm = Policial.query.filter_by(re=current_user.re).first()
    if admin_pm:
        admin_info = f"{admin_pm.posto} {admin_pm.nome_guerra} (RE {admin_pm.re})"
    else:
        admin_info = f"{current_user.nome} (RE {current_user.re})"
        
    validacao = ValidacaoMensal.query.filter_by(policial_id=id_pm, mes=mes, ano=ano).first()
    if not validacao:
        nova_validacao = ValidacaoMensal(
            policial_id=id_pm, mes=mes, ano=ano,
            admin_id=current_user.id, admin_info=admin_info
        )
        db.session.add(nova_validacao)
        db.session.commit()
        flash(f'✅ Escala de {mes:02d}/{ano} validada com sucesso por {admin_info}!', 'success')
        
    return redirect(request.referrer)

@app.route('/revogar_validacao/<int:id_pm>/<int:mes>/<int:ano>', methods=['POST'])
@login_required
def revogar_validacao(id_pm, mes, ano):
    if getattr(current_user, 'perfil', 'usuario') != 'admin':
        return redirect(request.referrer)
        
    validacao = ValidacaoMensal.query.filter_by(policial_id=id_pm, mes=mes, ano=ano).first()
    if validacao:
        db.session.delete(validacao)
        db.session.commit()
        flash('🔓 Validação revogada. A escala está aberta para edições novamente.', 'info')
        
    return redirect(request.referrer)

@app.route('/excluir_pm/<int:id_pm>')
@login_required
def excluir_pm(id_pm):
    pm = db.session.get(Policial, id_pm)
    if pm:
        pm.ativo = False
        db.session.commit()
        flash('Policial inativado.', 'success')
    return redirect(url_for('index'))

# ==============================================================================
# NOVAS ROTAS DE GERAÇÃO (EXCEL E RELATÓRIOS)
# ==============================================================================

@app.route('/visualizar_escala', methods=['POST'])
@login_required
def visualizar_escala():
    try:
        cia = request.form.get('cia_filtro')
        inicio_str = request.form.get('data_inicio_semana')
        tipo_filtro = request.form.get('tipo_filtro', 'TODOS')
        tipo_escala = (request.form.get('tipo_escala') or 'semanal').strip().lower()

        if not cia or not inicio_str:
            return redirect(url_for('index'))

        inicio = datetime.strptime(inicio_str, '%Y-%m-%d').date()

        if tipo_escala == 'diaria':
            return visualizar_escala_diaria(cia=cia, inicio=inicio, tipo_filtro=tipo_filtro)

        return visualizar_escala_semanal(cia=cia, inicio=inicio, tipo_filtro=tipo_filtro)

    except Exception as e:
        flash(f'Erro ao visualizar escala: {e}', 'danger')
        return redirect(url_for('index'))


def visualizar_escala_semanal(cia, inicio, tipo_filtro='TODOS'):
    try:
        fim = inicio + timedelta(days=6)

        tipo_txt = (
            "OFICIAIS" if tipo_filtro == 'OFICIAIS'
            else "PRAÇAS" if tipo_filtro == 'PRACAS'
            else "CFP" if tipo_filtro == 'CFP'
            else "POLICIAIS MILITARES"
        )

        numerador = f"{get_semana_mes(inicio)}/{str(inicio.year)[2:]}"
        periodo_txt = f"DE {formatar_data_militar(inicio)} A {formatar_data_militar(fim)}"

        dias_semana = []
        nomes_dias = ['SEG', 'TER', 'QUA', 'QUI', 'SEX', 'SAB', 'DOM']

        for i in range(7):
            d = inicio + timedelta(days=i)
            dias_semana.append({
                'dia_semana': nomes_dias[i],
                'data_str': f"{d.day:02d}/{d.month:02d}"
            })

        pms_raw = Policial.query.filter_by(cia=cia, ativo=True).all()
        pms = filtrar_por_tipo_hierarquico(pms_raw, tipo_filtro)

        if cia == 'EM':
            pms = ordenar_em_agrupado(pms)
        else:
            pms = ordenar_cia_agrupado(pms)

        mapa = agrupar_por_secao(pms, tipo_filtro)
        ordem = ordem_blocos_para_tipo(tipo_filtro)

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
                        is_afastamento = False
                    elif sigla_dia == 'FOLGA':
                        valor = 'FOLGA'
                        is_afastamento = False
                    else:
                        valor = descricao_dia if descricao_dia else sigla_dia
                        is_afastamento = (sigla_dia not in ('TRABALHO', 'FOLGA'))

                    pm.dias.append({
                        'valor': valor,
                        'is_afastamento': is_afastamento
                    })

                pms_bloco.append(pm)

            if pms_bloco:
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
            blocos=blocos
        )

    except Exception as e:
        flash(f'Erro ao carregar visualização semanal: {e}', 'danger')
        return redirect(url_for('index'))


def visualizar_escala_diaria(cia, inicio, tipo_filtro='TODOS'):
    try:
        tipo_txt = (
            "OFICIAIS" if tipo_filtro == 'OFICIAIS'
            else "PRAÇAS" if tipo_filtro == 'PRACAS'
            else "CFP" if tipo_filtro == 'CFP'
            else "POLICIAIS MILITARES"
        )

        numerador = f"{get_semana_mes(inicio)}/{str(inicio.year)[2:]}"
        periodo_txt = f"ESCALA DIÁRIA DE {formatar_data_militar(inicio)}"

        pms_raw = Policial.query.filter_by(cia=cia, ativo=True).all()
        pms = filtrar_por_tipo_hierarquico(pms_raw, tipo_filtro)

        if cia == 'EM':
            pms = ordenar_em_agrupado(pms)
        else:
            pms = ordenar_cia_agrupado(pms)

        mapa = agrupar_por_secao(pms, tipo_filtro)
        ordem = ordem_blocos_para_tipo(tipo_filtro)

        blocos = []
        secoes_vistas = set()

        for nome_bloco in ordem:
            if nome_bloco in secoes_vistas:
                continue

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
                pms_bloco.append(pm)

            if pms_bloco:
                blocos.append({
                    'nome_bloco': nome_bloco,
                    'horario_turno': horario_turno,
                    'pms': pms_bloco
                })
                secoes_vistas.add(nome_bloco)

        return render_template(
            'visualizar_escala_diaria.html',
            cia=cia,
            inicio=inicio,
            tipo_filtro=tipo_filtro,
            tipo_txt=tipo_txt,
            numerador=numerador,
            periodo_txt=periodo_txt,
            blocos=blocos
        )

    except Exception as e:
        flash(f'Erro ao carregar visualização diária: {e}', 'danger')
        return redirect(url_for('index'))

@app.route('/gerar_escala', methods=['POST'])
@login_required
def gerar_escala():
    try:
        cia = request.form.get('cia_filtro')
        inicio_str = request.form.get('data_inicio_semana')
        tipo_filtro = request.form.get('tipo_filtro', 'TODOS')
        tipo_escala = (request.form.get('tipo_escala') or 'semanal').strip().lower()

        if not cia or not inicio_str:
            return redirect(url_for('index'))

        inicio = datetime.strptime(inicio_str, '%Y-%m-%d').date()

        if tipo_escala == 'diaria':
            return _gerar_escala_diaria(cia, inicio, tipo_filtro)

        return _gerar_escala_semanal(cia, inicio, tipo_filtro)

    except Exception as e:
        flash(f'Erro ao gerar escala: {e}', 'danger')
        return redirect(url_for('index'))

@app.route('/gerar_escala_semanal', methods=['POST'])
@login_required
def gerar_escala_semanal():
    try:
        cia = request.form.get('cia_filtro')
        inicio_str = request.form.get('data_inicio_semana')
        tipo_filtro = request.form.get('tipo_filtro', 'TODOS')

        if not cia or not inicio_str:
            return redirect(url_for('index'))

        inicio = datetime.strptime(inicio_str, '%Y-%m-%d').date()
        return _gerar_escala_semanal(cia, inicio, tipo_filtro)

    except Exception as e:
        flash(f'Erro ao gerar escala semanal: {e}', 'danger')
        return redirect(url_for('index'))


def _gerar_escala_semanal(cia, inicio, tipo_filtro='TODOS'):
    try:
        fim = inicio + timedelta(days=6)
        versao_txt = get_proxima_versao(cia, inicio, tipo_filtro)

        tipo_txt = (
            "OFICIAIS" if tipo_filtro == 'OFICIAIS'
            else "PRAÇAS" if tipo_filtro == 'PRACAS'
            else "CFP" if tipo_filtro == 'CFP'
            else "POLICIAIS MILITARES"
        )

        numerador = f"{get_semana_mes(inicio)}/{str(inicio.year)[2:]}"
        periodo_txt = f"DE {formatar_data_militar(inicio)} A {formatar_data_militar(fim)}"

        output = BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        ws = workbook.add_worksheet("ESCALA")

        ws.set_paper(9)
        ws.set_landscape()
        ws.set_margins(0.5, 0.5, 0.5, 0.5)

        fmt_cabecalho = workbook.add_format({
            'bold': True, 'font_name': 'Arial', 'font_size': 10,
            'align': 'center', 'valign': 'vcenter'
        })
        fmt_numerador = workbook.add_format({
            'bold': True, 'font_name': 'Arial', 'font_size': 10,
            'align': 'center', 'valign': 'vcenter', 'underline': True
        })
        fmt_bloco = workbook.add_format({
            'bold': True, 'font_size': 9, 'align': 'center', 'valign': 'vcenter',
            'bg_color': '#595959', 'font_color': 'white', 'border': 1
        })
        fmt_th = workbook.add_format({
            'bold': True, 'font_size': 8, 'align': 'center', 'border': 1,
            'bg_color': '#EFEFEF', 'text_wrap': True
        })
        fmt_dia = workbook.add_format({
            'bold': True, 'font_size': 8, 'align': 'center', 'border': 1,
            'bg_color': '#D9D9D9'
        })
        fmt_center = workbook.add_format({
            'font_size': 8, 'align': 'center', 'border': 1
        })
        fmt_left = workbook.add_format({
            'font_size': 8, 'align': 'left', 'border': 1
        })
        fmt_afastamento = workbook.add_format({
            'font_size': 8, 'align': 'center', 'border': 1, 'bg_color': '#FFFFCC'
        })

        ws.set_column('A:A', 6)
        ws.set_column('B:B', 10)
        ws.set_column('C:C', 8)
        ws.set_column('D:D', 25)
        ws.set_column('E:E', 8)
        ws.set_column('F:G', 5)
        ws.set_column('H:H', 12)
        ws.set_column('I:I', 8)
        ws.set_column('J:P', 6)

        img_path = os.path.join(basedir, 'brasao.png')
        if os.path.exists(img_path):
            ws.insert_image('A1', img_path, {
                'x_scale': 0.22, 'y_scale': 0.22,
                'x_offset': 15, 'y_offset': 5
            })

        ws.merge_range('C1:P1', "SECRETARIA DA SEGURANÇA PÚBLICA", fmt_cabecalho)
        ws.merge_range('C2:P2', "POLÍCIA MILITAR DO ESTADO DE SÃO PAULO", fmt_cabecalho)
        ws.merge_range('C3:P3', "COMANDO DE POLICIAMENTO DE ÁREA METROPOLITANA DEZ", fmt_cabecalho)
        ws.merge_range('C4:P4', "TRIGÉSIMO SÉTIMO BATALHÃO DE POLICIA MILITAR METROPOLITANO", fmt_cabecalho)

        texto_escala = f"ESCALA DE SERVIÇO DE {tipo_txt} Nº 37BPMM-{numerador} - {versao_txt}"
        ws.merge_range('A6:P6', texto_escala, fmt_numerador)
        ws.merge_range('A7:P7', periodo_txt, fmt_cabecalho)

        linha = 8
        cols = ["EQ", "POSTO", "RE", "NOME", "VTR", "ENT", "SAI", "FUNÇÃO", "SEÇÃO"]
        for i, c in enumerate(cols):
            ws.write(linha, i, c, fmt_th)

        dias_semana = ['SEG', 'TER', 'QUA', 'QUI', 'SEX', 'SAB', 'DOM']
        for i in range(7):
            d = inicio + timedelta(days=i)
            ws.write(linha, 9 + i, f"{dias_semana[i]}\n{d.day}/{d.month}", fmt_dia)

        linha += 1

        pms_raw = Policial.query.filter_by(cia=cia, ativo=True).all()
        pms = filtrar_por_tipo_hierarquico(pms_raw, tipo_filtro)

        if cia == 'EM':
            pms = ordenar_em_agrupado(pms)
        else:
            pms = ordenar_cia_agrupado(pms)

        mapa = agrupar_por_secao(pms, tipo_filtro)
        ordem = ordem_blocos_para_tipo(tipo_filtro)

        for nome_bloco in ordem:
            lista_pms = mapa.get(nome_bloco, [])
            if not lista_pms:
                continue

            ws.merge_range(f'A{linha}:P{linha}', f'BLOCO: {nome_bloco}', fmt_bloco)
            linha += 1

            for pm in lista_pms:
                ws.write(linha, 0, getattr(pm, 'equipe', '') or '', fmt_center)
                ws.write(linha, 1, obter_posto_pm(pm), fmt_center)
                ws.write(linha, 2, getattr(pm, 're', '') or '', fmt_center)
                ws.write(linha, 3, getattr(pm, 'nome_guerra', None) or getattr(pm, 'nome', '') or '', fmt_left)
                ws.write(linha, 4, getattr(pm, 'viatura', '') or '', fmt_center)
                ws.write(linha, 5, "", fmt_center)
                ws.write(linha, 6, "", fmt_center)
                ws.write(linha, 7, getattr(pm, 'funcao', '') or '', fmt_center)
                ws.write(linha, 8, getattr(pm, 'secao_em', '') or '', fmt_center)

                for i in range(7):
                    data_dia = inicio + timedelta(days=i)
                    info_dia = calcular_escala_dia(pm, data_dia)
                    sigla = (info_dia.get('sigla') or '').strip().upper()
                    descricao = info_dia.get('descricao', sigla)

                    if sigla == 'TRABALHO':
                        vis = obter_texto_exibicao('Trabalho', pm)
                        estilo = fmt_center
                    elif sigla == 'FOLGA':
                        vis = 'FOLGA'
                        estilo = fmt_center
                    else:
                        vis = descricao if descricao else sigla
                        estilo = fmt_afastamento if is_afastamento(sigla) else fmt_center

                    ws.write(linha, 9 + i, vis, estilo)

                linha += 1

        linha += 3
        ws.merge_range(f'D{linha}:H{linha}', '_' * 30, fmt_center)
        ws.merge_range(f'D{linha+1}:H{linha+1}', 'ENCARREGADO DE ESCALA', fmt_center)
        ws.merge_range(f'K{linha}:O{linha}', '_' * 30, fmt_center)
        ws.merge_range(f'K{linha+1}:O{linha+1}', f'CMT DA {cia}', fmt_center)

        workbook.close()
        output.seek(0)

        return send_file(
            output,
            download_name=f"Escala_{cia}_{versao_txt}.xlsx",
            as_attachment=True
        )

    except Exception as e:
        flash(f'Erro: {e}', 'error')
        return redirect(url_for('index'))
    
@app.route('/gerar_escala_diaria', methods=['POST'])
@login_required
def gerar_escala_diaria():
    try:
        cia = request.form.get('cia_filtro')
        inicio_str = request.form.get('data_inicio_semana')
        tipo_filtro = request.form.get('tipo_filtro', 'TODOS')

        if not cia or not inicio_str:
            return redirect(url_for('index'))

        inicio = datetime.strptime(inicio_str, '%Y-%m-%d').date()
        return _gerar_escala_diaria(cia, inicio, tipo_filtro)

    except Exception as e:
        flash(f'Erro ao gerar escala diária: {e}', 'danger')
        return redirect(url_for('index'))


def _gerar_escala_diaria(cia, inicio, tipo_filtro='TODOS'):
    try:
        versao_txt = get_proxima_versao(cia, inicio, tipo_filtro)

        tipo_txt = (
            "OFICIAIS" if tipo_filtro == 'OFICIAIS'
            else "PRAÇAS" if tipo_filtro == 'PRACAS'
            else "CFP" if tipo_filtro == 'CFP'
            else "POLICIAIS MILITARES"
        )

        numerador = f"{get_semana_mes(inicio)}/{str(inicio.year)[2:]}"
        periodo_txt = f"ESCALA DIÁRIA DE {formatar_data_militar(inicio)}"

        output = BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        ws = workbook.add_worksheet("ESCALA DIÁRIA")

        ws.set_paper(9)
        ws.set_landscape()
        ws.set_margins(0.35, 0.35, 0.35, 0.35)
        ws.fit_to_pages(1, 0)
        ws.repeat_rows(0, 8)

        fmt_cabecalho = workbook.add_format({
            'bold': True, 'font_name': 'Arial', 'font_size': 10,
            'align': 'center', 'valign': 'vcenter'
        })
        fmt_numerador = workbook.add_format({
            'bold': True, 'font_name': 'Arial', 'font_size': 10,
            'align': 'center', 'valign': 'vcenter', 'underline': True
        })
        fmt_bloco = workbook.add_format({
            'bold': True, 'font_size': 9, 'align': 'center', 'valign': 'vcenter',
            'bg_color': '#595959', 'font_color': 'white', 'border': 1
        })
        fmt_th = workbook.add_format({
            'bold': True, 'font_size': 8, 'align': 'center', 'border': 1,
            'bg_color': '#EFEFEF', 'text_wrap': True, 'valign': 'vcenter'
        })
        fmt_center = workbook.add_format({
            'font_size': 8, 'align': 'center', 'valign': 'vcenter', 'border': 1
        })
        fmt_left = workbook.add_format({
            'font_size': 8, 'align': 'left', 'valign': 'vcenter', 'border': 1
        })
        fmt_servico = workbook.add_format({
            'font_size': 8, 'align': 'center', 'valign': 'vcenter',
            'border': 1, 'bg_color': '#D9EAD3'
        })
        fmt_folga = workbook.add_format({
            'font_size': 8, 'align': 'center', 'valign': 'vcenter',
            'border': 1, 'bg_color': '#F3F3F3'
        })
        fmt_afastamento = workbook.add_format({
            'font_size': 8, 'align': 'center', 'valign': 'vcenter',
            'border': 1, 'bg_color': '#FFFFCC'
        })

        ws.set_column('A:A', 5)
        ws.set_column('B:B', 11)
        ws.set_column('C:C', 8)
        ws.set_column('D:D', 22)
        ws.set_column('E:E', 9)
        ws.set_column('F:F', 18)
        ws.set_column('G:G', 12)
        ws.set_column('H:H', 14)
        ws.set_column('I:I', 12)
        ws.set_column('J:J', 12)

        img_path = os.path.join(basedir, 'brasao.png')
        if os.path.exists(img_path):
            ws.insert_image('A1', img_path, {
                'x_scale': 0.20, 'y_scale': 0.20,
                'x_offset': 12, 'y_offset': 4
            })

        ws.merge_range('C1:J1', "SECRETARIA DA SEGURANÇA PÚBLICA", fmt_cabecalho)
        ws.merge_range('C2:J2', "POLÍCIA MILITAR DO ESTADO DE SÃO PAULO", fmt_cabecalho)
        ws.merge_range('C3:J3', "COMANDO DE POLICIAMENTO DE ÁREA METROPOLITANA DEZ", fmt_cabecalho)
        ws.merge_range('C4:J4', "TRIGÉSIMO SÉTIMO BATALHÃO DE POLÍCIA MILITAR METROPOLITANO", fmt_cabecalho)

        texto_escala = f"ESCALA DIÁRIA DE SERVIÇO DE {tipo_txt} Nº 37BPMM-{numerador} - {versao_txt}"
        ws.merge_range('A6:J6', texto_escala, fmt_numerador)
        ws.merge_range('A7:J7', periodo_txt, fmt_cabecalho)

        linha = 8
        cols = ["EQ", "POSTO", "RE", "NOME", "VTR", "HAB", "SAT", "FUNÇÃO", "SEÇÃO", "STATUS"]
        for i, c in enumerate(cols):
            ws.write(linha, i, c, fmt_th)

        linha += 1

        pms_raw = Policial.query.filter_by(cia=cia, ativo=True).all()
        pms = filtrar_por_tipo_hierarquico(pms_raw, tipo_filtro)

        if cia == 'EM':
            pms = ordenar_em_agrupado(pms)
        else:
            pms = ordenar_cia_agrupado(pms)

        mapa = agrupar_por_secao(pms, tipo_filtro)
        ordem = ordem_blocos_para_tipo(tipo_filtro)

        for nome_bloco in ordem:
            lista_pms = mapa.get(nome_bloco, [])
            if not lista_pms:
                continue

            ws.merge_range(f'A{linha}:J{linha}', f'BLOCO: {nome_bloco}', fmt_bloco)
            linha += 1

            for pm in lista_pms:
                info_dia = calcular_escala_dia(pm, inicio)
                sigla = (info_dia.get('sigla') or '').strip().upper()
                descricao = info_dia.get('descricao', sigla)

                if sigla == 'TRABALHO':
                    vis = obter_texto_exibicao('Trabalho', pm)
                    estilo = fmt_servico
                elif sigla == 'FOLGA':
                    vis = 'FOLGA'
                    estilo = fmt_folga
                else:
                    vis = descricao if descricao else sigla
                    estilo = fmt_afastamento if is_afastamento(sigla) else fmt_folga

                ws.write(linha, 0, getattr(pm, 'equipe', '') or '', fmt_center)
                ws.write(linha, 1, obter_posto_pm(pm), fmt_center)
                ws.write(linha, 2, getattr(pm, 're', '') or '', fmt_center)
                ws.write(linha, 3, getattr(pm, 'nome_guerra', None) or getattr(pm, 'nome', '') or '', fmt_left)
                ws.write(linha, 4, getattr(pm, 'viatura', '') or '', fmt_center)
                ws.write(linha, 5, get_habilitacoes_texto(pm) or '', fmt_center)
                ws.write(linha, 6, get_sat_texto(pm) or '', fmt_center)
                ws.write(linha, 7, getattr(pm, 'funcao', '') or '', fmt_center)
                ws.write(linha, 8, getattr(pm, 'secao_em', '') or '', fmt_center)
                ws.write(linha, 9, vis, estilo)

                linha += 1

        linha += 2
        ws.merge_range(f'D{linha}:H{linha}', '_' * 30, fmt_center)
        ws.merge_range(f'D{linha+1}:H{linha+1}', 'ENCARREGADO DE ESCALA', fmt_center)

        workbook.close()
        output.seek(0)

        return send_file(
            output,
            download_name=f"Escala_Diaria_{cia}_{versao_txt}.xlsx",
            as_attachment=True
        )

    except Exception as e:
        flash(f'Erro: {e}', 'error')
        return redirect(url_for('index'))

# --- ROTA DE EXPORTAR FREQUÊNCIA (LAYOUT CONTROLADOR FINANCEIRO) ---
@app.route('/exportar_excel')
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
            if d <= u_dia:
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

                if d <= u_dia:
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

@app.route('/gerar_relatorio_efetivo', methods=['POST'])
@login_required
def gerar_relatorio_efetivo():
    # Rota corrigida para exibir o relatório (template)
    cia = request.form.get('cia_relatorio')
    if not cia: return redirect(url_for('index'))
    
    # Filtros via form para cada coluna (dropdown por coluna)
    filtro_posto = request.form.get('filtro_posto', '').strip()
    filtro_re = request.form.get('filtro_re', '').strip()
    filtro_nome_guerra = request.form.get('filtro_nome_guerra', '').strip()
    filtro_cia = request.form.get('filtro_cia', '').strip()
    filtro_funcao = request.form.get('filtro_funcao', '').strip()
    filtro_situacao = request.form.get('filtro_situacao', '').strip()
    filtro_afastamento = request.form.get('filtro_afastamento', '').strip()
    filtro_habilitacoes = request.form.get('filtro_habilitacoes', '').strip()
    filtro_sat = request.form.get('filtro_sat', '').strip()
    
    if cia == 'TODAS':
        pms_raw = Policial.query.filter_by(ativo=True).all()
        titulo = "TODAS AS CIAS"
    else:
        pms_raw = Policial.query.filter_by(cia=cia, ativo=True).all()
        titulo = cia.upper()
    
    # Computar hab_texto e sat_texto para TODOS os PMs raw (antes de filtrar)
    for pm in pms_raw:
        # Para Habilitações e SAT: use funções se existirem, senão campos diretos
        pm.hab_texto = (get_habilitacoes_texto(pm) if callable(get_habilitacoes_texto) else (getattr(pm, 'habilitacoes', '') or '')) or ''
        pm.sat_texto = (get_sat_texto(pm) if callable(get_sat_texto) else (getattr(pm, 'sat', '') or '')) or ''
    
    # Aplicar filtros em todas as colunas
    pms_filtrados = []
    for pm in pms_raw:
        nome_guerra = (pm.nome_guerra or pm.nome or '').lower()
        funcao = (pm.funcao or '').lower()
        afastamento = (pm.restricao_tipo or '').lower()
        hab_texto_lower = pm.hab_texto.lower()
        sat_texto_lower = pm.sat_texto.lower()
        
        if (not filtro_posto or pm.posto == filtro_posto) and \
           (not filtro_re or filtro_re in (pm.re or '')) and \
           (not filtro_nome_guerra or filtro_nome_guerra.lower() in nome_guerra) and \
           (not filtro_cia or pm.cia == filtro_cia) and \
           (not filtro_funcao or filtro_funcao.lower() in funcao) and \
           (not filtro_situacao or pm.status_saude == filtro_situacao) and \
           (not filtro_afastamento or filtro_afastamento.lower() in afastamento) and \
           (not filtro_habilitacoes or filtro_habilitacoes.lower() in hab_texto_lower) and \
           (not filtro_sat or filtro_sat.lower() in sat_texto_lower):
            pms_filtrados.append(pm)
    
    if cia == 'EM': 
        pms = ordenar_em_agrupado(pms_filtrados)
    else: 
        pms = ordenar_por_antiguidade(pms_filtrados)
    
    # Opções para dropdowns (únicos de cada coluna, sem vazios)
    opcoes_posto = sorted([p for p in set(pm.posto for pm in pms_raw if pm.posto and pm.posto.strip())])
    opcoes_re = sorted([r for r in set(pm.re for pm in pms_raw if pm.re and pm.re.strip())])
    opcoes_nome_guerra = sorted([n for n in set((pm.nome_guerra or pm.nome) for pm in pms_raw if (pm.nome_guerra or pm.nome) and (pm.nome_guerra or pm.nome).strip())])
    opcoes_cia = sorted([c for c in set(pm.cia for pm in pms_raw if pm.cia and pm.cia.strip())])
    opcoes_funcao = sorted([f for f in set(pm.funcao for pm in pms_raw if pm.funcao and pm.funcao.strip())])
    opcoes_situacao = sorted([s for s in set(pm.status_saude for pm in pms_raw if pm.status_saude and pm.status_saude.strip())])
    opcoes_afastamento = sorted([a for a in set(pm.restricao_tipo for pm in pms_raw if pm.restricao_tipo and pm.restricao_tipo.strip())])
    opcoes_habilitacoes = sorted([h for h in set(pm.hab_texto for pm in pms_raw if pm.hab_texto and pm.hab_texto.strip())])
    opcoes_sat = sorted([s for s in set(pm.sat_texto for pm in pms_raw if pm.sat_texto and pm.sat_texto.strip())])
    
    return render_template(
        'relatorio_efetivo.html', 
        pms=pms, 
        cia=titulo, 
        hoje=datetime.now(),
        # Opções para dropdowns
        opcoes_posto=opcoes_posto, opcoes_re=opcoes_re, opcoes_nome_guerra=opcoes_nome_guerra,
        opcoes_cia=opcoes_cia, opcoes_funcao=opcoes_funcao, opcoes_situacao=opcoes_situacao,
        opcoes_afastamento=opcoes_afastamento, opcoes_habilitacoes=opcoes_habilitacoes, opcoes_sat=opcoes_sat,
        # Filtros atuais para manter estado
        filtro_posto=filtro_posto, filtro_re=filtro_re, filtro_nome_guerra=filtro_nome_guerra,
        filtro_cia=filtro_cia, filtro_funcao=filtro_funcao, filtro_situacao=filtro_situacao,
        filtro_afastamento=filtro_afastamento, filtro_habilitacoes=filtro_habilitacoes, filtro_sat=filtro_sat
    )

# ==============================================================================
# GESTÃO DE USUÁRIOS (ADMIN)
# ==============================================================================

@app.route('/usuarios')
@login_required
def listar_usuarios():
    if getattr(current_user, 'perfil', 'usuario') != 'admin': return redirect(url_for('index'))
    return render_template('usuarios.html', usuarios=Usuario.query.all())

@app.route('/novo_usuario', methods=['POST'])
@login_required
def novo_usuario():
    if getattr(current_user, 'perfil', 'usuario') != 'admin': return redirect(url_for('index'))
    re = request.form['re']
    senha = request.form['senha']
    novo = Usuario(username=re, re=re, nome=request.form['nome'],
                   password=generate_password_hash(senha), perfil=request.form['perfil'])
    db.session.add(novo)
    db.session.commit()
    flash('Usuário criado com sucesso!', 'success')
    return redirect(url_for('listar_usuarios'))

@app.route('/excluir_usuario/<int:id_user>')
@login_required
def excluir_usuario(id_user):
    if getattr(current_user, 'perfil', 'usuario') != 'admin': return redirect(url_for('index'))
    if id_user == current_user.id:
        flash('Você não pode excluir a si mesmo.', 'warning')
        return redirect(url_for('listar_usuarios'))
        
    user = db.session.get(Usuario, id_user)
    if user:
        db.session.delete(user)
        db.session.commit()
        flash('Usuário excluído.', 'success')
    return redirect(url_for('listar_usuarios'))

@app.route('/alterar_minha_senha', methods=['POST'])
@login_required
def alterar_minha_senha():
    if not check_password_hash(current_user.password, request.form.get('senha_atual')):
        flash('Senha atual incorreta.', 'error')
        return redirect(request.referrer)
    current_user.password = generate_password_hash(request.form.get('nova_senha'))
    db.session.commit()
    flash('Sua senha foi alterada.', 'success')
    return redirect(request.referrer)

@app.route('/admin_reset_senha', methods=['POST'])
@login_required
def admin_reset_senha():
    if getattr(current_user, 'perfil', 'usuario') != 'admin': return redirect(url_for('index'))
    user = db.session.get(Usuario, int(request.form.get('user_id')))
    if user:
        user.password = generate_password_hash(request.form.get('nova_senha'))
        db.session.commit()
        flash(f'Senha do usuário {user.re} resetada.', 'success')
    return redirect(url_for('listar_usuarios'))


# ==============================================================================
# INICIALIZAÇÃO
# ==============================================================================

if __name__ == '__main__':
    with app.app_context():
        db.create_all()
        # Garante admin
        if not Usuario.query.filter_by(username='admin').first():
            db.session.add(Usuario(username='admin', re='admin', nome='Administrador', 
                                 password=generate_password_hash('admin'), perfil='admin'))
            db.session.commit()
            print("✅ Usuário admin criado.")
            
    # MODO DESENVOLVIMENTO (Apenas na sua máquina local)
    # app.run(debug=True) 

    # MODO PRODUÇÃO / SERVIDOR DE REDE
    from waitress import serve
    print("🚀 Servidor da P1 rodando! Aguardando conexões na porta 80...")
    serve(app, host='0.0.0.0', port=80)
    
'''if __name__ == '__main__':
    with app.app_context():
        db.create_all()
        # Garante admin
        if not Usuario.query.filter_by(username='admin').first():
            db.session.add(Usuario(username='admin', re='admin', nome='Administrador', 
                                 password=generate_password_hash('admin'), perfil='admin'))
            db.session.commit()
            print("✅ Usuário admin criado.")
            
    app.run(debug=True)'''