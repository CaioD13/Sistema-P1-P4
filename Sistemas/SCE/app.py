import os
import calendar
import xlsxwriter
from datetime import datetime, timedelta
from io import BytesIO
from flask import Flask, render_template, request, redirect, url_for, flash, send_file
from werkzeug.security import generate_password_hash, check_password_hash
from flask_login import LoginManager, login_user, login_required, logout_user, current_user, UserMixin
from sqlalchemy import desc
from flask_sqlalchemy import SQLAlchemy
from datetime import datetime, date
from datetime import timedelta

# ==============================================================================
# CONFIGURAÇÕES INICIAIS
# ==============================================================================

basedir = os.path.abspath(os.path.dirname(__file__))
db_path = os.path.join(basedir, 'frequencia.db')

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
    inicio_pares = db.Column(db.Boolean, default=True) # True = Dias Pares
    data_base_escala = db.Column(db.Date, nullable=True)
    funcao = db.Column(db.String(100))
    equipe = db.Column(db.String(50))
    viatura = db.Column(db.String(50))
    ativo = db.Column(db.Boolean, default=True)
        
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

ORDEM_SECOES_EM = {
    'Comando': 1, 'Subcomando': 2, 'Coord Op': 3,
    'P1': 4, 'P1/P5': 4, 'P2': 5, 'P3': 6, 'P4': 7, 'P5': 8,
    'SPJMD': 9, 'SJD': 9, 'SAD': 10, 'Tesouraria': 11, 'Motomec': 12, 'CFP': 13
}

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

def calcular_escala_dia(pm, data_obj):
    # Padronização da data
    if isinstance(data_obj, datetime): data = data_obj.date()
    else: data = data_obj

    # 1. PRIORIDADE MÁXIMA: Alterações manuais (JM, M1, M2, etc.)
    alteracao = AlteracaoEscala.query.filter_by(policial_id=pm.id, data=data).first()
    if alteracao: return alteracao.valor

    # 2. CONSULTA AO PASSADO: Verifica se o mês da data sendo calculada está no histórico
    historico = HistoricoEscala.query.filter_by(
        policial_id=pm.id, 
        mes=data.month, 
        ano=data.year
    ).first()
    
    # Se houver histórico para esse mês, usa a escala antiga. Senão, usa a atual do PM.
    tipo_vigente = historico.tipo_escala_antigo if historico else pm.tipo_escala

    # 3. CÁLCULO DAS ESCALAS OPERACIONAIS (12x36)
    escalas_12x36 = ['S/3', 'S/4', 'S/5', 'S/6', '12x36']
    
    if tipo_vigente in escalas_12x36:
        # Usa a data base (âncora) salva no cadastro
        if pm.data_base_escala:
            base = pm.data_base_escala
        else:
            # Segurança: Se não tiver âncora, assume janeiro como base
            base = date(data.year, 1, 2 if pm.inicio_pares else 1)

        # Lógica de Delta: Dias passados desde a âncora
        delta = (data - base).days
        trabalha_hoje = (delta % 2 == 0)
        
        return tipo_vigente if trabalha_hoje else 'F'
    
    # 4. CÁLCULO DAS ESCALAS ADMINISTRATIVAS
    else:
        # Folga FDS e Feriados
        if data.weekday() >= 5 or e_feriado(data): return 'F'
        return tipo_vigente
    
def traduzir_para_frequencia(sigla):
    """
    Traduz a sigla da escala para o código de alimentação baseado na DURAÇÃO DA JORNADA.
    
    Regras de Carga Horária:
    "0": < 8h (Meio Expediente)
    "1": >= 8h e < 12h (Expediente Administrativo)
    "2": >= 12h e < 18h (Plantões 12x36, Diurno/Noturno)
    "3": >= 18h (Plantões 24h, Dobras, Sobreaviso Longo)
    
    Códigos Especiais:
    "F": Férias
    "LP": Licença-Prêmio
    "FS": Falta ao Serviço
    "A": Demais Afastamentos
    """
    if not sigla: return '?'
    sigla = sigla.strip().upper()
    
    # === CÓDIGOS ESPECIAIS (LETRAS) ===
    if sigla in ['FN', 'FÉRIAS', 'FERIAS']: return 'FN'
    
    # Licença-Prêmio
    if sigla in ['LP', 'LICENÇA PRÊMIO', 'LICENÇA PREMIO']: return 'LP'
    
    # Falta ao Serviço
    if sigla in ['FALTA', 'FALTOU', 'F.S.', 'FS']: return 'FS'
    
    # Demais Afastamentos -> A
    lista_afastamentos = [
        'A', 'FOLGA', 'F', 'F*', 'JM', 'FM', 'DS', 'DR', 'FR', 'PF', 
        'CV', 'LTS', 'LT', 'LG', 'PT', 'LA', 'NP', 'LICENÇA', 'AFASTADO',
        'DISPENSA', 'LICENCA', 'LSV'
    ]
    # Se for F (Folga comum), não gera diária, mas no sistema financeiro pode ser A ou vazio
    if sigla == 'F': return 'A' 
    if sigla in lista_afastamentos: return 'A'

    # === CÁLCULO POR DURAÇÃO (NÚMEROS) ===
    
    # "0": < 8h (Meio Período 6h)
    if sigla in ['M1', 'M2', 'MEIO', '6H']: 
        return '0'

    # "1": >= 8h e < 12h (Expediente Administrativo 8h-10h)
    if sigla in ['S', 'S/1', 'S/2', 'S/8' 'ADM', 'EXP', '8H', '9H', '10H', 'EAP']: 
        return '1'

    # "2": >= 12h e < 18h (Plantões 12x36 Diurno/Noturno)
    # S/3, S/4, S/5, S/6 e S7 são escalas de 12 horas
    if sigla in ['S/3', 'S/4', 'S/5', 'S/6', 'S/7','12X36', 'PLANTÃO', '12H']: 
        return '2'

    # "3": >= 18h (Plantões 24h, Dobras, Sobreaviso Longo)
    # SR, EE, geralmente são escalas longas ou especiais
    if sigla in ['SR', 'PPJMD', '24H', '18H', 'DOBRA']: 
        return '3'
    
    # Se não reconhecer, retorna ?
    if sigla in ['SS', 'S*']:
        return '?'

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
    siglas_afastamento = ['FN', 'FS', 'LTS', 'CV', 'LG', 'PT', 'LA', 'NP', 'LT', 'LP', 'LSV', 'AD'] 
    
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
    """
    funcao = (pm.funcao or "").upper()
    tipo = (pm.tipo_escala or "").upper()
    
    # 1. BLOCO CFP
    if 'CFP' in funcao or 'AUX' in funcao or 'MOTORISTA' in funcao:
        if tipo == 'S/6' or 'NOTURNO' in funcao:
            return "CFP Noturno"
        return "CFP Diurno"

    # 2. BLOCO GUARDA / S.I.R.
    if 'S.I.R' in funcao or 'SIR' in funcao:
        return "Guarda do Quartel Diurno"
        
    if 'GUARDA' in funcao or 'SENTINELA' in funcao or 'PERMANÊNCIA' in funcao:
        if 'NOTURNO' in funcao or tipo == 'S/2':
            return "Guarda do Quartel Noturno"
        return "Guarda do Quartel Diurno"

    # 3. BLOCO MANUTENÇÃO / MOTOMEC
    if 'MOTOMEC' in funcao or 'MANUTEN' in funcao or 'VTR' in funcao:
        return "Manutenção/Motomec"

    # 4. BLOCO ADMINISTRAÇÃO
    termos_adm = ['P1', 'P2', 'P3', 'P4', 'P5', 'TESOURARIA', 'SECRETARIA', 'JUSTIÇA', 'CMT', 'SUB', 'CMD']
    if any(t in funcao for t in termos_adm):
        return "Administração"
    
    # 5. FALLBACK PELA SIGLA
    if tipo in ['S', 'S/1', 'S/2', 'M1', 'M2', 'ADM']:
        return "Administração"
    if tipo == 'S/3': return "Guarda do Quartel Diurno"
    if tipo == 'S/4': return "Manutenção/Motomec"
    if tipo == 'S/5': return "CFP Diurno"
    if tipo == 'S/6': return "CFP Noturno"
    
    return "Outros / A Definir"

INFO_BLOCOS = {
    "Administração": "das 08h00 às 18h00",
    "Manutenção/Motomec": "das 07h00 às 19h00",
    "Guarda do Quartel Diurno": "das 06h00 às 18h00",
    "Guarda do Quartel Noturno": "das 18h00 às 06h00",
    "CFP Diurno": "das 04h45 às 17h00",
    "CFP Noturno": "das 16h45 às 05h00",
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

# ==============================================================================
# ROTAS E VIEWS
# ==============================================================================

@app.route('/')
@login_required
def index():
    verificar_validade_restricoes()
    pms_raw = Policial.query.filter_by(ativo=True).all()
    
    hoje = datetime.now().date()
    
    # === AQUI ESTAVA O ERRO: Adicionadas as siglas FN, FS, CV e demais ===
    siglas_afastamento = ['F', 'FN', 'FS', 'CV', 'LP', 'LSV', 'LTS', 'AD', 'LG', 'LA', 'PT', 'NP', 'EAP', 'LT']
    
    # Nomes bonitos para aparecer na tela inicial em vez da sigla pura
    nomes_afastamento = {
        'F': 'Férias',
        'FN': 'Férias (FN)', 
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
    return render_template('index.html', pms=pms_ordenados, user=current_user, e_admin=e_admin, hoje=hoje)

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
@login_required
def salvar_pm():
    try:
        id_pm = request.form.get('id_pm')
        
        # Coletando dados do formulário
        re = request.form.get('re')
        posto = request.form.get('posto')
        nome = request.form.get('nome')
        nome_guerra = request.form.get('nome_guerra')
        cia = request.form.get('cia')
        secao_em = request.form.get('secao_em') if cia == 'EM' else ''
        funcao = request.form.get('funcao')
        equipe = request.form.get('equipe')
        viatura = request.form.get('viatura')
        tipo_escala = request.form.get('tipo_escala')
        inicio_pares_req = True if request.form.get('inicio_pares') else False
        status_saude = request.form.get('status_saude')
        restricao_tipo = request.form.get('restricao_tipo') if status_saude == 'Apto Com Restrições' else ''
        
        data_fim_str = request.form.get('data_fim_restricao')
        data_fim_restricao = None
        if status_saude == 'Apto Com Restrições' and data_fim_str:
            data_fim_restricao = datetime.strptime(data_fim_str, '%Y-%m-%d').date()
        
        hoje = datetime.now()

        if id_pm:
            pm = Policial.query.get(id_pm)
            if not pm:
                flash('Policial não encontrado!', 'danger')
                return redirect(url_for('index'))
            
            # === LÓGICA CORRIGIDA: BLOQUEIO DO PASSADO ===
            # Se o tipo de escala mudou, vamos "congelar" todos os meses passados
            # deste ano com a escala que o PM tinha até agora.
            if pm.tipo_escala and pm.tipo_escala != tipo_escala:
                for m in range(1, hoje.month):
                    # Verifica se já não existe histórico para este mês antes de criar
                    existe = HistoricoEscala.query.filter_by(policial_id=pm.id, mes=m, ano=hoje.year).first()
                    if not existe:
                        h = HistoricoEscala(
                            policial_id=pm.id,
                            mes=m,
                            ano=hoje.year,
                            tipo_escala_antigo=pm.tipo_escala # Guarda a escala velha nos meses passados
                        )
                        db.session.add(h)
            
            mensagem = 'Dados do policial atualizados com sucesso!'
        else:
            pm = Policial()
            db.session.add(pm)
            pm.ativo = True 
            mensagem = 'Novo policial cadastrado com sucesso!'
        
        # Atualiza os dados principais
        pm.re = re
        pm.posto = posto
        pm.nome = nome
        pm.nome_guerra = nome_guerra
        pm.cia = cia
        pm.secao_em = secao_em
        pm.funcao = funcao
        pm.equipe = equipe
        pm.viatura = viatura
        pm.status_saude = status_saude
        pm.restricao_tipo = restricao_tipo
        pm.data_fim_restricao = data_fim_restricao
        pm.tipo_escala = tipo_escala
        pm.inicio_pares = inicio_pares_req
        
        # Atualiza a âncora para o mês corrente (conforme seu pedido)
        if inicio_pares_req:
            pm.data_base_escala = date(hoje.year, hoje.month, 2)
        else:
            pm.data_base_escala = date(hoje.year, hoje.month, 1)

        db.session.commit()
        flash(mensagem, 'success')
        
    except Exception as e:
        db.session.rollback()
        flash(f'Erro ao salvar policial: {e}', 'danger')
        
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
    pm = db.session.get(Policial, id_pm)
    hoje = datetime.now()
    mes = request.args.get('mes', hoje.month, type=int)
    ano = request.args.get('ano', hoje.year, type=int)
    
    p_dia, u_dia = calendar.monthrange(ano, mes)
    dias_vazios = (p_dia + 1) % 7
    
    escala_vis = []; total_trabalhado = 0
    siglas_folga = ['F', 'A', 'F*', 'JM', 'FM', 'DS', 'DR', 'FR', 'PF', 'CV', 'LTS', 'EAP', 'LT', 'LG', 'PT', 'LA', 'NP', 'FN', 'FS']

    for d in range(1, u_dia + 1):
        dt = datetime(ano, mes, d)
        sigla = calcular_escala_dia(pm, dt)
        if sigla not in siglas_folga: total_trabalhado += 1
        
        cor = 'bg-light'
        if sigla in siglas_folga: cor = 'bg-white text-muted border-light'
        elif sigla in ['S', 'S/1', 'S/2', 'S/8' 'M1', 'M2']: cor = 'bg-success text-white'
        elif sigla in ['S/3', 'S/4', 'S/5', 'S/6', 'S/7']: cor = 'bg-primary text-white'
        elif sigla in ['LTS', 'CV', 'FN', 'FS']: cor = 'bg-warning text-dark'
        elif sigla in ['SS', 'SR', 'EE']: cor = 'bg-info text-dark'
        elif sigla in ['LP']: cor = 'bg-info text-white', 'Licença Prêmio'
        elif sigla in ['LSV']: cor = 'bg-secondary text-white', 'Licença S/ Venc.'
        elif sigla in ['AD']: cor = 'bg-primary text-white', 'Adido (Outra OPM)'
        
        escala_vis.append({'dia': d, 'valor': sigla, 'classe': cor, 'data_iso': dt.strftime('%Y-%m-%d')})

    nav = {'ant_m': mes-1 if mes>1 else 12, 'ant_a': ano if mes>1 else ano-1, 'prox_m': mes+1 if mes<12 else 1, 'prox_a': ano if mes<12 else ano+1}
    
    # Busca se este mês específico já foi validado para este PM
    validacao = ValidacaoMensal.query.filter_by(policial_id=id_pm, mes=mes, ano=ano).first()
    
    return render_template('editar.html', pm=pm, mes=mes, ano=ano, escala=escala_vis, dias_vazios=dias_vazios, totais={'trabalhado': total_trabalhado}, nav=nav, validacao=validacao, current_user=current_user)

@app.route('/salvar_alteracao', methods=['POST'])
@login_required
def salvar_alteracao():
    try:
        id_pm = int(request.form.get('id_pm'))
        inicio = datetime.strptime(request.form.get('data_inicio'), '%Y-%m-%d').date()
        fim_str = request.form.get('data_fim')
        fim = datetime.strptime(fim_str, '%Y-%m-%d').date() if fim_str else inicio
        valor = request.form.get('valor')
        
        pm = db.session.get(Policial, id_pm)
        if not pm:
            flash('Erro: Policial não encontrado.', 'error')
            return redirect(request.referrer)
        
        # --- TRAVA DE VALIDAÇÃO (CHEFE P1) ---
        mes_inicio = inicio.month
        ano_inicio = inicio.year
        validacao = ValidacaoMensal.query.filter_by(policial_id=id_pm, mes=mes_inicio, ano=ano_inicio).first()
        
        e_admin = getattr(current_user, 'perfil', 'usuario') == 'admin'
        
        if validacao and not e_admin:
            flash(f'⚠️ Bloqueado: A escala deste mês já foi VALIDADA por {validacao.admin_info}. Apenas o administrador pode alterá-la.', 'danger')
            return redirect(request.referrer)

        # --- INÍCIO DAS VALIDAÇÕES INTELIGENTES ---
        
        escalas_operacionais = ['S/3', 'S/4', 'S/5', 'S/6', '12x36']
        e_administrativo = pm.tipo_escala not in escalas_operacionais

        # REGRA ZERO: Operacional/12x36 NÃO PODE ter JM, M1 ou M2
        if not e_administrativo and valor in ['JM', 'M1', 'M2']:
            flash(f'⚠️ Bloqueado: Policiais em escala Operacional ({pm.tipo_escala}) não têm direito a {valor}.', 'danger')
            return redirect(request.referrer)

        # Validação para Administrativos (Regras Semanais para JM/M1/M2)
        if e_administrativo and valor in ['JM', 'M1', 'M2'] and (fim - inicio).days == 0:
            
            # Encontrar a Segunda-feira da semana da data solicitada
            dia_da_semana = inicio.weekday() 
            segunda = inicio - timedelta(days=dia_da_semana)
            
            qtd_meio_exp = 0
            has_jm_na_semana = False
            contagem_folgas_outros_dias = 0
            contagem_trabalho = 0
            
            todas_folgas = ['F', 'A', 'JM', 'FM', 'DS', 'DR', 'FR', 'PF', 'CV', 'LTS', 'EAP', 'LT', 'LG', 'PT', 'LA', 'NP', 'FN', 'FS', 'LP', 'LSV', 'AD']
            
            # Analisamos a SEMANA INTEIRA (7 dias) para validar as travas
            for d in range(7):
                dia_analise = segunda + timedelta(days=d)
                
                # Se for o dia que estamos editando agora, usamos o 'valor'. 
                # Caso contrário, usamos o que já está salvo no sistema (calcular_escala_dia)
                sigla_dia = valor if dia_analise == inicio else calcular_escala_dia(pm, dia_analise)
                
                # 1. Contagem de Meio Expediente (M1 ou M2)
                if sigla_dia in ['M1', 'M2']:
                    qtd_meio_exp += 1
                
                # 2. Verificação de existência de JM
                if sigla_dia == 'JM':
                    has_jm_na_semana = True
                
                # 3. Contagens auxiliares (para regras de dias trabalhados e folgas em dias úteis)
                if dia_analise.weekday() < 5: # Segunda a Sexta
                    if sigla_dia in todas_folgas and dia_analise != inicio:
                        contagem_folgas_outros_dias += 1
                    
                    # Se for código de trabalho normal (1, 2 ou 3)
                    if traduzir_para_frequencia(sigla_dia) in ['1', '2', '3']:
                        contagem_trabalho += 1

            # --- APLICAÇÃO DAS REGRAS REQUISITADAS ---
            
            # REGRA 1: JM e M1/M2 são EXCLUSIVOS (Não pode ter os dois na mesma semana)
            if has_jm_na_semana and qtd_meio_exp > 0:
                flash('⚠️ Bloqueado: Não é permitido combinar Junção de Meios (JM) com Meio Expediente (M1/M2) na mesma semana.', 'danger')
                return redirect(request.referrer)

            # REGRA 2: Limite máximo de 2 Meios Expedientes por semana
            if qtd_meio_exp > 2:
                flash('⚠️ Bloqueado: Limite máximo de 2 Meios Expedientes (M1/M2) por semana atingido.', 'danger')
                return redirect(request.referrer)

            # REGRA 3: Validações específicas para quando o valor atual é JM
            if valor == 'JM':
                # Verifica se já existe outro JM (além do que estamos tentando salvar)
                # Fazemos uma busca direta no banco/escala para os outros dias
                outros_jm = [segunda + timedelta(days=d) for d in range(7) if (segunda + timedelta(days=d)) != inicio and calcular_escala_dia(pm, segunda + timedelta(days=d)) == 'JM']
                if outros_jm:
                    flash('⚠️ Bloqueado: Não é permitido lançar mais de uma Junção de Meios (JM) na mesma semana.', 'danger')
                    return redirect(request.referrer)
                
                if contagem_folgas_outros_dias >= 2: 
                    flash('⚠️ Bloqueado: O PM já possui 2 ou mais dias de folga/afastamento na semana (Seg-Sex). Junção de Meios não permitida.', 'danger')
                    return redirect(request.referrer)

            # REGRA 4: Validação de trabalho mínimo para M1/M2
            elif valor in ['M1', 'M2']:
                if contagem_trabalho < 2:
                    flash('⚠️ Bloqueado: Para fazer jus a Meio Expediente, o PM deve trabalhar pelo menos 2 dias inteiros na semana.', 'danger')
                    return redirect(request.referrer)

        # --- FIM DAS VALIDAÇÕES ---

        # Execução da gravação no banco
        for i in range((fim - inicio).days + 1):
            dia = inicio + timedelta(days=i)
            alt = AlteracaoEscala.query.filter_by(policial_id=id_pm, data=dia).first()
            if valor == 'LIMPAR':
                if alt: 
                    db.session.delete(alt)
            else:
                if not alt: 
                    db.session.add(AlteracaoEscala(policial_id=id_pm, data=dia, valor=valor))
                else: 
                    alt.valor = valor
                    
        db.session.commit()
        flash('✅ Escala atualizada com sucesso!', 'success')
        
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

@app.route('/visualizar_escala_semanal', methods=['POST'])
@login_required
def visualizar_escala_semanal():
    try:
        cia = request.form.get('cia_filtro')
        inicio_str = request.form.get('data_inicio_semana')
        tipo_filtro = request.form.get('tipo_filtro', 'TODOS')
        
        if not cia or not inicio_str: 
            return redirect(url_for('index'))
        
        inicio = datetime.strptime(inicio_str, '%Y-%m-%d').date()
        
        # Gera os dias da semana para o cabeçalho
        dias_semana = []
        nomes_dias = ['SEG', 'TER', 'QUA', 'QUI', 'SEX', 'SAB', 'DOM']
        for i in range(7):
            d = inicio + timedelta(days=i)
            dias_semana.append({'data_str': f"{d.day:02d}/{d.month:02d}", 'dia_semana': nomes_dias[i]})
            
        # Busca e filtra os PMs (Muda a query dependendo da CIA)
        pms_raw = Policial.query.filter_by(ativo=True)
        if cia != 'TODAS' and cia != 'EM': 
            pms_raw = pms_raw.filter_by(cia=cia)
        pms_raw = pms_raw.all()
        
        if cia == 'EM': 
            pms_raw = [p for p in pms_raw if p.cia == 'EM']
            
        todos_pms = filtrar_por_tipo_hierarquico(pms_raw, tipo_filtro)
        
        if cia == 'EM': 
            todos_pms = ordenar_em_agrupado(todos_pms)
        else: 
            todos_pms = ordenar_cia_agrupado(todos_pms)

        # Agrupa nos blocos de horário
        mapa_pms = {k: [] for k in ORDEM_BLOCOS}
        for pm in todos_pms:
            nome_bloco = identificar_bloco_horario(pm)
            if nome_bloco in mapa_pms: 
                mapa_pms[nome_bloco].append(pm)
            else: 
                mapa_pms["Outros / A Definir"].append(pm)

        # Monta a estrutura para enviar ao HTML
        blocos_view = []
        siglas_afastamento = ['LTS', 'CV', 'FN', 'FS', 'LG', 'LA', 'PT', 'NP', 'LP', 'EAP', 'LT', 'JM', 'FM', 'DS', 'DR', 'FR', 'PF', 'F*']
        
        for nome_bloco in ORDEM_BLOCOS:
            lista_pms = mapa_pms[nome_bloco]
            if not lista_pms: continue
            
            pms_view = []
            for pm in lista_pms:
                escala_pm = []
                for i in range(7):
                    data_atual = inicio + timedelta(days=i)
                    sigla = calcular_escala_dia(pm, data_atual)
                    vis = 'F' if sigla in ['A', 'Folga', 'FOLGA'] else sigla
                    is_afastamento = vis in siglas_afastamento
                    escala_pm.append({'valor': vis, 'is_afastamento': is_afastamento})
                
                pms_view.append({
                    'equipe': getattr(pm, 'equipe', '') or '',
                    'posto': pm.posto,
                    're': pm.re,
                    'nome': pm.nome_guerra or pm.nome,
                    'viatura': getattr(pm, 'viatura', '') or '',
                    'funcao': pm.funcao or '',
                    'secao_em': getattr(pm, 'secao_em', '') or '',
                    'dias': escala_pm
                })
            
            blocos_view.append({
                'nome': nome_bloco,
                'horario': INFO_BLOCOS.get(nome_bloco, ""),
                'pms': pms_view
            })
            
        return render_template('visualizar_escala.html', 
                               cia=cia, 
                               data_inicio=inicio, 
                               data_fim=inicio + timedelta(days=6),
                               dias_semana=dias_semana, 
                               blocos=blocos_view, 
                               tipo_filtro=tipo_filtro)
                               
    except Exception as e:
        flash(f'Erro na pré-visualização: {e}', 'error')
        return redirect(url_for('index'))

@app.route('/gerar_escala_semanal', methods=['POST'])
@login_required
def gerar_escala_semanal():
    try:
        cia = request.form.get('cia_filtro')
        inicio_str = request.form.get('data_inicio_semana')
        tipo_filtro = request.form.get('tipo_filtro', 'TODOS')
        
        if not cia or not inicio_str: return redirect(url_for('index'))
        
        inicio = datetime.strptime(inicio_str, '%Y-%m-%d').date()
        fim = inicio + timedelta(days=6)
        versao_txt = get_proxima_versao(cia, inicio, tipo_filtro)
        
        tipo_txt = "OFICIAIS" if tipo_filtro == 'OFICIAIS' else "PRAÇAS" if tipo_filtro == 'PRACAS' else "POLICIAIS MILITARES"
        numerador = f"{get_semana_mes(inicio)}/{str(inicio.year)[2:]}"
        periodo_txt = f"DE {formatar_data_militar(inicio)} A {formatar_data_militar(fim)}"
        
        # Configuração do Arquivo Excel
        output = BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        ws = workbook.add_worksheet("ESCALA")
        ws.set_paper(9) # A4
        ws.set_landscape() # Paisagem
        ws.set_margins(0.5, 0.5, 0.5, 0.5)
        
        # Estilos
        fmt_cabecalho = workbook.add_format({'bold': True, 'font_name': 'Arial', 'font_size': 10, 'align': 'center', 'valign': 'vcenter'})
        fmt_numerador = workbook.add_format({'bold': True, 'font_name': 'Arial', 'font_size': 10, 'align': 'center', 'valign': 'vcenter', 'underline': True})
        fmt_bloco = workbook.add_format({'bold': True, 'font_size': 9, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#595959', 'font_color': 'white', 'border': 1})
        fmt_th = workbook.add_format({'bold': True, 'font_size': 8, 'align': 'center', 'border': 1, 'bg_color': '#EFEFEF', 'text_wrap': True})
        fmt_dia = workbook.add_format({'bold': True, 'font_size': 8, 'align': 'center', 'border': 1, 'bg_color': '#D9D9D9'})
        fmt_center = workbook.add_format({'font_size': 8, 'align': 'center', 'border': 1})
        fmt_left = workbook.add_format({'font_size': 8, 'align': 'left', 'border': 1})
        fmt_afastamento = workbook.add_format({'font_size': 8, 'align': 'center', 'border': 1, 'bg_color': '#FFFFCC'}) # Amarelo Claro
        fmt_padrao = fmt_center

        # Colunas
        ws.set_column('A:A', 6); ws.set_column('B:B', 10); ws.set_column('C:C', 8)
        ws.set_column('D:D', 25); ws.set_column('E:E', 8); ws.set_column('F:G', 5)
        ws.set_column('H:H', 12); ws.set_column('I:I', 8); ws.set_column('J:P', 6)

        # Cabeçalho Oficial
        img_path = os.path.join(basedir, 'brasao.png')
        if os.path.exists(img_path):
            ws.insert_image('A1', img_path, {'x_scale': 0.22, 'y_scale': 0.22, 'x_offset': 15, 'y_offset': 5})
        
        ws.merge_range('C1:P1', "SECRETARIA DA SEGURANÇA PÚBLICA", fmt_cabecalho)
        ws.merge_range('C2:P2', "POLÍCIA MILITAR DO ESTADO DE SÃO PAULO", fmt_cabecalho)
        ws.merge_range('C3:P3', "COMANDO DE POLICIAMENTO DE ÁREA METROPOLITANA DEZ", fmt_cabecalho)
        ws.merge_range('C4:P4', "TRIGÉSIMO SÉTIMO BATALHÃO DE POLICIA MILITAR METROPOLITANO", fmt_cabecalho)
        texto_escala = f"ESCALA DE SERVIÇO DE {tipo_txt} Nº 37BPMM-{numerador} - {versao_txt}"
        ws.merge_range('A6:P6', texto_escala, fmt_numerador)
        ws.merge_range('A7:P7', periodo_txt, fmt_cabecalho)

        # Tabela e Dados
        linha = 8
        cols = ["EQ", "POSTO", "RE", "NOME", "VTR", "ENT", "SAI", "FUNÇÃO", "SEÇÃO"]
        for i, c in enumerate(cols): ws.write(linha, i, c, fmt_th)
        dias_semana = ['SEG', 'TER', 'QUA', 'QUI', 'SEX', 'SAB', 'DOM']
        for i in range(7):
            d = inicio + timedelta(days=i)
            ws.write(linha, 9+i, f"{dias_semana[i]}\n{d.day}/{d.month}", fmt_dia)
        linha += 1

        pms_raw = Policial.query.filter_by(cia=cia, ativo=True).all()
        todos_pms = filtrar_por_tipo_hierarquico(pms_raw, tipo_filtro)
        if cia == 'EM': todos_pms = ordenar_em_agrupado(todos_pms)
        else: todos_pms = ordenar_cia_agrupado(todos_pms)

        mapa_pms = {k: [] for k in ORDEM_BLOCOS}
        for pm in todos_pms:
            nome_bloco = identificar_bloco_horario(pm)
            if nome_bloco in mapa_pms: mapa_pms[nome_bloco].append(pm)
            else: mapa_pms["Outros / A Definir"].append(pm)

        for nome_bloco in ORDEM_BLOCOS:
            lista_pms = mapa_pms[nome_bloco]
            if not lista_pms: continue
            
            horario = INFO_BLOCOS.get(nome_bloco, "")
            texto_bloco = f"{nome_bloco}          {horario}"
            ws.merge_range(f'A{linha+1}:P{linha+1}', texto_bloco, fmt_bloco)
            linha += 1
            
            for pm in lista_pms:
                ws.write(linha, 0, getattr(pm, 'equipe', '') or '', fmt_center)
                ws.write(linha, 1, pm.posto, fmt_center)
                ws.write(linha, 2, pm.re, fmt_center)
                ws.write(linha, 3, pm.nome_guerra or pm.nome, fmt_left)
                ws.write(linha, 4, getattr(pm, 'viatura', '') or '', fmt_center)
                ws.write(linha, 5, "", fmt_center)
                ws.write(linha, 6, "", fmt_center)
                ws.write(linha, 7, pm.funcao or '', fmt_center)
                ws.write(linha, 8, getattr(pm, 'secao_em', '') or '', fmt_center)

                for i in range(7):
                    sigla = calcular_escala_dia(pm, inicio + timedelta(days=i))
                    vis = 'F' if sigla in ['A', 'Folga', 'FOLGA'] else sigla
                    estilo = fmt_padrao
                    siglas_afastamento = ['LTS', 'CV', 'FN', 'FS', 'LG', 'LA', 'PT', 'NP', 'LP', 'EAP', 'LT', 'JM', 'FM', 'DS', 'DR', 'FR', 'PF', 'F*']
                    if vis in siglas_afastamento: estilo = fmt_afastamento
                    ws.write(linha, 9+i, vis, estilo)
                linha += 1

        linha += 3
        ws.merge_range(f'D{linha}:H{linha}', '_' * 30, fmt_center)
        ws.merge_range(f'D{linha+1}:H{linha+1}', 'ENCARREGADO DE ESCALA', fmt_center)
        ws.merge_range(f'K{linha}:O{linha}', '_' * 30, fmt_center)
        ws.merge_range(f'K{linha+1}:O{linha+1}', f'CMT DA {cia}', fmt_center)

        workbook.close()
        output.seek(0)
        return send_file(output, download_name=f"Escala_{cia}_{versao_txt}.xlsx", as_attachment=True)
    except Exception as e:
        flash(f'Erro: {e}', 'error')
        return redirect(url_for('index'))

# --- ROTA DE EXPORTAR FREQUÊNCIA (LAYOUT CONTROLADOR FINANCEIRO) ---
@app.route('/exportar_excel')
@login_required
def exportar_excel():
    try:
        # 1. Pega os parâmetros da URL (do Modal)
        filtro_posto = request.args.get('filtro_posto', 'Todos')
        mes = request.args.get('mes', datetime.now().month, type=int)
        ano = request.args.get('ano', datetime.now().year, type=int)
        cia = request.args.get('cia_filtro', 'TODAS')
        
        import calendar
        u_dia = calendar.monthrange(ano, mes)[1]
        
        # 2. Configuração inicial do Excel
        output = BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        ws = workbook.add_worksheet("Frequencia")
        
        ws.set_paper(9) # A4
        ws.set_landscape()
        ws.set_margins(0.3, 0.3, 0.3, 0.3)
        ws.fit_to_pages(1, 0)
        
        # --- ESTILOS ---
        fmt_titulo = workbook.add_format({'bold': True, 'font_name': 'Arial', 'font_size': 12, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#D9D9D9'})
        fmt_header = workbook.add_format({'bold': True, 'font_name': 'Arial', 'font_size': 8, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#EFEFEF', 'text_wrap': True})
        fmt_dia_header = workbook.add_format({'bold': True, 'font_name': 'Arial', 'font_size': 8, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#FFFFFF'})
        fmt_texto_centro = workbook.add_format({'font_name': 'Arial', 'font_size': 8, 'align': 'center', 'valign': 'vcenter', 'border': 1})
        fmt_texto_esq = workbook.add_format({'font_name': 'Arial', 'font_size': 8, 'align': 'left', 'valign': 'vcenter', 'border': 1, 'indent': 1})
        fmt_total = workbook.add_format({'bold': True, 'font_name': 'Arial', 'font_size': 8, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#F2F2F2'})
        fmt_bonus = workbook.add_format({'bold': True, 'font_name': 'Arial', 'font_size': 8, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#E6E6FF'})

        # --- LARGURAS ---
        ws.set_column('A:A', 8)   # POSTO
        ws.set_column('B:B', 10)  # RE
        ws.set_column('C:C', 35)  # NOME
        ws.set_column(3, 33, 2.8) # Dias 1-31
        ws.set_column(34, 35, 6)  # DIAS e VALOR
        ws.set_column(36, 36, 7)  # BÔNUS

        # --- CABEÇALHO ---
        titulo_texto = f'CONTROLE DE FREQUÊNCIA - {cia if cia != "TODAS" else "37º BPM/M"}'
        subtitulo = f'REFERÊNCIA: {mes:02d}/{ano}'
        
        ws.merge_range('A1:AK1', titulo_texto, fmt_titulo)
        ws.merge_range('A2:AK2', subtitulo, fmt_titulo)

        ws.write(2, 0, "POSTO", fmt_header)
        ws.write(2, 1, "RE", fmt_header)
        ws.write(2, 2, "NOME DE GUERRA", fmt_header)
        
        col_idx = 3
        for d in range(1, 32):
            if d <= u_dia: ws.write(2, col_idx, str(d), fmt_dia_header)
            else: ws.write(2, col_idx, "X", fmt_header)
            col_idx += 1
            
        ws.write(2, col_idx, "DIAS", fmt_header)
        ws.write(2, col_idx+1, "VALOR", fmt_header)
        ws.write(2, col_idx+2, "BÔNUS", fmt_header)

        # --- FILTRO DE DADOS (CORRIGIDO) ---
        lista_oficiais = [
            'Coronel PM', 'Cel PM', 
            'Tenente Coronel PM', 'Ten Cel PM', 
            'Major PM', 'Maj PM', 
            'Capitão PM', 'Cap PM', 
            '1º Tenente PM', '1º Ten PM', 
            '2º Tenente PM', '2º Ten PM', 
            'Aspirante a Oficial', 'Asp Of PM'
        ]
        
        # Pega todos os ativos e já ordena por nome para ficar bonitinho
        q = Policial.query.filter_by(ativo=True).order_by(Policial.nome)
        
        # Filtra Cia
        if cia != 'TODAS': 
            q = q.filter_by(cia=cia)
            
        # Filtra Oficiais ou Praças
        if filtro_posto == 'Oficiais':
            q = q.filter(Policial.posto.in_(lista_oficiais))
        elif filtro_posto == 'Pracas':
            q = q.filter(Policial.posto.notin_(lista_oficiais))
            
        # AQUI ESTAVA O SEU ERRO (O nome da variável pms estava faltando!)
        pms = q.all()

        # --- ESCRITA DOS DADOS ---
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
                    data_dia = datetime(ano, mes, d)
                    sigla = calcular_escala_dia(pm, data_dia)
                    cod = traduzir_para_frequencia(sigla)
                    
                    if cod in ['0', '1', '2', '3']:
                        dias_trabalhados += 1
                        soma_valor += int(cod)
                        val_cell = cod
                    else:
                        # Para aceitar o '?', 'A', 'LP', etc (que você tinha corrigido mais cedo)
                        val_cell = cod
                    
                    # Cálculo Bônus
                    if cod in ['0', '1', '2', '3', 'F', 'A', 'LP']:
                        bonus_count += 1
                    
                ws.write(linha, c, val_cell, fmt_texto_centro)
                c += 1
            
            ws.write(linha, c, dias_trabalhados, fmt_total)   
            ws.write(linha, c+1, soma_valor, fmt_total)       
            ws.write(linha, c+2, bonus_count, fmt_bonus)      
            
            linha += 1

        # --- ASSINATURAS ---
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
    
    if cia == 'TODAS':
        pms_raw = Policial.query.filter_by(ativo=True).all()
        titulo = "TODAS AS CIAS"
    else:
        pms_raw = Policial.query.filter_by(cia=cia, ativo=True).all()
        titulo = cia.upper()
    
    if cia == 'EM': pms = ordenar_em_agrupado(pms_raw)
    else: pms = ordenar_por_antiguidade(pms_raw)
    
    # Você precisará ter o template 'relatorio_efetivo.html' na pasta templates
    # Se não tiver, avise que crio.
    return render_template('relatorio_efetivo.html', pms=pms, cia=titulo, hoje=datetime.now())

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