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

# --- NOVO MODELO: CONTROLE DE VERSÃO DA ESCALA ---
class EscalaVersao(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    cia = db.Column(db.String(50))
    data_inicio = db.Column(db.Date)
    tipo_filtro = db.Column(db.String(20))
    versao_atual = db.Column(db.Integer, default=1)

# --- NOVO MODELO: PROTOCOLO DE DOCUMENTOS ---
class DocumentoProtocolo(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    numero_sequencial = db.Column(db.Integer)
    ano = db.Column(db.Integer)
    numero_protocolo_formatado = db.Column(db.String(20), unique=True)
    
    tipo_documento = db.Column(db.String(50))
    numero_documento_origem = db.Column(db.String(50))
    
    # QUEM ENTREGOU (Entrada)
    entregue_posto = db.Column(db.String(20))
    entregue_re = db.Column(db.String(20))
    entregue_nome = db.Column(db.String(100))
    entregue_cia = db.Column(db.String(50))
    
    # QUEM RECEBEU (Entrada)
    recebido_posto = db.Column(db.String(20))
    recebido_re = db.Column(db.String(20))
    recebido_nome = db.Column(db.String(100))
    recebido_cia = db.Column(db.String(50))
    
    data_entrada = db.Column(db.DateTime, default=datetime.now)
    
    # --- DADOS DE SAÍDA (QUEM LEVOU) ---
    data_saida = db.Column(db.DateTime)
    
    # Quem Retirou o documento (Saída)
    retirado_posto = db.Column(db.String(20))
    retirado_re = db.Column(db.String(20))
    retirado_nome = db.Column(db.String(100))
    retirado_cia = db.Column(db.String(50))
    
    # Quem Entregou na Saída (Operador da Seção)
    despachante_posto = db.Column(db.String(20))
    despachante_re = db.Column(db.String(20))
    despachante_nome = db.Column(db.String(100))
    
    # --- DADOS DE ARQUIVAMENTO ---
    data_arquivamento = db.Column(db.DateTime)
    local_arquivo = db.Column(db.String(100)) # Ex: Pasta 3, Gaveta 2
    
    status = db.Column(db.String(20), default='Aberto') # Aberto, Saiu, Arquivado
    observacao = db.Column(db.Text)

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

def calcular_escala_dia(policial, data_obj):
    if isinstance(data_obj, datetime): data = data_obj.date()
    else: data = data_obj

    alteracao = AlteracaoEscala.query.filter_by(policial_id=policial.id, data=data).first()
    if alteracao: return alteracao.valor

    tipo = policial.tipo_escala
    escalas_12x36 = ['S/3', 'S/4', 'S/5', 'S/6', '12x36']
    
    if tipo in escalas_12x36:
        mes = data.month
        dia_mes = data.day
        inicio_pares = policial.inicio_pares
        if mes % 2 == 0: comecar_pares = not inicio_pares
        else: comecar_pares = inicio_pares
        trabalha_hoje = (dia_mes % 2 == 0) if comecar_pares else (dia_mes % 2 == 1)
        return tipo if trabalha_hoje else 'F'
    else:
        if data.weekday() >= 5 or e_feriado(data): return 'F'
        return tipo 

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
    # Férias (FN, FS no sistema antigo, mas Férias é F)
    if sigla in ['FN', 'FS', 'FÉRIAS', 'FERIAS']: return 'F'
    
    # Licença-Prêmio
    if sigla == 'LP': return 'LP'
    
    # Falta ao Serviço
    if sigla in ['FALTA', 'FALTOU', 'F.S.', 'FS']: return 'FS'
    
    # Demais Afastamentos -> A
    lista_afastamentos = [
        'A', 'FOLGA', 'FOLGA', 'JM', 'FM', 'DS', 'DR', 'FR', 'PF', 
        'CV', 'LTS', 'EAP', 'LT', 'LG', 'PT', 'LA', 'NP', 'LICENÇA', 'AFASTADO',
        'DISPENSA', 'LICENCA'
    ]
    # Se for F (Folga comum), não gera diária, mas no sistema financeiro pode ser A ou vazio
    if sigla == 'F': return 'A' 
    if sigla in lista_afastamentos: return 'A'

    # === CÁLCULO POR DURAÇÃO (NÚMEROS) ===
    
    # "0": < 8h (Meio Período 6h)
    if sigla in ['M1', 'M2', 'MEIO', '6H']: 
        return '0'

    # "1": >= 8h e < 12h (Expediente Administrativo 8h-10h)
    if sigla in ['S', 'S/1', 'S/2', 'ADM', 'EXP', '8H', '9H', '10H']: 
        return '1'

    # "2": >= 12h e < 18h (Plantões 12x36 Diurno/Noturno)
    # S/3, S/4, S/5, S/6 são escalas de 12 horas
    if sigla in ['S/3', 'S/4', 'S/5', 'S/6', '12X36', 'PLANTÃO', '12H']: 
        return '2'

    # "3": >= 18h (Plantões 24h, Dobras, Sobreaviso Longo)
    # SS, SR, EE, S* geralmente são escalas longas ou especiais
    if sigla in ['SS', 'SR', 'EE', 'S*', '24H', '18H', 'DOBRA']: 
        return '3'
    
    # Se não reconhecer, retorna ?
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
    
    tem_em = any(pm.cia == 'EM' for pm in pms_raw)
    if tem_em:
        pms_em = [p for p in pms_raw if p.cia == 'EM']
        pms_outros = [p for p in pms_raw if p.cia != 'EM']
        pms_ordenados = ordenar_em_agrupado(pms_em) + ordenar_por_antiguidade(pms_outros)
    else:
        pms_ordenados = ordenar_por_antiguidade(pms_raw)
    
    e_admin = (getattr(current_user, 'perfil', 'usuario') == 'admin')
    return render_template('index.html', pms=pms_ordenados, user=current_user, e_admin=e_admin, hoje=datetime.now().date())

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
        re_form = request.form.get('re', '').strip()
        if not re_form: return redirect(url_for('index'))

        dados = {
            're': re_form, 'posto': request.form.get('posto'), 'nome': request.form.get('nome'),
            'nome_guerra': request.form.get('nome_guerra'), 'cia': request.form.get('cia'),
            'secao_em': request.form.get('secao_em'), 'funcao': request.form.get('funcao'),
            'viatura': request.form.get('viatura'), 'equipe': request.form.get('equipe'),
            'tipo_escala': request.form.get('tipo_escala'), 
            'inicio_pares': (request.form.get('inicio_pares') == 'on'),
            'status_saude': request.form.get('status_saude'), 
            'restricao_tipo': request.form.get('restricao_tipo'),
            'data_fim_restricao': datetime.strptime(request.form.get('data_fim_restricao'), '%Y-%m-%d').date() if request.form.get('data_fim_restricao') else None
        }

        if id_pm:
            pm = db.session.get(Policial, int(id_pm))
            if pm:
                if pm.cia != dados['cia']: registrar_movimentacao(pm.id, "Mudança Cia", pm.cia, dados['cia'])
                for k, v in dados.items(): setattr(pm, k, v)
                msg = 'Cadastro atualizado com sucesso!'
        else:
            if Policial.query.filter_by(re=re_form).first():
                flash('Já existe um PM com este RE!', 'error'); return redirect(url_for('index'))
            novo = Policial(**dados)
            db.session.add(novo); db.session.commit()
            registrar_movimentacao(novo.id, "Admissão", "N/A", dados['cia'])
            msg = 'Novo policial cadastrado!'
            
        db.session.commit()
        flash(msg, 'success')
    except Exception as e:
        db.session.rollback()
        flash(f'Erro ao salvar: {e}', 'error')
        
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
        elif sigla in ['S', 'S/1', 'S/2', 'M1', 'M2']: cor = 'bg-success text-white'
        elif sigla in ['S/3', 'S/4', 'S/5', 'S/6']: cor = 'bg-primary text-white'
        elif sigla in ['LTS', 'CV', 'FN', 'FS']: cor = 'bg-warning text-dark'
        elif sigla in ['SS', 'SR', 'EE']: cor = 'bg-info text-dark'
        
        escala_vis.append({'dia': d, 'valor': sigla, 'classe': cor, 'data_iso': dt.strftime('%Y-%m-%d')})

    nav = {'ant_m': mes-1 if mes>1 else 12, 'ant_a': ano if mes>1 else ano-1, 'prox_m': mes+1 if mes<12 else 1, 'prox_a': ano if mes<12 else ano+1}
    return render_template('editar.html', pm=pm, mes=mes, ano=ano, escala=escala_vis, dias_vazios=dias_vazios, totais={'trabalhado': total_trabalhado}, nav=nav)

@app.route('/salvar_alteracao', methods=['POST'])
@login_required
def salvar_alteracao():
    try:
        id_pm = request.form.get('id_pm')
        inicio = datetime.strptime(request.form.get('data_inicio'), '%Y-%m-%d').date()
        fim_str = request.form.get('data_fim')
        fim = datetime.strptime(fim_str, '%Y-%m-%d').date() if fim_str else inicio
        valor = request.form.get('valor')
        
        for i in range((fim - inicio).days + 1):
            dia = inicio + timedelta(days=i)
            alt = AlteracaoEscala.query.filter_by(policial_id=id_pm, data=dia).first()
            if valor == 'LIMPAR':
                if alt: db.session.delete(alt)
            else:
                if not alt: db.session.add(AlteracaoEscala(policial_id=id_pm, data=dia, valor=valor))
                else: alt.valor = valor
        db.session.commit()
        flash('Atualizado!', 'success')
    except: pass
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
        # 1. Parâmetros
        mes = request.args.get('mes', datetime.now().month, type=int)
        ano = request.args.get('ano', datetime.now().year, type=int)
        tipo = request.args.get('tipo_filtro', 'TODOS')
        cia = request.args.get('cia_filtro', 'TODAS')
        
        u_dia = calendar.monthrange(ano, mes)[1]
        
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
        
        # Estilo BÔNUS (AK)
        fmt_bonus = workbook.add_format({'bold': True, 'font_name': 'Arial', 'font_size': 8, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#E6E6FF'})

        # --- LARGURAS ---
        ws.set_column('A:A', 8)   # POSTO
        ws.set_column('B:B', 10)  # RE
        ws.set_column('C:C', 35)  # NOME
        ws.set_column(3, 33, 2.8) # Dias 1-31
        ws.set_column(34, 35, 6)  # DIAS e VALOR
        ws.set_column(36, 36, 7)  # BÔNUS (Mais larguinho)

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
        ws.write(2, col_idx+2, "BÔNUS", fmt_header) # Coluna AK

        # --- DADOS ---
        q = Policial.query.filter_by(ativo=True)
        if cia != 'TODAS': q = q.filter_by(cia=cia)
        pms_raw = q.all()
        pms = filtrar_por_tipo_hierarquico(pms_raw, tipo)
        
        if cia == 'EM': pms = ordenar_em_agrupado(pms)
        elif cia != 'TODAS': pms = ordenar_cia_agrupado(pms)
        else: pms = ordenar_por_antiguidade(pms)

        linha = 3
        for pm in pms:
            ws.write(linha, 0, pm.posto, fmt_texto_centro)
            ws.write(linha, 1, pm.re, fmt_texto_centro)
            ws.write(linha, 2, pm.nome_guerra or pm.nome, fmt_texto_esq)
            
            c = 3
            dias_trabalhados = 0
            soma_valor = 0
            bonus_count = 0 # Inicializa contador do Bônus (AK8)
            
            for d in range(1, 32):
                val_cell = ""
                if d <= u_dia:
                    data_dia = datetime(ano, mes, d)
                    sigla = calcular_escala_dia(pm, data_dia)
                    cod = traduzir_para_frequencia(sigla)
                    
                    # Cálculo Padrão (DIAS e VALOR)
                    if cod in ['0', '1', '2', '3']:
                        dias_trabalhados += 1
                        soma_valor += int(cod)
                        val_cell = cod
                    elif cod in ['A', 'F', 'LP', 'FS']:
                        val_cell = cod
                    
                    # --- CÁLCULO DA COLUNA BÔNUS (AK8) ---
                    # Fórmula Excel: =CONT.SE("0")*1 + CONT.SE("1")*1 ... + CONT.SE("F")*1 + CONT.SE("A")*1
                    # Basicamente: Tudo que é trabalho ou afastamento remunerado conta 1.
                    # FS (Falta) não conta.
                    if cod in ['0', '1', '2', '3', 'F', 'A', 'LP']:
                        bonus_count += 1
                    
                ws.write(linha, c, val_cell, fmt_texto_centro)
                c += 1
            
            # Totais
            ws.write(linha, c, dias_trabalhados, fmt_total)   # Coluna DIAS
            ws.write(linha, c+1, soma_valor, fmt_total)       # Coluna VALOR
            ws.write(linha, c+2, bonus_count, fmt_bonus)      # Coluna BÔNUS (Nova)
            
            linha += 1

        # Assinaturas
        linha += 3
        ws.merge_range(f'C{linha}:J{linha}', '_' * 30, fmt_texto_centro)
        ws.merge_range(f'C{linha+1}:J{linha+1}', 'ENCARREGADO DE SECRETARIA', fmt_texto_centro)
        ws.merge_range(f'R{linha}:Y{linha}', '_' * 30, fmt_texto_centro)
        ws.merge_range(f'R{linha+1}:Y{linha+1}', f'CMT DA {cia if cia != "TODAS" else "UNIDADE"}', fmt_texto_centro)

        workbook.close()
        output.seek(0)
        return send_file(output, download_name=f"Frequencia_{cia}_{mes}_{ano}.xlsx", as_attachment=True)

    except Exception as e:
        flash(f"Erro Excel: {e}", 'error')
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

# --- ROTAS DE PROTOCOLO (NOVA FUNCIONALIDADE) ---

@app.route('/protocolo')
@login_required
def protocolo():
    docs = DocumentoProtocolo.query.order_by(desc(DocumentoProtocolo.id)).all()
    return render_template('protocolo.html', docs=docs, hoje=datetime.now())

@app.route('/novo_protocolo', methods=['POST'])
@login_required
def novo_protocolo():
    try:
        ano_atual = datetime.now().year
        ultimo = DocumentoProtocolo.query.filter_by(ano=ano_atual).order_by(desc(DocumentoProtocolo.numero_sequencial)).first()
        novo_seq = (ultimo.numero_sequencial + 1) if ultimo else 1
        protocolo_fmt = f"{ano_atual}/{novo_seq:04d}"
        
        novo_doc = DocumentoProtocolo(
            ano=ano_atual, numero_sequencial=novo_seq, numero_protocolo_formatado=protocolo_fmt,
            tipo_documento=request.form.get('tipo'),
            numero_documento_origem=request.form.get('numero_origem'),
            entregue_posto=request.form.get('ent_posto'), entregue_re=request.form.get('ent_re'),
            entregue_nome=request.form.get('ent_nome'), entregue_cia=request.form.get('ent_cia'),
            recebido_posto=request.form.get('rec_posto'), recebido_re=request.form.get('rec_re'),
            recebido_nome=request.form.get('rec_nome'), recebido_cia=request.form.get('rec_cia'),
            observacao=request.form.get('observacao')
        )
        db.session.add(novo_doc); db.session.commit()
        flash(f'Entrada {protocolo_fmt} registrada!', 'success')
    except Exception as e: flash(f'Erro: {e}', 'error')
    return redirect(url_for('protocolo'))

@app.route('/nova_saida_direta', methods=['POST'])
@login_required
def nova_saida_direta():
    try:
        ano_atual = datetime.now().year
        
        # Gera Numerador Único
        ultimo = DocumentoProtocolo.query.filter_by(ano=ano_atual).order_by(desc(DocumentoProtocolo.numero_sequencial)).first()
        novo_seq = (ultimo.numero_sequencial + 1) if ultimo else 1
        protocolo_fmt = f"{ano_atual}/{novo_seq:04d}"
        
        novo_doc = DocumentoProtocolo(
            ano=ano_atual,
            numero_sequencial=novo_seq,
            numero_protocolo_formatado=protocolo_fmt,
            
            tipo_documento=request.form.get('tipo'),
            numero_documento_origem=request.form.get('numero_origem'),
            
            # Como é saída direta, definimos data de entrada = agora (para constar no registro)
            data_entrada=datetime.now(),
            
            # DADOS DE SAÍDA (O Caminho Inverso)
            data_saida=datetime.now(),
            status='Saiu',
            
            # QUEM LEVOU (Destinatário)
            retirado_posto=request.form.get('ret_posto'),
            retirado_re=request.form.get('ret_re'),
            retirado_nome=request.form.get('ret_nome'),
            retirado_cia=request.form.get('ret_cia'),
            
            # QUEM DESPACHOU (Sua Seção)
            despachante_posto=request.form.get('desp_posto'),
            despachante_re=request.form.get('desp_re'),
            despachante_nome=request.form.get('desp_nome'),
            
            observacao=f"[Expediente Próprio/Saída Direta]: {request.form.get('observacao')}"
        )
        
        db.session.add(novo_doc)
        db.session.commit()
        flash(f'Saída Direta {protocolo_fmt} registrada!', 'success')
        
    except Exception as e:
        flash(f'Erro: {e}', 'error')
        
    return redirect(url_for('protocolo'))

@app.route('/registrar_saida', methods=['POST'])
@login_required
def registrar_saida():
    try:
        doc = db.session.get(DocumentoProtocolo, request.form.get('doc_id'))
        if doc:
            doc.data_saida = datetime.now()
            doc.status = 'Saiu'
            # Quem Retirou (Destinatário)
            doc.retirado_posto = request.form.get('ret_posto')
            doc.retirado_re = request.form.get('ret_re')
            doc.retirado_nome = request.form.get('ret_nome')
            doc.retirado_cia = request.form.get('ret_cia')
            # Quem Entregou (Operador)
            doc.despachante_posto = request.form.get('desp_posto')
            doc.despachante_re = request.form.get('desp_re')
            doc.despachante_nome = request.form.get('desp_nome')
            
            obs_extra = request.form.get('obs_saida')
            if obs_extra: doc.observacao = (doc.observacao or '') + f"\n[Saída]: {obs_extra}"
            
            db.session.commit()
            flash('Saída registrada com sucesso!', 'success')
    except Exception as e: flash(f'Erro: {e}', 'error')
    return redirect(url_for('protocolo'))

@app.route('/arquivar_documento', methods=['POST'])
@login_required
def arquivar_documento():
    try:
        doc = db.session.get(DocumentoProtocolo, request.form.get('doc_id'))
        if doc:
            doc.data_arquivamento = datetime.now()
            doc.status = 'Arquivado'
            doc.local_arquivo = request.form.get('local_arquivo')
            
            obs_arq = request.form.get('obs_arquivo')
            if obs_arq: doc.observacao = (doc.observacao or '') + f"\n[Arquivo]: {obs_arq}"
            
            db.session.commit()
            flash('Documento arquivado.', 'secondary')
    except Exception as e: flash(f'Erro: {e}', 'error')
    return redirect(url_for('protocolo'))

# --- ROTA DE RELATÓRIO DE PROTOCOLO (EXCEL) ---
@app.route('/gerar_relatorio_protocolo', methods=['POST'])
@login_required
def gerar_relatorio_protocolo():
    try:
        data_inicio_str = request.form.get('data_inicio')
        data_fim_str = request.form.get('data_fim')
        
        if not data_inicio_str or not data_fim_str:
            flash('Selecione o período inicial e final.', 'error')
            return redirect(url_for('protocolo'))
            
        # Converte para objetos datetime
        inicio = datetime.strptime(data_inicio_str, '%Y-%m-%d')
        # Ajusta o fim para pegar o dia inteiro (23:59:59)
        fim = datetime.strptime(data_fim_str, '%Y-%m-%d').replace(hour=23, minute=59, second=59)
        
        # Busca no Banco (Filtrando pela Data de Entrada)
        docs = DocumentoProtocolo.query.filter(
            DocumentoProtocolo.data_entrada >= inicio,
            DocumentoProtocolo.data_entrada <= fim
        ).order_by(DocumentoProtocolo.numero_sequencial).all()
        
        # --- GERAÇÃO DO EXCEL ---
        output = BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        ws = workbook.add_worksheet("Protocolo")
        
        # Configuração de Impressão (Paisagem)
        ws.set_landscape()
        ws.set_paper(9) # A4
        ws.set_margins(0.2, 0.2, 0.2, 0.2)
        
        # Estilos
        fmt_titulo = workbook.add_format({'bold': True, 'font_name': 'Arial', 'font_size': 14, 'align': 'center', 'bg_color': '#D9D9D9', 'border': 1})
        fmt_header = workbook.add_format({'bold': True, 'font_name': 'Arial', 'font_size': 9, 'align': 'center', 'bg_color': '#4a148c', 'font_color': 'white', 'border': 1})
        fmt_cell = workbook.add_format({'font_name': 'Arial', 'font_size': 9, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'text_wrap': True})
        fmt_date = workbook.add_format({'font_name': 'Arial', 'font_size': 9, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'num_format': 'dd/mm/yyyy hh:mm'})
        
        # Cabeçalho
        texto_titulo = f"RELATÓRIO DE PROTOCOLO - DE {inicio.strftime('%d/%m/%Y')} A {fim.strftime('%d/%m/%Y')}"
        ws.merge_range('A1:J1', texto_titulo, fmt_titulo)
        
        # Colunas
        headers = [
            "Protocolo", "Data Entrada", "Tipo Doc", "Nº Origem", 
            "Entregue Por (Origem)", "Recebido Por", 
            "Status", "Data Saída", "Destino / Retirado Por", "Observações"
        ]
        
        ws.set_column('A:A', 12) # Protocolo
        ws.set_column('B:B', 15) # Data Ent
        ws.set_column('C:D', 12) # Tipo/Num
        ws.set_column('E:F', 25) # Nomes
        ws.set_column('G:G', 10) # Status
        ws.set_column('H:H', 15) # Data Sai
        ws.set_column('I:I', 25) # Destino
        ws.set_column('J:J', 30) # Obs
        
        ws.write_row('A2', headers, fmt_header)
        
        linha = 2
        for doc in docs:
            ws.write(linha, 0, doc.numero_protocolo_formatado, fmt_cell)
            ws.write(linha, 1, doc.data_entrada, fmt_date)
            ws.write(linha, 2, doc.tipo_documento, fmt_cell)
            ws.write(linha, 3, doc.numero_documento_origem or '-', fmt_cell)
            
            # Entregue Por
            if doc.entregue_nome:
                origem = f"{doc.entregue_posto} {doc.entregue_nome}\n({doc.entregue_cia})"
            else:
                origem = "Produção Interna"
            ws.write(linha, 4, origem, fmt_cell)
            
            # Recebido Por
            if doc.recebido_nome:
                rec = f"{doc.recebido_re} {doc.recebido_nome}"
            else:
                rec = f"{doc.despachante_nome} (Despacho)"
            ws.write(linha, 5, rec, fmt_cell)
            
            ws.write(linha, 6, doc.status, fmt_cell)
            
            # Saída
            ws.write(linha, 7, doc.data_saida if doc.data_saida else "-", fmt_date if doc.data_saida else fmt_cell)
            
            # Destino
            if doc.status == 'Arquivado':
                destino = f"ARQUIVO: {doc.local_arquivo}"
            elif doc.status == 'Saiu':
                destino = f"{doc.retirado_posto} {doc.retirado_nome}\n({doc.retirado_cia})"
            else:
                destino = "-"
            ws.write(linha, 8, destino, fmt_cell)
            
            ws.write(linha, 9, doc.observacao or '', fmt_cell)
            linha += 1
            
        workbook.close()
        output.seek(0)
        
        nome_arquivo = f"Protocolo_{inicio.strftime('%d-%m')}_a_{fim.strftime('%d-%m')}.xlsx"
        return send_file(output, download_name=nome_arquivo, as_attachment=True)
        
    except Exception as e:
        flash(f'Erro ao gerar relatório: {e}', 'error')
        return redirect(url_for('protocolo'))

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
            
    app.run(debug=True)