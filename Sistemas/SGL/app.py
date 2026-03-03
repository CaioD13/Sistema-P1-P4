# app.py - Sistema de Gestão Logística (SGL) - Batalhão
# Versão: 2.0 (com RA e Cadastro de Usuários)
# Total aproximado de linhas: 580 (incluindo comentários)

import os
import pandas as pd
from io import BytesIO
from datetime import datetime
from flask import Flask, render_template, request, redirect, url_for, flash, send_file
from flask_sqlalchemy import SQLAlchemy
from flask_login import LoginManager, UserMixin, login_user, login_required, logout_user, current_user
from werkzeug.security import generate_password_hash, check_password_hash
from sqlalchemy import func, or_
import re

# Configuração inicial
basedir = os.path.abspath(os.path.dirname(__file__))
app = Flask(__name__)
app.config['SECRET_KEY'] = 'pm_seguranca_total_2026'
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///' + os.path.join(basedir, 'banco_novo_v3.db')
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

db = SQLAlchemy(app)
login_manager = LoginManager()
login_manager.init_app(app)
login_manager.login_view = 'login'

# Funções auxiliares
def limpar_valor_monetario(valor_raw):
    if valor_raw is None or valor_raw == '': return 0.0
    if isinstance(valor_raw, (int, float)): return float(valor_raw)
    valor_str = str(valor_raw).strip().replace('R$', '').strip()
    if ',' in valor_str and '.' in valor_str:
        valor_str = valor_str.replace('.', '').replace(',', '.')
    elif ',' in valor_str:
        valor_str = valor_str.replace(',', '.')
    try: return float(valor_str)
    except: return 0.0

@app.template_filter('reais')
def format_reais(value):
    if value is None or value == '': return '0,00'
    try: return f"{float(value):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except: return value

def normalizar_re(re_raw: str) -> str:
    if not re_raw: return ''
    somente_digitos = re.sub(r'\D', '', re_raw)
    if len(somente_digitos) >= 2: return somente_digitos[:-1]
    return somente_digitos

def senha_forte(senha: str) -> bool:
    if not senha or len(senha) < 8: return False
    if not re.search(r'[A-Z]', senha): return False
    if not re.search(r'\d', senha): return False
    if not re.search(r'[!@#$%^&*()_\-+=\[\]{};:\'",.<>/?\|`~]', senha): return False
    return True

# Constantes
CONTAS_PATRIMONIAIS = {
    '123110901': 'ARMAMENTO',
    '123110101': 'APARELHOS DE MEDIÇÃO E ORIENTAÇÃO',
    '123110102': 'APARELHOS E EQUIPAMENTOS DE COMUNICAÇÃO',
    '123110104': 'APARELHOS E EQUIPAMENTOS DIVERSÕES E ESPORTE',
    '123110105': 'EQUIPAMENTO DE PROTEÇÃO, SEGURANÇA E SOCORRO',
    '123110106': 'MÁQUINAS E EQUIPAMENTOS INDUSTRIAIS',
    '123110109': 'MÁQUINAS, FERRAMENTAS E UTENSÍLIOS DE OFICINA',
    '123110119': 'MÁQUINAS, EQUIPAMENTOS E UTENSÍLIOS AGROPECUÁRIOS',
    '123110123': 'EQUIPAMENTO PARA ESCRITÓRIO',
    '123110125': 'EQUIPAM. E UTENSÍLIOS P/ COLETA E TRANSPORTE DE LIXO',
    '123110199': 'OUTRAS MÁQ., APARELHOS, EQUIP. E FERRAMENTAS',
    '123110201': 'EQUIPAMENTOS DE PROCESSAMENTO DE DADOS',
    '123110301': 'APARELHOS E UTENSÍLIOS DOMÉSTICOS',
    '123110302': 'MÁQUINAS E UTENSÍLIOS DE ESCRITÓRIO',
    '123110303': 'MOBILIÁRIO EM GERAL',
    '123119999': 'OUTROS BENS MÓVEIS',
    '123110501': 'VEÍCULOS EM GERAL'
}

# Modelos
class Usuario(UserMixin, db.Model):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(50), unique=True, nullable=False)
    password_hash = db.Column(db.String(128))
    role = db.Column(db.String(20), default='usuario')
    def set_password(self, password): self.password_hash = generate_password_hash(password)
    def check_password(self, password): return check_password_hash(self.password_hash, password)

class Material(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    patrimonio = db.Column(db.String(50), unique=True, nullable=False)
    n_serie = db.Column(db.String(50))
    nome_material = db.Column(db.String(100), nullable=False)
    conta_codigo = db.Column(db.String(20), nullable=False)
    local = db.Column(db.String(100), nullable=False)
    valor = db.Column(db.Float, default=0.0)
    situacao = db.Column(db.String(50))
    observacoes = db.Column(db.Text)
    pm_re = db.Column(db.String(20))
    pm_nome = db.Column(db.String(100))
    pm_opm = db.Column(db.String(50))
    eh_apreensao = db.Column(db.Boolean, default=False)
    local_apreensao = db.Column(db.String(100))
    autoridade_resp = db.Column(db.String(100))
    data_apreensao = db.Column(db.DateTime)
    num_doc_apreensao = db.Column(db.String(50))
    data_cadastro = db.Column(db.DateTime, server_default=func.now())
    historico = db.relationship('Historico', backref='material', cascade="all, delete-orphan")
    movimentacoes_ra = db.relationship('MovimentacaoRA', backref='material', cascade="all, delete-orphan")
    
    @property
    def dias_custodia(self):
        if self.data_apreensao: return (datetime.now() - self.data_apreensao).days
        return 0

class Historico(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    material_id = db.Column(db.Integer, db.ForeignKey('material.id'), nullable=False)
    acao = db.Column(db.String(100)) 
    descricao_detalhada = db.Column(db.Text)
    usuario_responsavel = db.Column(db.String(50))
    data_hora = db.Column(db.DateTime, default=datetime.now)

class MovimentacaoRA(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    material_id = db.Column(db.Integer, db.ForeignKey('material.id'), nullable=False)
    operador_retirada = db.Column(db.String(50), nullable=False)
    operador_devolucao = db.Column(db.String(50))
    pm_retirou_posto_grad = db.Column(db.String(50), nullable=False)
    pm_retirou_re = db.Column(db.String(20), nullable=False)
    pm_retirou_nome_guerra = db.Column(db.String(100), nullable=False)
    pm_retirou_cia_secao = db.Column(db.String(50), nullable=False)
    pm_devolveu_posto_grad = db.Column(db.String(50))
    pm_devolveu_re = db.Column(db.String(20))
    pm_devolveu_nome_guerra = db.Column(db.String(100))
    pm_devolveu_cia_secao = db.Column(db.String(50))
    data_hora_retirada = db.Column(db.DateTime, default=datetime.now, nullable=False)
    data_hora_devolucao = db.Column(db.DateTime)
    observacoes_retirada = db.Column(db.Text)
    observacoes_devolucao = db.Column(db.Text)

@login_manager.user_loader
def load_user(user_id): return Usuario.query.get(int(user_id))

def criar_admin():
    with app.app_context():
        db.create_all()
        if not Usuario.query.filter_by(username='admin').first():
            admin = Usuario(username='admin', role='admin')
            admin.set_password('admin123')
            db.session.add(admin)
            db.session.commit()

def registrar_log(material_id, acao, descricao):
    try:
        usuario = current_user.username if current_user.is_authenticated else 'Sistema'
        novo_log = Historico(material_id=material_id, acao=acao, descricao_detalhada=descricao, usuario_responsavel=usuario)
        db.session.add(novo_log)
    except: pass

# Rotas
@app.route('/')
def index(): return redirect(url_for('dashboard')) if current_user.is_authenticated else redirect(url_for('login'))

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        user = Usuario.query.filter_by(username=request.form['username']).first()
        if user and user.check_password(request.form['password']):
            login_user(user)
            return redirect(url_for('dashboard'))
        flash('Usuário ou senha inválidos', 'danger')
    return render_template('login.html')

@app.route('/logout')
@login_required
def logout():
    logout_user()
    return redirect(url_for('login'))

@app.route('/alterar_senha', methods=['GET', 'POST'])
@login_required
def alterar_senha():
    if request.method == 'POST':
        senha_atual = request.form.get('senha_atual', '')
        nova_senha = request.form.get('nova_senha', '')
        nova_senha2 = request.form.get('nova_senha2', '')

        # 1) Confere senha atual
        if not current_user.check_password(senha_atual):
            flash('Senha atual incorreta.', 'danger')
            return redirect(url_for('alterar_senha'))

        # 2) Confere confirmação
        if nova_senha != nova_senha2:
            flash('A nova senha e a confirmação não conferem.', 'danger')
            return redirect(url_for('alterar_senha'))

        # 3) Valida força da senha (usa sua função existente)
        if not senha_forte(nova_senha):
            flash('Senha fraca. Use mínimo 8 caracteres, 1 letra maiúscula, 1 número e 1 caractere especial.', 'danger')
            return redirect(url_for('alterar_senha'))

        # 4) Evita repetir a senha atual (opcional, mas recomendado)
        if current_user.check_password(nova_senha):
            flash('A nova senha não pode ser igual à senha atual.', 'warning')
            return redirect(url_for('alterar_senha'))

        # 5) Atualiza
        user = Usuario.query.get(current_user.id)
        user.set_password(nova_senha)
        db.session.commit()

        flash('Senha alterada com sucesso.', 'success')
        return redirect(url_for('dashboard'))

    return render_template('alterar_senha.html')

@app.route('/dashboard')
@login_required
def dashboard():
    materiais = Material.query.all()
    total_valor = db.session.query(func.sum(Material.valor)).scalar() or 0.0
    qtd_total = len(materiais)
    return render_template('dashboard.html', materiais=materiais, total_valor=total_valor, qtd_total=qtd_total)

@app.route('/resetar_banco', methods=['POST'])
@login_required
def resetar_banco():
    try:
        db.session.query(Historico).delete()
        db.session.query(MovimentacaoRA).delete()
        db.session.query(Material).delete()
        db.session.commit()
        flash('🧹 Banco de dados limpo com sucesso! Importe a planilha novamente.', 'warning')
    except Exception as e:
        db.session.rollback()
        flash(f'Erro ao limpar: {e}', 'danger')
    return redirect(url_for('dashboard'))

@app.route('/cadastro', methods=['GET', 'POST'])
@login_required
def cadastro():
    if request.method == 'POST':
        nome_material = request.form['nome_material'].upper()
        materiais_controlados = ['PISTOLA', 'COLETE', 'ALGEMA']
        e_controlado = any(item in nome_material for item in materiais_controlados)
        pm_re = request.form.get('pm_re')
        if e_controlado and not pm_re:
            flash('⚠️ ERRO: Armamentos exigem RE.', 'danger')
            return redirect(url_for('cadastro'))

        check_apreensao = request.form.get('check_apreensao')
        data_apreensao_obj = None
        if check_apreensao and request.form.get('data_apreensao'):
            data_apreensao_obj = datetime.strptime(request.form['data_apreensao'], '%Y-%m-%d')
            
        valor_corrigido = limpar_valor_monetario(request.form['valor'])

        patrimonio = (request.form.get('patrimonio') or '').strip()

        if not patrimonio:
            flash('Informe o número de patrimônio.', 'danger')
            return redirect(url_for('cadastro'))

        # ✅ BLOQUEIO DE DUPLICIDADE (antes do commit)
        if Material.query.filter_by(patrimonio=patrimonio).first():
            flash('Já existe um material cadastrado com este número de patrimônio.', 'warning')
            return redirect(url_for('cadastro'))

        novo_item = Material(
            patrimonio=request.form['patrimonio'],
            n_serie=request.form['n_serie'],
            nome_material=nome_material,
            conta_codigo=request.form['conta_codigo'],
            local=request.form['local'],
            valor=valor_corrigido,
            situacao=request.form['situacao'],
            observacoes=request.form['observacoes'],
            pm_re=request.form.get('pm_re'),
            pm_nome=request.form.get('pm_nome'),
            pm_opm=request.form.get('pm_opm'),
            eh_apreensao=True if check_apreensao else False,
            local_apreensao=request.form.get('local_apreensao'),
            autoridade_resp=request.form.get('autoridade_resp'),
            num_doc_apreensao=request.form.get('num_doc_apreensao'),
            data_apreensao=data_apreensao_obj
        )
        
        try:
            db.session.add(novo_item)
            db.session.commit()
            registrar_log(novo_item.id, "Criação", "Item cadastrado.")
            db.session.commit()
            flash('Cadastrado com sucesso!', 'success')
            return redirect(url_for('cadastro'))
        except Exception as e:
            db.session.rollback()
            flash(f'Erro: {e}', 'danger')

    return render_template('cadastro.html', contas=CONTAS_PATRIMONIAIS)

@app.route('/importar', methods=['POST'])
@login_required
def importar_lcm():
    file = request.files.get('arquivo_excel')
    if not file: return redirect(url_for('dashboard'))
    try:
        df = pd.read_excel(file)
        df.columns = [c.strip() for c in df.columns]
        sucesso = 0
        for index, row in df.iterrows():
            patrimonio = str(row.get('Nº Patrimônio', '')).strip()
            if not patrimonio or patrimonio == 'nan': continue
            if Material.query.filter_by(patrimonio=patrimonio).first(): continue
            
            val = limpar_valor_monetario(row.get('Valor', 0))
                
            db.session.add(Material(
                patrimonio=patrimonio,
                n_serie=str(row.get('Nº Série', '')).replace('nan', ''),
                nome_material=str(row.get('Descrição', 'IMP')).replace('nan', ''),
                conta_codigo='123119999', local='TRIAGEM', situacao='Bom',
                valor=val
            ))
            sucesso += 1
            
        db.session.commit()
        flash(f'Importação concluída! {sucesso} itens adicionados.', 'success')
        
    except Exception as e:
        db.session.rollback()
        flash(f'Erro na importação: {str(e)}', 'danger')
        
    return redirect(url_for('dashboard'))

@app.route('/editar/<int:id>', methods=['GET', 'POST'])
@login_required
def editar(id):
    material = Material.query.get_or_404(id)
    if request.method == 'POST':
        local_antigo = material.local
        situacao_antiga = material.situacao
        obs_antiga = material.observacoes
        apreensao_antiga = material.eh_apreensao
        
        material.local = request.form['local']
        material.situacao = request.form['situacao']
        material.observacoes = request.form['observacoes']
        
        check_apreensao = request.form.get('check_apreensao')
        if check_apreensao:
            material.eh_apreensao = True
            material.local_apreensao = request.form['local_apreensao']
            material.autoridade_resp = request.form['autoridade_resp']
            material.num_doc_apreensao = request.form['num_doc_apreensao']
            if request.form.get('data_apreensao'):
                material.data_apreensao = datetime.strptime(request.form['data_apreensao'], '%Y-%m-%d')
        else:
            material.eh_apreensao = False

        if local_antigo != material.local:
            registrar_log(material.id, "Movimentação Física", f"De '{local_antigo}' para '{material.local}'")
        if situacao_antiga != material.situacao:
            registrar_log(material.id, "Alteração de Status", f"De '{situacao_antiga}' para '{material.situacao}'")
        if obs_antiga != material.observacoes:
            registrar_log(material.id, "Nota", "Observações editadas")
        if not apreensao_antiga and material.eh_apreensao:
            registrar_log(material.id, "Início Custódia", "Marcado como apreensão")

        db.session.commit()
        flash('Atualizado!', 'success')
        return redirect(url_for('dashboard'))
    return render_template('editar.html', material=material, contas=CONTAS_PATRIMONIAIS)

@app.route('/pesquisa', methods=['GET'])
@login_required
def pesquisa():
    termo = request.args.get('q')
    resultados = []
    if termo:
        termo_like = f"%{termo}%"
        resultados = Material.query.filter(
            (Material.patrimonio.like(termo_like)) | 
            (Material.n_serie.like(termo_like)) | 
            (Material.pm_re.like(termo_like)) |
            (Material.nome_material.like(termo_like))
        ).all()
    return render_template('pesquisa.html', resultados=resultados, termo=termo)

@app.route('/historico_busca', methods=['GET'])
@login_required
def historico_busca():
    termo = request.args.get('q')
    resultados = []
    if termo:
        termo_like = f"%{termo}%"
        resultados = Material.query.filter(
            (Material.patrimonio.like(termo_like)) | 
            (Material.n_serie.like(termo_like)) | 
            (Material.pm_re.like(termo_like)) |
            (Material.nome_material.like(termo_like))
        ).all()
    return render_template('historico_busca.html', resultados=resultados, termo=termo)

@app.route('/historico_detalhes/<int:id>')
@login_required
def historico_detalhes(id):
    material = Material.query.get_or_404(id)
    historico = Historico.query.filter_by(material_id=id).order_by(Historico.data_hora.desc()).all()
    return render_template('historico_detalhes.html', material=material, historico=historico)

@app.route('/excluir/<int:id>', methods=['POST'])
@login_required
def excluir(id):
    material = Material.query.get_or_404(id)
    try:
        db.session.delete(material)
        db.session.commit()
        flash('Material excluído.', 'success')
    except:
        db.session.rollback()
        flash('Erro ao excluir.', 'danger')
    return redirect(url_for('dashboard'))

@app.route('/exportar_excel')
@login_required
def exportar_excel():
    materiais = Material.query.all()
    dados = []
    for m in materiais:
        dados.append({
            'Patrimônio': m.patrimonio,
            'Descrição': m.nome_material,
            'Valor': m.valor,
            'Local': m.local,
            'Situação': m.situacao,
            'Apreensão?': 'SIM' if m.eh_apreensao else 'NÃO',
            'Dias Custódia': m.dias_custodia if m.eh_apreensao else 0,
            'Responsável': m.pm_re or ''
        })
    df = pd.DataFrame(dados)
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Inventario')
    output.seek(0)
    return send_file(output, download_name='Inventario_Batalhao.xlsx', as_attachment=True, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

# Novas rotas: Usuários e RA
@app.route('/usuarios/novo', methods=['GET', 'POST'])
@login_required
def usuario_novo():
    if current_user.role != 'admin':
        flash('Acesso restrito ao administrador.', 'danger')
        return redirect(url_for('dashboard'))

    if request.method == 'POST':
        re_input = request.form.get('re', '')
        senha = request.form.get('senha', '')
        senha2 = request.form.get('senha2', '')
        role = request.form.get('role', 'usuario')

        re_login = normalizar_re(re_input)

        if not re_login or len(re_login) < 4:
            flash('RE inválido. Informe o RE com dígito para o sistema remover automaticamente.', 'danger')
            return redirect(url_for('usuario_novo'))

        if senha != senha2:
            flash('As senhas não conferem.', 'danger')
            return redirect(url_for('usuario_novo'))

        if not senha_forte(senha):
            flash('Senha fraca. Use mínimo 8 caracteres, 1 letra maiúscula, 1 número e 1 caractere especial.', 'danger')
            return redirect(url_for('usuario_novo'))

        if Usuario.query.filter_by(username=re_login).first():
            flash('Já existe usuário com esse RE (login).', 'danger')
            return redirect(url_for('usuario_novo'))

        novo = Usuario(username=re_login, role=role)
        novo.set_password(senha)
        db.session.add(novo)
        db.session.commit()

        flash('Usuário criado com sucesso!', 'success')
        return redirect(url_for('dashboard'))

    return render_template('usuario_novo.html')

@app.route('/usuarios/listar', methods=['GET'])
@login_required
def usuarios_listar():
    if current_user.role != 'admin':
        flash('Acesso restrito ao administrador.', 'danger')
        return redirect(url_for('dashboard'))
    
    usuarios = Usuario.query.all()
    return render_template('usuarios_listar.html', usuarios=usuarios)

@app.route('/usuarios/editar/<int:id>', methods=['GET', 'POST'])
@login_required
def usuarios_editar(id):
    if current_user.role != 'admin':
        flash('Acesso restrito ao administrador.', 'danger')
        return redirect(url_for('dashboard'))
    
    usuario = Usuario.query.get_or_404(id)
    
    if request.method == 'POST':
        novo_role = request.form.get('role', 'usuario')
        
        usuario.role = novo_role
        db.session.commit()
        
        flash(f'Usuário {usuario.username} atualizado com sucesso!', 'success')
        return redirect(url_for('usuarios_listar'))
    
    return render_template('usuarios_editar.html', usuario=usuario)

@app.route('/usuarios/excluir/<int:id>', methods=['POST'])
@login_required
def usuarios_excluir(id):
    if current_user.role != 'admin':
        flash('Acesso restrito ao administrador.', 'danger')
        return redirect(url_for('dashboard'))
    
    usuario = Usuario.query.get_or_404(id)
    
    # Impede excluir o próprio usuário
    if usuario.id == current_user.id:
        flash('Você não pode excluir sua própria conta.', 'warning')
        return redirect(url_for('usuarios_listar'))
    
    username = usuario.username
    db.session.delete(usuario)
    db.session.commit()
    
    flash(f'Usuário {username} excluído com sucesso.', 'success')
    return redirect(url_for('usuarios_listar'))

@app.route('/ra/retirada', methods=['GET', 'POST'])
@login_required
def ra_retirada():
    if request.method == 'POST':
        busca = (request.form.get('busca_material') or '').strip()
        posto_grad = (request.form.get('posto_grad') or '').strip()
        pm_re = (request.form.get('pm_re') or '').strip()
        nome_guerra = (request.form.get('nome_guerra') or '').strip()
        cia_secao = (request.form.get('cia_secao') or '').strip()
        obs = (request.form.get('observacoes') or '').strip()

        if not busca:
            flash('Informe Patrimônio ou Nº de Série.', 'danger')
            return redirect(url_for('ra_retirada'))

        material = Material.query.filter(
            or_(Material.patrimonio == busca, Material.n_serie == busca)
        ).first()

        if not material:
            flash('Material não encontrado (patrimônio ou série).', 'danger')
            return redirect(url_for('ra_retirada'))
        
                # ✅ BLOQUEIO DE DUPLICIDADE: já existe retirada em aberto?
        retirada_em_aberto = MovimentacaoRA.query.filter_by(
            material_id=material.id,
            data_hora_devolucao=None
        ).first()

        if retirada_em_aberto:
            flash('Este material já possui uma retirada em aberto. Registre a devolução antes de uma nova retirada.', 'warning')
            return redirect(url_for('ra_consulta'))

        if not (posto_grad and pm_re and nome_guerra and cia_secao):
            flash('Preencha Posto/Grad, RE, Nome de Guerra e Cia/Seção.', 'danger')
            return redirect(url_for('ra_retirada'))

        mov = MovimentacaoRA(
            material_id=material.id,
            operador_retirada=current_user.username,
            pm_retirou_posto_grad=posto_grad,
            pm_retirou_re=normalizar_re(pm_re),
            pm_retirou_nome_guerra=nome_guerra,
            pm_retirou_cia_secao=cia_secao,
            observacoes_retirada=obs
        )
        db.session.add(mov)

        registrar_log(material.id, "Saída RA", f"Retirado por {posto_grad} {nome_guerra} (RE {normalizar_re(pm_re)}) | Operador: {current_user.username}")

        db.session.commit()
        flash('Retirada registrada com sucesso.', 'success')
        return redirect(url_for('ra_consulta'))

    return render_template('ra_retirada.html')

@app.route('/ra/devolucao', methods=['GET', 'POST'])
@login_required
def ra_devolucao():
    if request.method == 'POST':
        busca = (request.form.get('busca_material') or '').strip()
        posto_grad = (request.form.get('posto_grad') or '').strip()
        pm_re = (request.form.get('pm_re') or '').strip()
        nome_guerra = (request.form.get('nome_guerra') or '').strip()
        cia_secao = (request.form.get('cia_secao') or '').strip()
        obs = (request.form.get('observacoes') or '').strip()

        if not busca:
            flash('Informe Patrimônio ou Nº de Série.', 'danger')
            return redirect(url_for('ra_devolucao'))

        material = Material.query.filter(
            or_(Material.patrimonio == busca, Material.n_serie == busca)
        ).first()

        if not material:
            flash('Material não encontrado (patrimônio ou série).', 'danger')
            return redirect(url_for('ra_devolucao'))

        movs_abertas = MovimentacaoRA.query.filter_by(
            material_id=material.id,
            data_hora_devolucao=None
        ).order_by(MovimentacaoRA.data_hora_retirada.desc()).all()

        if not movs_abertas:
            flash('Não existe retirada em aberto para este material (ou já foi devolvido).', 'danger')
            return redirect(url_for('ra_devolucao'))

        if len(movs_abertas) > 1:
            flash('Atenção: existem múltiplas retiradas em aberto para este material (duplicidade). Procure o administrador para corrigir.', 'danger')
            return redirect(url_for('ra_consulta'))

        mov = movs_abertas[0]

        if not (posto_grad and pm_re and nome_guerra and cia_secao):
            flash('Preencha Posto/Grad, RE, Nome de Guerra e Cia/Seção do PM que devolveu.', 'danger')
            return redirect(url_for('ra_devolucao'))

        mov.pm_devolveu_posto_grad = posto_grad
        mov.pm_devolveu_re = normalizar_re(pm_re)
        mov.pm_devolveu_nome_guerra = nome_guerra
        mov.pm_devolveu_cia_secao = cia_secao
        mov.operador_devolucao = current_user.username
        mov.observacoes_devolucao = obs
        mov.data_hora_devolucao = datetime.now()

        registrar_log(material.id, "Entrada RA", f"Devolvido por {posto_grad} {nome_guerra} (RE {normalizar_re(pm_re)}) | Operador: {current_user.username}")

        db.session.commit()
        flash('Devolução registrada com sucesso.', 'success')
        return redirect(url_for('ra_consulta'))

    return render_template('ra_devolucao.html')

@app.route('/ra/consulta', methods=['GET'])
@login_required
def ra_consulta():
    termo = (request.args.get('q') or '').strip()
    query = MovimentacaoRA.query.join(Material, MovimentacaoRA.material_id == Material.id)

    if termo:
        like = f"%{termo}%"
        query = query.filter(
            or_(
                Material.patrimonio.like(like),
                Material.n_serie.like(like),
                Material.nome_material.like(like),
                MovimentacaoRA.pm_retirou_re.like(like),
                MovimentacaoRA.pm_retirou_nome_guerra.like(like),
                MovimentacaoRA.pm_devolveu_re.like(like),
                MovimentacaoRA.pm_devolveu_nome_guerra.like(like),
            )
        )

    registros = query.order_by(MovimentacaoRA.data_hora_retirada.desc()).limit(200).all()
    return render_template('ra_consulta.html', registros=registros, termo=termo)

if __name__ == '__main__':
    with app.app_context():
        db.create_all()
        # Se tiver criação de usuário admin no protocolo, fica aqui dentro também
            
    # MODO DESENVOLVIMENTO (Comentado)
    # app.run(debug=True) 

    # MODO PRODUÇÃO / SERVIDOR DE REDE (Porta 8080)
    from waitress import serve
    print("🚀 Servidor de Protocolos rodando! Aguardando conexões na porta 8084...")
    serve(app, host='0.0.0.0', port=8084)

'''if __name__ == '__main__':
    criar_admin()
    app.run(host='0.0.0.0', port=5001, debug=False)'''