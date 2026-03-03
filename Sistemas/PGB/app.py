from flask import Flask, render_template, request, redirect, url_for, flash
from flask_sqlalchemy import SQLAlchemy
from flask_login import LoginManager, UserMixin, login_user, login_required, logout_user, current_user
from werkzeug.security import generate_password_hash, check_password_hash
from datetime import datetime
from sqlalchemy import desc
import os

app = Flask(__name__)
app.secret_key = 'chave_secreta_protocolo_bpm'
basedir = os.path.abspath(os.path.dirname(__file__))
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///' + os.path.join(basedir, 'protocolo.db')
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

db = SQLAlchemy(app)
login_manager = LoginManager(app)
login_manager.login_view = 'login'

# --- MODELOS DE BANCO DE DADOS ---

class Usuario(UserMixin, db.Model):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(50), unique=True, nullable=False)
    password = db.Column(db.String(100), nullable=False)
    secao = db.Column(db.String(20)) # Ex: P1, P2, 1ª Cia...
    
    # NOVOS CAMPOS:
    posto = db.Column(db.String(20))
    nome = db.Column(db.String(100))
    re = db.Column(db.String(20))

class DocumentoProtocolo(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    numero_sequencial = db.Column(db.Integer)
    ano = db.Column(db.Integer)
    numero_protocolo_formatado = db.Column(db.String(20), unique=True)
    tipo_documento = db.Column(db.String(50))
    numero_documento_origem = db.Column(db.String(50))
    
    # Entrada
    entregue_posto = db.Column(db.String(20))
    entregue_re = db.Column(db.String(20))
    entregue_nome = db.Column(db.String(100))
    entregue_cia = db.Column(db.String(50))
    recebido_posto = db.Column(db.String(20))
    recebido_re = db.Column(db.String(20))
    recebido_nome = db.Column(db.String(100))
    recebido_secao = db.Column(db.String(50)) # Nova coluna: Qual seção recebeu
    data_entrada = db.Column(db.DateTime, default=datetime.now)
    
    # Saída
    data_saida = db.Column(db.DateTime)
    retirado_posto = db.Column(db.String(20))
    retirado_re = db.Column(db.String(20))
    retirado_nome = db.Column(db.String(100))
    retirado_cia = db.Column(db.String(50))
    despachante_posto = db.Column(db.String(20))
    despachante_re = db.Column(db.String(20))
    despachante_nome = db.Column(db.String(100))
    
    # Arquivo
    data_arquivamento = db.Column(db.DateTime)
    local_arquivo = db.Column(db.String(100))
    
    status = db.Column(db.String(20), default='Aberto') # Aberto, Saiu, Arquivado
    observacao = db.Column(db.Text)

@login_manager.user_loader
def load_user(user_id):
    return db.session.get(Usuario, int(user_id))

# --- ROTAS DE AUTENTICAÇÃO ---

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        user = Usuario.query.filter_by(username=request.form.get('username')).first()
        if user and user.password == request.form.get('password'):
            login_user(user)
            return redirect(url_for('index'))
        flash('Usuário ou senha incorretos.', 'danger')
    return render_template('login.html')

@app.route('/logout')
@login_required
def logout():
    logout_user()
    return redirect(url_for('login'))

@app.route('/alterar_senha', methods=['POST'])
@login_required
def alterar_senha():
    senha_atual = request.form.get('senha_atual')
    nova_senha = request.form.get('nova_senha')
    confirmar_senha = request.form.get('confirmar_senha')
    
    # 1. Verifica se as novas senhas batem
    if nova_senha != confirmar_senha:
        flash('⚠️ Erro: As novas senhas não coincidem.', 'danger')
        return redirect(request.referrer)
        
    # 2. Verifica se a senha atual digitada está correta (Texto puro)
    if current_user.password != senha_atual:
        flash('⚠️ Erro: A senha atual está incorreta.', 'danger')
        return redirect(request.referrer)
        
    # 3. Salva a nova senha
    current_user.password = nova_senha
    db.session.commit()
    
    flash('✅ Senha alterada com sucesso!', 'success')
    return redirect(request.referrer)

# --- ROTAS DO PROTOCOLO ---

@app.route('/')
@login_required
def index():
    # Pega o termo de busca digitado pelo usuário (se houver)
    busca = request.args.get('busca', '')
    
    query = DocumentoProtocolo.query
    
    # Se o usuário digitou algo, filtra o banco de dados
    if busca:
        query = query.filter(db.or_(
            DocumentoProtocolo.numero_protocolo_formatado.ilike(f'%{busca}%'),
            DocumentoProtocolo.numero_documento_origem.ilike(f'%{busca}%')
        ))
        
    # Ordena do mais recente para o mais antigo
    docs = query.order_by(desc(DocumentoProtocolo.id)).all()
    
    # Passamos a variável 'busca' para a tela para manter o texto na barra após pesquisar
    return render_template('index.html', docs=docs, user=current_user, busca=busca)

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
            tipo_documento=request.form.get('tipo'), numero_documento_origem=request.form.get('numero_origem'),
            entregue_posto=request.form.get('ent_posto'), entregue_re=request.form.get('ent_re'),
            entregue_nome=request.form.get('ent_nome'), entregue_cia=request.form.get('ent_cia'),
            recebido_posto=request.form.get('rec_posto'), recebido_re=request.form.get('rec_re'),
            recebido_nome=request.form.get('rec_nome'), recebido_secao=current_user.secao,
            observacao=request.form.get('observacao')
        )
        db.session.add(novo_doc); db.session.commit()
        flash(f'Protocolo {protocolo_fmt} registrado na {current_user.secao}!', 'success')
    except Exception as e: flash(f'Erro: {e}', 'danger')
    return redirect(url_for('index'))

@app.route('/nova_saida_direta', methods=['POST'])
@login_required
def nova_saida_direta():
    try:
        ano_atual = datetime.now().year
        ultimo = DocumentoProtocolo.query.filter_by(ano=ano_atual).order_by(desc(DocumentoProtocolo.numero_sequencial)).first()
        novo_seq = (ultimo.numero_sequencial + 1) if ultimo else 1
        protocolo_fmt = f"{ano_atual}/{novo_seq:04d}"
        
        novo_doc = DocumentoProtocolo(
            ano=ano_atual, numero_sequencial=novo_seq, numero_protocolo_formatado=protocolo_fmt,
            tipo_documento=request.form.get('tipo'), numero_documento_origem=request.form.get('numero_origem'),
            data_entrada=datetime.now(), data_saida=datetime.now(), status='Saiu',
            retirado_posto=request.form.get('ret_posto'), retirado_re=request.form.get('ret_re'),
            retirado_nome=request.form.get('ret_nome'), retirado_cia=request.form.get('ret_cia'),
            despachante_posto=request.form.get('desp_posto'), despachante_re=request.form.get('desp_re'),
            despachante_nome=request.form.get('desp_nome'), recebido_secao=current_user.secao,
            observacao=f"[{current_user.secao} - Expediente Próprio]: {request.form.get('observacao')}"
        )
        db.session.add(novo_doc); db.session.commit()
        flash(f'Saída Direta {protocolo_fmt} expedida pela {current_user.secao}!', 'warning')
    except Exception as e: flash(f'Erro: {e}', 'danger')
    return redirect(url_for('index'))

@app.route('/registrar_saida', methods=['POST'])
@login_required
def registrar_saida():
    try:
        doc = db.session.get(DocumentoProtocolo, request.form.get('doc_id'))
        if doc:
            doc.data_saida = datetime.now()
            doc.status = 'Saiu'
            doc.retirado_posto = request.form.get('ret_posto')
            doc.retirado_re = request.form.get('ret_re')
            doc.retirado_nome = request.form.get('ret_nome')
            doc.retirado_cia = request.form.get('ret_cia')
            doc.despachante_posto = request.form.get('desp_posto')
            doc.despachante_re = request.form.get('desp_re')
            doc.despachante_nome = request.form.get('desp_nome')
            
            obs_extra = request.form.get('obs_saida')
            if obs_extra: 
                doc.observacao = (doc.observacao or '') + f"\n[Saída da {current_user.secao}]: {obs_extra}"
            
            db.session.commit()
            flash('Trâmite de saída registrado com sucesso!', 'success')
    except Exception as e: 
        flash(f'Erro: {e}', 'danger')
    return redirect(url_for('index'))

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
            if obs_arq: 
                doc.observacao = (doc.observacao or '') + f"\n[Arquivo {current_user.secao}]: {obs_arq}"
            
            db.session.commit()
            flash('Documento arquivado.', 'secondary')
    except Exception as e: 
        flash(f'Erro: {e}', 'danger')
    return redirect(url_for('index'))

@app.route('/gerar_relatorio_protocolo', methods=['POST', 'GET'])
@login_required
def gerar_relatorio_protocolo():
    try:
        # Pega o filtro se existir no seu modal (ex: Todos, Abertos, Arquivados)
        status_filtro = request.form.get('status_filtro', 'TODOS')
        
        query = DocumentoProtocolo.query
        
        if status_filtro != 'TODOS':
            query = query.filter_by(status=status_filtro)
            
        # Ordena do mais novo para o mais antigo
        docs = query.order_by(desc(DocumentoProtocolo.id)).all()
        
        # Renderiza uma página limpa pronta para impressão
        return render_template('relatorio_protocolo.html', docs=docs, filtro=status_filtro, hoje=datetime.now())
        
    except Exception as e:
        flash(f'Erro ao gerar relatório: {e}', 'danger')
        return redirect(url_for('index'))
    
# --- ROTAS DE GERENCIAMENTO DE USUÁRIOS (APENAS ADMIN) ---

@app.route('/usuarios')
@login_required
def gerenciar_usuarios():
    # Bloqueia o acesso se não for o admin
    if current_user.username != 'admin':
        flash('Acesso negado. Apenas o administrador pode acessar esta página.', 'danger')
        return redirect(url_for('index'))
    
    usuarios = Usuario.query.all()
    return render_template('usuarios.html', usuarios=usuarios, user=current_user)

@app.route('/novo_usuario', methods=['POST'])
@login_required
def novo_usuario():
    if current_user.username != 'admin':
        return redirect(url_for('index'))
        
    username = request.form.get('username')
    password = request.form.get('password')
    secao = request.form.get('secao')
    posto = request.form.get('posto')
    nome = request.form.get('nome')
    re = request.form.get('re')
    
    if Usuario.query.filter_by(username=username).first():
        flash('Erro: Este nome de usuário já está em uso!', 'danger')
    else:
        novo_user = Usuario(username=username, password=password, secao=secao, posto=posto, nome=nome, re=re)
        db.session.add(novo_user)
        db.session.commit()
        flash(f'Usuário "{username}" ({secao}) criado com sucesso!', 'success')
        
    return redirect(url_for('gerenciar_usuarios'))

@app.route('/editar_usuario', methods=['POST'])
@login_required
def editar_usuario():
    if current_user.username != 'admin':
        return redirect(url_for('index'))
        
    user_id = request.form.get('user_id')
    novo_username = request.form.get('username')
    nova_secao = request.form.get('secao')
    novo_posto = request.form.get('posto')
    novo_nome = request.form.get('nome')
    novo_re = request.form.get('re')
    
    usuario = db.session.get(Usuario, user_id)
    if usuario:
        existente = Usuario.query.filter_by(username=novo_username).first()
        if existente and existente.id != usuario.id:
            flash('Erro: Este nome de usuário já está em uso por outra pessoa!', 'danger')
        else:
            usuario.username = novo_username
            usuario.secao = nova_secao
            usuario.posto = novo_posto
            usuario.nome = novo_nome
            usuario.re = novo_re
            db.session.commit()
            flash(f'Dados do usuário atualizados com sucesso!', 'success')
    else:
        flash('Usuário não encontrado.', 'danger')
        
    return redirect(url_for('gerenciar_usuarios'))

@app.route('/alterar_senhaadmin', methods=['POST'])
@login_required
def alterar_senhaadmin():
    if current_user.username != 'admin':
        return redirect(url_for('index'))
        
    user_id = request.form.get('user_id')
    nova_senha = request.form.get('nova_senha')
    
    usuario = db.session.get(Usuario, user_id)
    if usuario:
        usuario.password = nova_senha
        db.session.commit()
        flash(f'Senha do usuário "{usuario.username}" alterada com sucesso!', 'success')
    else:
        flash('Usuário não encontrado.', 'danger')
        
    return redirect(url_for('gerenciar_usuarios'))

@app.route('/excluir_usuario/<int:user_id>', methods=['POST'])
@login_required
def excluir_usuario(user_id):
    if current_user.username != 'admin':
        return redirect(url_for('index'))
        
    usuario = db.session.get(Usuario, user_id)
    
    if usuario:
        # Trava de segurança: impede a exclusão do admin principal
        if usuario.username == 'admin':
            flash('Ação bloqueada: Não é permitido excluir o administrador principal do sistema.', 'danger')
        else:
            nome_excluido = usuario.username
            db.session.delete(usuario)
            db.session.commit()
            flash(f'Usuário "{nome_excluido}" excluído com sucesso.', 'info')
    else:
        flash('Usuário não encontrado.', 'danger')
        
    return redirect(url_for('gerenciar_usuarios'))

if __name__ == '__main__':
    with app.app_context():
        db.create_all()
        
        # Cria o admin se não existir
        if not Usuario.query.filter_by(username='admin').first():
            admin = Usuario(username='admin', password='123', secao='P1', posto='', nome='Administrador', re='')
            db.session.add(admin)
            db.session.commit()
            print("Usuário Admin criado com sucesso!")
            
    # MODO PRODUÇÃO / SERVIDOR DE REDE (Porta 8080)
    from waitress import serve
    print("🚀 Servidor de Protocolos rodando! Aguardando conexões na porta 8080...")
    serve(app, host='0.0.0.0', port=8080)

'''if __name__ == '__main__':
    with app.app_context():
        db.create_all()
        # Cria um usuário admin padrão se não existir
        if not Usuario.query.first():
            admin = Usuario(username='admin', password='123', secao='Comando')
            db.session.add(admin)
            db.session.commit()
    app.run(debug=True, port=5001) # Rodando na porta 5001 para não conflitar com o sistema da P1'''