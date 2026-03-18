from app import app, db, Usuario
from werkzeug.security import generate_password_hash

with app.app_context():
    # Cria as tabelas caso não existam
    db.create_all()
    
    # Verifica se já existe algum usuário
    if Usuario.query.filter_by(username='admin').first():
        print("❌ O usuário 'admin' já existe!")
    else:
        # CRIA O ADMINISTRADOR PADRÃO
        senha_hash = generate_password_hash('123456') # <--- SENHA INICIAL
        
        admin = Usuario(
            username='admin',  # Este será o RE/Login
            re='admin',
            nome='Administrador do Sistema',
            password=senha_hash,
            perfil='admin'
        )
        
        db.session.add(admin)
        db.session.commit()
        
        print("\n✅ SUCESSO! Usuário criado.")
        print("👤 Login (RE): admin")
        print("🔑 Senha: 123456")