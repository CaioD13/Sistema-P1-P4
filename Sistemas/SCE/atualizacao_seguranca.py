from app import app, db
from sqlalchemy import text

print("🛡️ Iniciando blindagem do banco de dados...")

with app.app_context():
    # 1. Atualizar Tabela de Usuários (Adicionar RE e Perfil)
    try:
        with db.engine.connect() as conn:
            conn.execute(text("ALTER TABLE usuario ADD COLUMN re VARCHAR(20)"))
            conn.execute(text("ALTER TABLE usuario ADD COLUMN perfil VARCHAR(20) DEFAULT 'usuario'"))
            # Atualiza o admin existente para ser superusuário
            conn.execute(text("UPDATE usuario SET perfil='admin', re='000000' WHERE username='admin'"))
            conn.commit()
        print("✅ Tabela Usuário atualizada (RE e Perfil).")
    except Exception as e:
        print(f"ℹ️ Tabela Usuário já estava atualizada ou erro: {e}")

    # 2. Atualizar Tabela de Logs (Adicionar RE do autor)
    try:
        with db.engine.connect() as conn:
            conn.execute(text("ALTER TABLE log_auditoria ADD COLUMN usuario_re VARCHAR(20)"))
            conn.commit()
        print("✅ Tabela Logs atualizada (RE).")
    except Exception as e:
        print(f"ℹ️ Tabela Logs já estava atualizada ou erro: {e}")

print("🏁 Banco de dados pronto para o novo sistema de segurança!")