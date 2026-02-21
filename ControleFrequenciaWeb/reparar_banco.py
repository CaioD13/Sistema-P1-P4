import sqlite3
import os
from werkzeug.security import generate_password_hash

# 1. Ajuste aqui se seu banco estiver em outra pasta (ex: 'instance/frequencia.db')
CAMINHO_BANCO = 'frequencia.db'

print(f"🔍 Verificando banco de dados: {os.path.abspath(CAMINHO_BANCO)}")

if not os.path.exists(CAMINHO_BANCO):
    print("❌ ERRO CRÍTICO: O arquivo do banco não existe neste local!")
    print("Certifique-se de que copiou seu banco antigo para cá.")
    exit()

conn = sqlite3.connect(CAMINHO_BANCO)
cursor = conn.cursor()

# 2. Criar tabela USUARIO se não existir
print("🛠️ Verificando tabela 'usuario'...")
cursor.execute('''
    CREATE TABLE IF NOT EXISTS usuario (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        username VARCHAR(80) UNIQUE NOT NULL,
        password VARCHAR(120) NOT NULL,
        nome VARCHAR(100),
        re VARCHAR(20),
        perfil VARCHAR(20) DEFAULT 'usuario'
    )
''')

# 3. Criar tabela LOG se não existir (apenas para garantir)
cursor.execute('''
    CREATE TABLE IF NOT EXISTS log_auditoria (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        data_hora DATETIME DEFAULT CURRENT_TIMESTAMP,
        usuario_id INTEGER,
        usuario_nome VARCHAR(100),
        usuario_re VARCHAR(20),
        acao VARCHAR(50),
        detalhes VARCHAR(200)
    )
''')

# 4. Criar Admin se não existir
cursor.execute("SELECT * FROM usuario WHERE username = 'admin'")
if not cursor.fetchone():
    print("👤 Criando usuário admin padrão...")
    senha_hash = generate_password_hash('123456')
    cursor.execute(
        "INSERT INTO usuario (username, password, nome, re, perfil) VALUES (?, ?, ?, ?, ?)",
        ('admin', senha_hash, 'Administrador', 'admin', 'admin')
    )
    print("✅ Admin criado! (Login: admin / Senha: 123456)")
else:
    print("✅ Usuário admin já existe.")

conn.commit()
conn.close()

print("\n🚀 Banco reparado com sucesso!")
print("Tente rodar o app.py novamente.")