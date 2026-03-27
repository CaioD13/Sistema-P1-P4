import sqlite3

db = sqlite3.connect("frequencia.db")
cursor = db.cursor()

print("🔧 Criando tabelas faltantes...\n")

# 1. Tabela VALIDACAO_MENSAL
cursor.execute("""
CREATE TABLE IF NOT EXISTS validacao_mensal (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    policial_id INTEGER,
    mes INTEGER,
    ano INTEGER,
    admin_id INTEGER,
    admin_info TEXT,
    data_validacao DATETIME
)
""")
print("✔ Tabela validacao_mensal verificada/criada.")

# 2. Tabela HISTORICO_ESCALA
cursor.execute("""
CREATE TABLE IF NOT EXISTS historico_escala (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    policial_id INTEGER,
    mes INTEGER,
    ano INTEGER,
    tipo_escala_antigo TEXT
)
""")
print("✔ Tabela historico_escala verificada/criada.")

db.commit()
db.close()

print("\n🎉 Tudo pronto! Reinicie seu app.py.")