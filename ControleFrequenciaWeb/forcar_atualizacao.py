import sqlite3
import os

# 1. Garante que pegamos o arquivo na MESMA pasta deste script
basedir = os.path.abspath(os.path.dirname(__file__))
db_path = os.path.join(basedir, 'frequencia.db')

print(f"📂 Alvo do Banco de Dados: {db_path}")

if not os.path.exists(db_path):
    print("❌ ERRO: O arquivo frequencia.db não foi encontrado nesta pasta!")
    print("   Certifique-se de salvar este script na mesma pasta do app.py")
    input("Pressione Enter para sair...")
    exit()

conn = sqlite3.connect(db_path)
cursor = conn.cursor()

try:
    print("🔍 Verificando tabela 'alteracao_escala'...")
    
    # Verifica quais colunas existem
    cursor.execute("PRAGMA table_info(alteracao_escala)")
    colunas = [info[1] for info in cursor.fetchall()]
    print(f"   Colunas encontradas: {colunas}")

    if 'motivo' not in colunas:
        print("🔧 Coluna 'motivo' ausente. Adicionando agora...")
        cursor.execute("ALTER TABLE alteracao_escala ADD COLUMN motivo VARCHAR(200)")
        conn.commit()
        print("✅ SUCESSO! Coluna 'motivo' adicionada.")
    else:
        print("✅ A coluna 'motivo' JÁ EXISTE neste banco.")

except Exception as e:
    print(f"❌ Erro ao alterar tabela: {e}")
    # Se der erro que a tabela não existe, vamos criá-la
    if "no such table" in str(e):
        print("⚠️ Tabela não existe. Criando do zero...")
        cursor.execute("""
        CREATE TABLE alteracao_escala (
            id INTEGER PRIMARY KEY,
            policial_id INTEGER,
            data DATE,
            valor VARCHAR(10),
            motivo VARCHAR(200),
            FOREIGN KEY(policial_id) REFERENCES policial(id)
        )
        """)
        conn.commit()
        print("✅ Tabela criada do zero!")

conn.close()
input("\nProcesso finalizado. Pressione Enter...")