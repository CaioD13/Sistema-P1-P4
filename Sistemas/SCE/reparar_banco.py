import sqlite3
from pathlib import Path

DB_PATH = Path(r"C:\Users\33320866850\Desktop\Sistemas\SCE\frequencia.db")

if not DB_PATH.exists():
    raise FileNotFoundError(f"Banco não encontrado: {DB_PATH}")

conn = sqlite3.connect(DB_PATH)
cursor = conn.cursor()

cursor.execute("SELECT name FROM sqlite_master WHERE type='table'")
tabelas = [row[0] for row in cursor.fetchall()]
print("Tabelas encontradas:", tabelas)

if "policial" not in tabelas:
    raise Exception("A tabela 'policial' não foi encontrada nesse banco.")

cursor.execute("PRAGMA table_info(policial)")
colunas = [row[1] for row in cursor.fetchall()]
print("Colunas atuais:", colunas)

if "regime_escala" not in colunas:
    cursor.execute("ALTER TABLE policial ADD COLUMN regime_escala TEXT")
    print("Coluna regime_escala adicionada.")

if "data_base_escala" not in colunas:
    cursor.execute("ALTER TABLE policial ADD COLUMN data_base_escala DATE")
    print("Coluna data_base_escala adicionada.")

if "habilitacoes" not in colunas:
    cursor.execute("ALTER TABLE policial ADD COLUMN habilitacoes TEXT")
    print("Coluna habilitacoes adicionada.")

if "sat" not in colunas:
    cursor.execute("ALTER TABLE policial ADD COLUMN sat TEXT")
    print("Coluna sat adicionada.")


if 'hora_inicio_ee' not in colunas:
    cursor.execute("ALTER TABLE alteracao_escala ADD COLUMN hora_inicio_ee VARCHAR(5)")

if 'hora_fim_ee' not in colunas:
    cursor.execute("ALTER TABLE alteracao_escala ADD COLUMN hora_fim_ee VARCHAR(5)")


conn.commit()
conn.close()

print("Banco atualizado com sucesso.")