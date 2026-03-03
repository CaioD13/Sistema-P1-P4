import sqlite3

# Conecta ao banco de dados do protocolo
conn = sqlite3.connect('protocolo.db')
cursor = conn.cursor()

# Tenta adicionar as novas colunas (ignora se já existirem)
try: 
    cursor.execute("ALTER TABLE usuario ADD COLUMN posto VARCHAR(20)")
    print("Coluna 'posto' adicionada.")
except: pass

try: 
    cursor.execute("ALTER TABLE usuario ADD COLUMN nome VARCHAR(100)")
    print("Coluna 'nome' adicionada.")
except: pass

try: 
    cursor.execute("ALTER TABLE usuario ADD COLUMN re VARCHAR(20)")
    print("Coluna 're' adicionada.")
except: pass

conn.commit()
conn.close()
print("✅ Atualização do banco concluída com sucesso!")