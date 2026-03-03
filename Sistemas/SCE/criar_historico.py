import sqlite3
import os

db_path = 'frequencia.db' # Ou 'instance/frequencia.db' se estiver lá

print(f"Atualizando banco em: {db_path}")
conn = sqlite3.connect(db_path)
cursor = conn.cursor()

# Cria a tabela de histórico se não existir
cursor.execute('''
    CREATE TABLE IF NOT EXISTS historico_movimentacao (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        policial_id INTEGER NOT NULL,
        data_mudanca DATETIME DEFAULT CURRENT_TIMESTAMP,
        tipo_mudanca VARCHAR(50),
        anterior VARCHAR(100),
        novo VARCHAR(100),
        usuario_responsavel VARCHAR(100),
        FOREIGN KEY(policial_id) REFERENCES policial(id)
    )
''')

conn.commit()
conn.close()
print("✅ Tabela 'HistoricoMovimentacao' criada com sucesso!")