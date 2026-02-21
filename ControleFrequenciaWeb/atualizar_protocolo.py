import sqlite3
import os

basedir = os.path.abspath(os.path.dirname(__file__))
db_path = os.path.join(basedir, 'frequencia.db')

conn = sqlite3.connect(db_path)
cursor = conn.cursor()

try:
    print("Criando tabela de Protocolo...")
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS documento_protocolo (
        id INTEGER PRIMARY KEY,
        numero_sequencial INTEGER,
        ano INTEGER,
        numero_protocolo_formatado VARCHAR(20) UNIQUE,
        tipo_documento VARCHAR(50),
        numero_documento_origem VARCHAR(50),
        entregue_posto VARCHAR(20),
        entregue_re VARCHAR(20),
        entregue_nome VARCHAR(100),
        entregue_cia VARCHAR(50),
        recebido_posto VARCHAR(20),
        recebido_re VARCHAR(20),
        recebido_nome VARCHAR(100),
        recebido_cia VARCHAR(50),
        data_entrada DATETIME,
        data_saida DATETIME,
        saida_destino VARCHAR(100),
        status VARCHAR(20) DEFAULT 'Aberto',
        observacao TEXT
    )
    """)
    print("✅ Tabela 'documento_protocolo' criada com sucesso!")
    conn.commit()
except Exception as e:
    print(f"❌ Erro: {e}")

conn.close()
input("Enter para sair...")