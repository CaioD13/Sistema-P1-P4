import sqlite3
import os

basedir = os.path.abspath(os.path.dirname(__file__))
db_path = os.path.join(basedir, 'frequencia.db')
conn = sqlite3.connect(db_path)
cursor = conn.cursor()

try:
    print("🗑️ Apagando tabela antiga de protocolo (se existir)...")
    cursor.execute("DROP TABLE IF EXISTS documento_protocolo")
    
    print("🆕 Criando nova tabela completa...")
    cursor.execute("""
    CREATE TABLE documento_protocolo (
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
        retirado_posto VARCHAR(20),
        retirado_re VARCHAR(20),
        retirado_nome VARCHAR(100),
        retirado_cia VARCHAR(50),
        
        despachante_posto VARCHAR(20),
        despachante_re VARCHAR(20),
        despachante_nome VARCHAR(100),
        
        data_arquivamento DATETIME,
        local_arquivo VARCHAR(100),
        
        status VARCHAR(20) DEFAULT 'Aberto',
        observacao TEXT
    )
    """)
    print("✅ Tabela recriada com sucesso!")
    conn.commit()
except Exception as e:
    print(f"❌ Erro: {e}")

conn.close()
input("Enter para sair...")