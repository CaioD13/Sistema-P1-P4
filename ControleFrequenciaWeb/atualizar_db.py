import sqlite3

# Conecta no banco (ou cria se não existir)
conn = sqlite3.connect('frequencia.db')
cursor = conn.cursor()

print("🔍 Iniciando reparo do banco de dados...")

# 1. Tabela USUARIO
try:
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS usuario (
        id INTEGER PRIMARY KEY,
        username VARCHAR(20) UNIQUE NOT NULL,
        re VARCHAR(20) UNIQUE NOT NULL,
        nome VARCHAR(100) NOT NULL,
        password VARCHAR(100) NOT NULL,
        perfil VARCHAR(20) DEFAULT 'comum'
    )
    """)
    print("✅ Tabela 'usuario' verificada.")
except Exception as e:
    print(f"❌ Erro em usuario: {e}")

# 2. Tabela POLICIAL
try:
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS policial (
        id INTEGER PRIMARY KEY,
        posto VARCHAR(20),
        re VARCHAR(20) UNIQUE,
        nome VARCHAR(100),
        nome_guerra VARCHAR(50),
        cia VARCHAR(20),
        secao_em VARCHAR(50),
        tipo_escala VARCHAR(20),
        inicio_pares BOOLEAN DEFAULT 1,
        funcao VARCHAR(100),
        equipe VARCHAR(50),
        viatura VARCHAR(50),
        ativo BOOLEAN DEFAULT 1,
        status_saude VARCHAR(50) DEFAULT 'Apto Sem Restrições',
        restricao_tipo VARCHAR(100),
        data_fim_restricao DATE
    )
    """)
    print("✅ Tabela 'policial' verificada.")
except Exception as e:
    print(f"❌ Erro em policial: {e}")

# 3. Tabela ALTERACAO_ESCALA (A que estava faltando!)
try:
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS alteracao_escala (
        id INTEGER PRIMARY KEY,
        policial_id INTEGER,
        data DATE,
        valor VARCHAR(10),
        motivo VARCHAR(200),
        FOREIGN KEY(policial_id) REFERENCES policial(id)
    )
    """)
    print("✅ Tabela 'alteracao_escala' verificada/criada.")
    
    # Se ela já existia mas sem 'motivo', tenta adicionar
    try:
        cursor.execute("ALTER TABLE alteracao_escala ADD COLUMN motivo VARCHAR(200)")
        print("   -> Coluna 'motivo' adicionada.")
    except:
        pass # Coluna já existe ou tabela acabou de ser criada
        
except Exception as e:
    print(f"❌ Erro em alteracao_escala: {e}")

# 4. Tabela ESCALA_VERSAO
try:
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS escala_versao (
        id INTEGER PRIMARY KEY,
        cia VARCHAR(50),
        data_inicio DATE,
        tipo_filtro VARCHAR(20),
        versao_atual INTEGER DEFAULT 1
    )
    """)
    print("✅ Tabela 'escala_versao' verificada.")
except Exception as e:
    print(f"❌ Erro em escala_versao: {e}")

# 5. Tabelas de Log (Opcional, mas bom ter)
try:
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS log_auditoria (
        id INTEGER PRIMARY KEY,
        data_hora DATETIME,
        usuario_id INTEGER,
        usuario_nome VARCHAR(100),
        usuario_re VARCHAR(20),
        acao VARCHAR(50),
        detalhes VARCHAR(500)
    )
    """)
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS historico_movimentacao (
        id INTEGER PRIMARY KEY,
        policial_id INTEGER,
        data_mudanca DATETIME,
        tipo_mudanca VARCHAR(50),
        anterior VARCHAR(100),
        novo VARCHAR(100),
        usuario_responsavel VARCHAR(100),
        FOREIGN KEY(policial_id) REFERENCES policial(id)
    )
    """)
    print("✅ Tabelas de Log verificadas.")
except Exception as e:
    print(f"❌ Erro em logs: {e}")

conn.commit()
conn.close()
print("\n🚀 Banco de dados reparado com sucesso!")
input("Pressione Enter para sair...")