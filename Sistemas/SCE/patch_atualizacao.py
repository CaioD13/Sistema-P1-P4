import sqlite3
import os
from pathlib import Path

# =====================================================
# 1) DESCOBRIR O CAMINHO REAL DO BANCO QUE O FLASK USA
# =====================================================

def localizar_frequencia_db():
    print("🔎 Procurando automaticament o arquivo 'frequencia.db'...")

    # Locais comuns onde o Flask pode estar carregando o banco
    caminhos_possiveis = [
        Path(os.getcwd()) / "frequencia.db",                    # pasta atual
        Path(__file__).parent / "frequencia.db",                # pasta onde o script está
        Path(os.getcwd()).parent / "frequencia.db",             # pasta acima
    ]

    # Caminhos adicionais: escaneie toda a pasta do usuário
    user_home = Path.home()
    for root, dirs, files in os.walk(user_home):
        if "frequencia.db" in files:
            caminhos_possiveis.append(Path(root) / "frequencia.db")

    encontrados = []

    for caminho in caminhos_possiveis:
        if caminho.exists():
            encontrados.append(caminho)

    if not encontrados:
        print("\n❌ Nenhum banco 'frequencia.db' encontrado no sistema.")
        print("Coloque este script na MESMA pasta onde o Flask está rodando.")
        exit()

    print("\n📌 Bancos encontrados:")
    for c in encontrados:
        print(f"   → {c}")

    # Caso tenha mais de 1 banco, escolher o provável
    escolhido = encontrados[0]
    print(f"\n✅ Usando este banco para patch: {escolhido}\n")
    return escolhido


db_path = localizar_frequencia_db()


# =====================================================
# 2) DEFINIÇÃO DAS TABELAS E COLUNAS ESPERADAS
# =====================================================

schema_updates = {
    "policial": [
        ("hora_inicio_escala", "TIME"),
        ("hora_fim_escala", "TIME"),
        ("duracao_horas_escala", "FLOAT"),
        ("status_saude", "TEXT"),
        ("restricao_tipo", "TEXT"),
        ("data_fim_restricao", "DATE"),
        ("tipo_escala", "TEXT"),
        ("inicio_pares", "BOOLEAN"),
        ("data_base_escala", "DATE"),
    ],
    "alteracao_escala": [
        ("motivo", "TEXT"),
    ],
    "validacao_mensal": [
        ("id", "INTEGER PRIMARY KEY AUTOINCREMENT"),
        ("policial_id", "INTEGER"),
        ("mes", "INTEGER"),
        ("ano", "INTEGER"),
        ("admin_id", "INTEGER"),
        ("admin_info", "TEXT"),
        ("data_validacao", "DATETIME"),
    ],
    "historico_escala": [
        ("id", "INTEGER PRIMARY KEY AUTOINCREMENT"),
        ("policial_id", "INTEGER"),
        ("mes", "INTEGER"),
        ("ano", "INTEGER"),
        ("tipo_escala_antigo", "TEXT"),
    ],
}


# =====================================================
# 3) FUNÇÕES DE UTILIDADE
# =====================================================

def tabela_existe(cursor, tabela):
    cursor.execute(f"SELECT name FROM sqlite_master WHERE type='table' AND name='{tabela}'")
    return cursor.fetchone() is not None

def coluna_existe(cursor, tabela, coluna):
    cursor.execute(f"PRAGMA table_info({tabela})")
    colunas = [info[1] for info in cursor.fetchall()]
    return coluna in colunas


# =====================================================
# 4) EXECUTAR O PATCH
# =====================================================

def aplicar_patch():
    print("🚀 Iniciando patch automático completo...\n")

    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    for tabela, colunas in schema_updates.items():

        print(f"\n📌 Verificando tabela: {tabela}")

        # ================================
        # Criar tabela se não existir
        # ================================
        if not tabela_existe(cursor, tabela):
            print(f"   ⚠️ Tabela '{tabela}' não existe → criando...")
            colunas_def = ", ".join([f"{nome} {tipo}" for nome, tipo in colunas])
            cursor.execute(f"CREATE TABLE {tabela} ({colunas_def})")
            print(f"   ✔ Tabela '{tabela}' criada.")
            conn.commit()
            continue  # Pode pular adição de colunas pois acabou de criar

        # ================================
        # Adicionar colunas faltantes
        # ================================
        for coluna, tipo in colunas:
            if coluna_existe(cursor, tabela, coluna):
                print(f"   ✔ Coluna '{coluna}' já existe.")
            else:
                try:
                    cursor.execute(f"ALTER TABLE {tabela} ADD COLUMN {coluna} {tipo}")
                    print(f"   ➕ Coluna '{coluna}' adicionada ({tipo}).")
                except Exception as e:
                    print(f"   ❌ Erro ao adicionar '{coluna}': {e}")

    print("\n🧹 Rodando VACUUM (otimização)...")
    cursor.execute("VACUUM")

    conn.commit()
    conn.close()

    print("\n🎉 PATCH FINALIZADO!")
    print("Reinicie seu servidor Flask agora.\n")


# =====================================================
# 5) EXECUÇÃO
# =====================================================

if __name__ == "__main__":
    aplicar_patch()