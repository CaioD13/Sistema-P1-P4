import sqlite3
import os

DB_NAME = "frequencia.db"

# === TABELAS E COLUNAS QUE PRECISAM EXISTIR ===
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
        ("admin_info", "TEXT"),
        ("data_validacao", "DATETIME"),
    ],
    "historico_escala": [
        ("tipo_escala_antigo", "TEXT"),
    ],
    "historico_movimentacao": [
        ("usuario_responsavel", "TEXT"),
    ],
    "escala_versao": [
        ("tipo_filtro", "TEXT"),
        ("versao_atual", "INTEGER"),
    ],
    "log_auditoria": [
        ("usuario_re", "TEXT"),
        ("detalhes", "TEXT"),
    ]
}

def coluna_existe(cursor, tabela, coluna):
    cursor.execute(f"PRAGMA table_info({tabela})")
    colunas = [info[1] for info in cursor.fetchall()]
    return coluna in colunas

def aplicar_patch():
    print("\n=== Atualizador de Banco SCE - 37º BPM/M ===\n")

    if not os.path.exists(DB_NAME):
        print(f"❌ ERRO: Banco '{DB_NAME}' não encontrado!")
        return

    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()

    for tabela, colunas in schema_updates.items():
        print(f"\n📌 Verificando tabela: {tabela}")
        for coluna, tipo in colunas:
            if coluna_existe(cursor, tabela, coluna):
                print(f"   ✔ Coluna '{coluna}' já existe.")
            else:
                try:
                    cursor.execute(f"ALTER TABLE {tabela} ADD COLUMN {coluna} {tipo}")
                    print(f"   ➕ Coluna '{coluna}' adicionada ({tipo}).")
                except Exception as e:
                    print(f"   ❌ Erro ao adicionar coluna '{coluna}': {e}")

    print("\n🧹 Otimizando banco (VACUUM)...")
    try:
        cursor.execute("VACUUM")
        print("   ✔ Banco otimizado.")
    except Exception as e:
        print(f"   ❌ Erro no VACUUM: {e}")

    conn.commit()
    conn.close()

    print("\n🎉 Atualização finalizada com sucesso!")
    print("Reinicie seu servidor Flask.\n")

if __name__ == "__main__":
    aplicar_patch()