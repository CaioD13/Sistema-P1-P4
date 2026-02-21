import sqlite3

# Nome do seu arquivo de banco de dados
banco = 'frequencia.db' 
# ATENÇÃO: Se o seu banco estiver na raiz, mude para 'frequencia.db'

print(f"🔄 Conectando ao banco: {banco}...")

try:
    conn = sqlite3.connect(banco)
    cursor = conn.cursor()

    # Adiciona coluna secao_em
    try:
        print("Adicionando coluna 'secao_em'...")
        cursor.execute("ALTER TABLE policial ADD COLUMN secao_em VARCHAR(20)")
        print("✅ Coluna 'secao_em' criada!")
    except sqlite3.OperationalError:
        print("⚠️ Coluna 'secao_em' já existe.")

    # Adiciona coluna viatura
    try:
        print("Adicionando coluna 'viatura'...")
        cursor.execute("ALTER TABLE policial ADD COLUMN viatura VARCHAR(20)")
        print("✅ Coluna 'viatura' criada!")
    except sqlite3.OperationalError:
        print("⚠️ Coluna 'viatura' já existe.")

    # Adiciona coluna equipe
    try:
        print("Adicionando coluna 'equipe'...")
        cursor.execute("ALTER TABLE policial ADD COLUMN equipe VARCHAR(20)")
        print("✅ Coluna 'equipe' criada!")
    except sqlite3.OperationalError:
        print("⚠️ Coluna 'equipe' já existe.")

    conn.commit()
    conn.close()
    print("\n🚀 ATUALIZAÇÃO CONCLUÍDA! Seus dados foram preservados.")

except Exception as e:
    print(f"\n❌ Erro ao conectar: {e}")
    print("Verifique se o arquivo 'frequencia.db' está na pasta correta.")