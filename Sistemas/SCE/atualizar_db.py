# Importa o seu aplicativo e o banco de dados para usar as mesmas configurações
try:
    from app import app, db
except ImportError:
    print("❌ Erro: O arquivo 'app.py' não foi encontrado na mesma pasta.")
    input("Aperte ENTER para sair...")
    exit()

from sqlalchemy import text
from datetime import datetime

def executar_atualizacao():
    with app.app_context():
        print("🔄 Conectado ao banco de dados oficial...")
        
        # 1. CRIAR A TABELA DE HISTÓRICO (Caso não exista)
        try:
            db.session.execute(text("""
                CREATE TABLE IF NOT EXISTS historico_escala (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    policial_id INTEGER NOT NULL,
                    mes INTEGER NOT NULL,
                    ano INTEGER NOT NULL,
                    tipo_escala_antigo VARCHAR(20) NOT NULL,
                    FOREIGN KEY(policial_id) REFERENCES policial(id)
                );
            """))
            db.session.commit()
            print("  ✅ Tabela 'historico_escala' verificada/criada com sucesso!")
        except Exception as e:
            print(f"  ❌ Erro ao criar tabela de histórico: {e}")

        # 2. ADICIONAR A COLUNA DATA_BASE_ESCALA NA TABELA POLICIAL
        try:
            db.session.execute(text("ALTER TABLE policial ADD COLUMN data_base_escala DATE;"))
            db.session.commit()
            print("  ✅ Coluna 'data_base_escala' adicionada com sucesso!")
        except Exception as e:
            db.session.rollback()
            if "duplicate column name" in str(e).lower():
                print("  ⚠️ A coluna 'data_base_escala' já existe. (Tudo certo)")
            else:
                print(f"  ❌ Erro ao criar a coluna: {e}")

        # 3. CONFIGURAR DATAS INICIAIS PARA O EFETIVO ANTIGO (Janeiro do ano atual)
        try:
            ano_atual = datetime.now().year
            data_pares = f"{ano_atual}-01-02"
            data_impares = f"{ano_atual}-01-01"
            
            # Atualiza quem é Pares
            db.session.execute(text(f"""
                UPDATE policial 
                SET data_base_escala = '{data_pares}' 
                WHERE inicio_pares = 1 AND data_base_escala IS NULL;
            """))
            
            # Atualiza quem é Ímpares
            db.session.execute(text(f"""
                UPDATE policial 
                SET data_base_escala = '{data_impares}' 
                WHERE (inicio_pares = 0 OR inicio_pares IS NULL) AND data_base_escala IS NULL;
            """))
            
            db.session.commit()
            print("  ✅ Escalas do efetivo atual protegidas e migradas com sucesso!")
        except Exception as e:
            db.session.rollback()
            print(f"  ❌ Erro ao configurar datas iniciais: {e}")

        print("\n🎉 BANCO DE DADOS ATUALIZADO E PRONTO PARA USO!")

if __name__ == '__main__':
    executar_atualizacao()
    input("\nFim do processo. Aperte ENTER para fechar...")
