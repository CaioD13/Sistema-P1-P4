# Importa o seu próprio aplicativo e o banco de dados oficial dele
from app import app, db
from sqlalchemy import text
from datetime import datetime

def atualizar_banco_seguro():
    # Entra no "contexto" do seu sistema para usar as configurações dele
    with app.app_context():
        print("🔄 Conectado ao banco de dados oficial do sistema...")
        
        # 1. TENTAR CRIAR A COLUNA
        try:
            db.session.execute(text("ALTER TABLE policial ADD COLUMN data_base_escala DATE;"))
            db.session.commit()
            print("  ✅ Coluna 'data_base_escala' criada com sucesso!")
        except Exception as e:
            db.session.rollback()
            if "duplicate column name" in str(e).lower():
                print("  ⚠️ A coluna 'data_base_escala' já existe no banco oficial. (Tudo certo)")
            else:
                print(f"  ❌ Erro ao criar a coluna: {e}")
                
        # 2. ATUALIZAR AS DATAS ANTIGAS PARA JANEIRO
        try:
            ano_atual = datetime.now().year
            data_pares = f"{ano_atual}-01-02"
            data_impares = f"{ano_atual}-01-01"
            
            db.session.execute(text(f"UPDATE policial SET data_base_escala = '{data_pares}' WHERE inicio_pares = 1 AND data_base_escala IS NULL;"))
            db.session.execute(text(f"UPDATE policial SET data_base_escala = '{data_impares}' WHERE (inicio_pares = 0 OR inicio_pares IS NULL) AND data_base_escala IS NULL;"))
            db.session.commit()
            print("  ✅ Dados da escala antiga protegidos e configurados com sucesso!")
        except Exception as e:
            db.session.rollback()
            print(f"  ❌ Erro ao atualizar os dados: {e}")
            
        print("\n🎉 BANCO DE DADOS OFICIAL ATUALIZADO SEM PERDER NENHUM POLICIAL!")

if __name__ == '__main__':
    atualizar_banco_seguro()
    input("\nAperte ENTER para fechar esta janela...") 