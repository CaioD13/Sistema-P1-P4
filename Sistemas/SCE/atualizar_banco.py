from app import app, db

def atualizar():
    with app.app_context():
        # O comando create_all() verifica quais tabelas do app.py não existem no frequencia.db e as cria, 
        # SEM apagar ou modificar os dados das tabelas que já existem!
        db.create_all()
        print("✅ Banco de dados atualizado com sucesso!")
        print("✅ A tabela 'validacao_mensal' foi criada sem alterar seus dados antigos.")

if __name__ == '__main__':
    atualizar()