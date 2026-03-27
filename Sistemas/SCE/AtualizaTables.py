from sqlalchemy import text
from app import db

def adicionar_colunas_ee_se_necessario():
    with db.engine.connect() as conn:
        # Verifica colunas existentes na tabela alteracao_escala
        resultado = conn.execute(text("PRAGMA table_info(alteracao_escala)"))
        colunas_existentes = [row[1] for row in resultado.fetchall()]

        comandos = []

        if 'hora_inicio_ee' not in colunas_existentes:
            comandos.append("ALTER TABLE alteracao_escala ADD COLUMN hora_inicio_ee VARCHAR(5)")

        if 'hora_fim_ee' not in colunas_existentes:
            comandos.append("ALTER TABLE alteracao_escala ADD COLUMN hora_fim_ee VARCHAR(5)")

        if not comandos:
            print("Nenhuma alteração necessária. As colunas já existem.")
            return

        for comando in comandos:
            conn.execute(text(comando))

        conn.commit()
        print("Colunas atualizadas com sucesso.")

if __name__ == "__main__":
    adicionar_colunas_ee_se_necessario()