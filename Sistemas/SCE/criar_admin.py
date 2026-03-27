from app import db, Usuario
from werkzeug.security import generate_password_hash

def criar_admin():
    print("\n🔧 Criando usuário ADMIN no novo banco...\n")

    # Verifica se já existe
    admin = Usuario.query.filter_by(username="admin").first()
    if admin:
        print("✔ Já existe um admin cadastrado.")
        return

    novo = Usuario(
        username="admin",
        re="admin",
        nome="Administrador",
        password=generate_password_hash("admin"),
        perfil="admin",
    )

    db.session.add(novo)
    db.session.commit()

    print("🎉 ADMIN criado com sucesso!")
    print("Login:")
    print("  Usuário: admin")
    print("  Senha:   admin\n")

if __name__ == "__main__":
    from app import app
    with app.app_context():
        criar_admin()