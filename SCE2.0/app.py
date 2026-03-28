from flask import Flask
from extensions import db, login_manager
from config import Config
from routes.auth import auth_bp
from routes.index import index_bp
from routes.policiais import policiais_bp
from routes.escala import escala_bp
from routes.frequencia import frequencia_bp
from routes.usuarios import usuarios_bp


def create_app():
    app = Flask(__name__)
    app.config.from_object(Config)

    db.init_app(app)
    login_manager.init_app(app)

    app.register_blueprint(auth_bp)
    app.register_blueprint(index_bp)
    app.register_blueprint(policiais_bp)
    app.register_blueprint(escala_bp)
    app.register_blueprint(frequencia_bp)
    app.register_blueprint(usuarios_bp)

    with app.app_context():
        from models import Usuario

        if not Usuario.query.filter_by(username='admin').first():
            from werkzeug.security import generate_password_hash

            admin = Usuario(
                username='admin',
                re='admin',
                nome='Administrador',
                password=generate_password_hash('admin'),
                perfil='admin'
            )
            db.session.add(admin)
            db.session.commit()

    return app


app = create_app()


if __name__ == '__main__':
    from waitress import serve
    serve(app, host='0.0.0.0', port=80)
    # app.run(debug=True)