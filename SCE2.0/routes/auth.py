from flask import Blueprint, render_template, redirect, url_for, flash, request
from flask_login import login_user, logout_user, login_required, current_user
from werkzeug.security import check_password_hash

from models import Usuario
from extensions import db

auth_bp = Blueprint('auth', __name__, url_prefix='/auth')


@auth_bp.route('/login', methods=['GET', 'POST'])
def login():
    if current_user.is_authenticated:
        return redirect(url_for('index.index'))

    if request.method == 'POST':
        re_login = request.form.get('username')
        senha = request.form.get('password')
        user = Usuario.query.filter_by(username=re_login).first()

        if user and check_password_hash(user.password, senha):
            login_user(user)
            flash('Login realizado com sucesso.', 'success')
            return redirect(url_for('index.index'))

        flash('RE ou Senha incorretos.', 'error')

    return render_template('login.html')


@auth_bp.route('/logout')
@login_required
def logout():
    logout_user()
    return redirect(url_for('auth.login'))