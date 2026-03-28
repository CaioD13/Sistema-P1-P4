from flask import Blueprint, render_template, redirect, url_for, flash, request
from flask_login import login_required
from werkzeug.security import generate_password_hash

from models import Usuario
from extensions import db

usuarios_bp = Blueprint('usuarios', __name__, url_prefix='/usuarios')


@usuarios_bp.route('/')
@login_required
def listar():
    usuarios = Usuario.query.order_by(Usuario.nome).all()
    return render_template('usuarios.html', usuarios=usuarios)


@usuarios_bp.route('/novo', methods=['GET', 'POST'])
@login_required
def novo():
    if request.method == 'POST':
        username = (request.form.get('username') or '').strip()
        re = (request.form.get('re') or '').strip()
        nome = (request.form.get('nome') or '').strip()
        senha = request.form.get('password') or ''
        perfil = request.form.get('perfil') or 'comum'

        if not username or not re or not nome or not senha:
            flash('Preencha todos os campos obrigatórios.', 'error')
            return redirect(url_for('usuarios.novo'))

        if Usuario.query.filter_by(username=username).first():
            flash('Já existe um usuário com esse login.', 'error')
            return redirect(url_for('usuarios.novo'))

        if Usuario.query.filter_by(re=re).first():
            flash('Já existe um usuário com esse RE.', 'error')
            return redirect(url_for('usuarios.novo'))

        usuario = Usuario(
            username=username,
            re=re,
            nome=nome,
            password=generate_password_hash(senha),
            perfil=perfil
        )
        db.session.add(usuario)
        db.session.commit()

        flash('Usuário cadastrado com sucesso.', 'success')
        return redirect(url_for('usuarios.listar'))

    return render_template('usuario_form.html')


@usuarios_bp.route('/editar/<int:id>', methods=['GET', 'POST'])
@login_required
def editar(id):
    usuario = Usuario.query.get_or_404(id)

    if request.method == 'POST':
        usuario.username = (request.form.get('username') or '').strip()
        usuario.re = (request.form.get('re') or '').strip()
        usuario.nome = (request.form.get('nome') or '').strip()
        usuario.perfil = request.form.get('perfil') or 'comum'

        nova_senha = request.form.get('password') or ''
        if nova_senha:
            usuario.password = generate_password_hash(nova_senha)

        db.session.commit()
        flash('Usuário atualizado com sucesso.', 'success')
        return redirect(url_for('usuarios.listar'))

    return render_template('usuario_form.html', usuario=usuario)


@usuarios_bp.route('/excluir/<int:id>', methods=['POST'])
@login_required
def excluir(id):
    usuario = Usuario.query.get_or_404(id)
    db.session.delete(usuario)
    db.session.commit()
    flash('Usuário removido com sucesso.', 'success')
    return redirect(url_for('usuarios.listar'))