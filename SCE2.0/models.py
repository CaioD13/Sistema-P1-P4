from datetime import datetime
from flask_login import UserMixin
from extensions import db


class HistoricoEscala(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    policial_id = db.Column(db.Integer, db.ForeignKey('policial.id'), nullable=False)
    mes = db.Column(db.Integer, nullable=False)
    ano = db.Column(db.Integer, nullable=False)
    tipo_escala_antigo = db.Column(db.String(20), nullable=False)


class Usuario(UserMixin, db.Model):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(20), unique=True, nullable=False)
    re = db.Column(db.String(20), unique=True, nullable=False)
    nome = db.Column(db.String(100), nullable=False)
    password = db.Column(db.String(100), nullable=False)
    perfil = db.Column(db.String(20), default='comum')  # admin, operador, comum


class Policial(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    posto = db.Column(db.String(20))
    re = db.Column(db.String(20), unique=True)
    nome = db.Column(db.String(100))
    nome_guerra = db.Column(db.String(50))
    cia = db.Column(db.String(20))
    secao_em = db.Column(db.String(50))
    tipo_escala = db.Column(db.String(20))
    data_base_escala = db.Column(db.Date, nullable=True)
    funcao = db.Column(db.String(100))
    equipe = db.Column(db.String(50))
    viatura = db.Column(db.String(50))
    ativo = db.Column(db.Boolean, default=True)
    hora_inicio_escala = db.Column(db.Time, nullable=True)
    hora_fim_escala = db.Column(db.Time, nullable=True)
    duracao_horas_escala = db.Column(db.Float, nullable=True)
    regime_escala = db.Column(db.String(50), nullable=True)
    habilitacoes = db.Column(db.String(500), nullable=True)
    sat = db.Column(db.String(500), nullable=True)

    status_saude = db.Column(db.String(50), default='Apto Sem Restrições')
    restricao_tipo = db.Column(db.String(100))
    data_fim_restricao = db.Column(db.Date)


class AlteracaoEscala(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    policial_id = db.Column(db.Integer, db.ForeignKey('policial.id'))
    data = db.Column(db.Date)
    valor = db.Column(db.String(10))
    motivo = db.Column(db.String(200))
    hora_inicio_ee = db.Column(db.String(5))
    hora_fim_ee = db.Column(db.String(5))


class LogAuditoria(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    data_hora = db.Column(db.DateTime, default=datetime.now)
    usuario_id = db.Column(db.Integer)
    usuario_nome = db.Column(db.String(100))
    usuario_re = db.Column(db.String(20))
    acao = db.Column(db.String(50))
    detalhes = db.Column(db.String(500))


class HistoricoMovimentacao(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    policial_id = db.Column(db.Integer, db.ForeignKey('policial.id'))
    data_mudanca = db.Column(db.DateTime, default=datetime.now)
    tipo_mudanca = db.Column(db.String(50))
    anterior = db.Column(db.String(100))
    novo = db.Column(db.String(100))
    usuario_responsavel = db.Column(db.String(100))


class EscalaVersao(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    cia = db.Column(db.String(50))
    data_inicio = db.Column(db.Date)
    tipo_filtro = db.Column(db.String(20))
    versao_atual = db.Column(db.Integer, default=1)


class ValidacaoMensal(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    policial_id = db.Column(db.Integer, db.ForeignKey('policial.id'))
    mes = db.Column(db.Integer)
    ano = db.Column(db.Integer)
    admin_id = db.Column(db.Integer)
    admin_info = db.Column(db.String(200))
    data_validacao = db.Column(db.DateTime, default=datetime.now)