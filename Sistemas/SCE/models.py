from flask_sqlalchemy import SQLAlchemy
from flask_login import UserMixin
from datetime import datetime

db = SQLAlchemy()

class Usuario(db.Model, UserMixin):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(80), unique=True, nullable=False)
    password = db.Column(db.String(120), nullable=False)
    nome = db.Column(db.String(100))
    re = db.Column(db.String(20))
    perfil = db.Column(db.String(20), default='usuario')

class Policial(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    re = db.Column(db.String(20), unique=True, nullable=False)
    posto = db.Column(db.String(20))
    nome = db.Column(db.String(100))
    nome_guerra = db.Column(db.String(50))
    cia = db.Column(db.String(20))
    funcao = db.Column(db.String(50))
    
    # NOVOS CAMPOS (Já devem estar no banco se rodou o script)
    secao_em = db.Column(db.String(20))
    viatura = db.Column(db.String(20))
    equipe = db.Column(db.String(20))
    
    tipo_escala = db.Column(db.String(20))
    inicio_pares = db.Column(db.Boolean, default=False)
    ativo = db.Column(db.Boolean, default=True)
    
    # SAÚDE
    status_saude = db.Column(db.String(50), default='Apto Sem Restrições')
    restricao_tipo = db.Column(db.String(100), nullable=True)
    data_fim_restricao = db.Column(db.Date, nullable=True)

    def to_dict(self):
        return {
            'id': self.id,
            're': self.re,
            'posto': self.posto,
            'nome': self.nome,
            'nome_guerra': self.nome_guerra,
            'cia': self.cia,
            'secao_em': self.secao_em,
            'funcao': self.funcao,
            'viatura': self.viatura,
            'equipe': self.equipe,
            'tipo_escala': self.tipo_escala,
            'inicio_pares': self.inicio_pares,
            'status_saude': self.status_saude,
            'restricao_tipo': self.restricao_tipo,
            'data_fim_restricao': str(self.data_fim_restricao) if self.data_fim_restricao else None
        }

class AlteracaoEscala(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    policial_id = db.Column(db.Integer, db.ForeignKey('policial.id'), nullable=False)
    data = db.Column(db.Date, nullable=False)
    valor = db.Column(db.String(10), nullable=False)

class LogAuditoria(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    data_hora = db.Column(db.DateTime, default=db.func.now())
    usuario_id = db.Column(db.Integer)
    usuario_nome = db.Column(db.String(100))
    usuario_re = db.Column(db.String(20))
    acao = db.Column(db.String(50))
    detalhes = db.Column(db.String(200))

class HistoricoMovimentacao(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    policial_id = db.Column(db.Integer, db.ForeignKey('policial.id'), nullable=False)
    data_mudanca = db.Column(db.DateTime, default=datetime.now)
    tipo_mudanca = db.Column(db.String(50)) # Ex: Transferência Cia, Promoção
    anterior = db.Column(db.String(100))
    novo = db.Column(db.String(100))
    usuario_responsavel = db.Column(db.String(100))

class EscalaVersao(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    cia = db.Column(db.String(50))
    data_inicio = db.Column(db.Date)
    tipo_filtro = db.Column(db.String(20)) # OFICIAIS, PRACAS, TODOS
    versao_atual = db.Column(db.Integer, default=1)