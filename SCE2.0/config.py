import os
from datetime import timedelta


class Config:
    BASEDIR = os.path.abspath(os.path.dirname(__file__))
    SECRET_KEY = "chave_super_secreta_37bpmm_2026_final"
    SQLALCHEMY_DATABASE_URI = f"sqlite:///{os.path.join(BASEDIR, 'frequencia.db')}"
    SQLALCHEMY_TRACK_MODIFICATIONS = False
    SESSION_PERMANENT = False
    PERMANENT_SESSION_LIFETIME = timedelta(minutes=30)

    # Dados institucionais padrão do SCE
    OPM_CODIGO = "510370000"
    OPM_NOME = "37º BPM/M"