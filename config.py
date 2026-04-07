"""
Configurações globais — URL, seletores, credenciais via .env
"""
import os

# Seletores do site (se o Labimed atualizar, mudar aqui)
URL_SITE = "https://lis.labimed.com.br/shift/lis/labimed/elis/s01.iu.web.Login.cls?config=UNICO#dadosOs"
ID_CAMPO_LOGIN = "control_40"
ID_CAMPO_SENHA = "control_42"
ID_BOTAO_ENTRAR = "control_51"
NOME_SELETOR_LISTA = "listaOS"
CLASSE_DATA = "lbDataOs"
XPATH_IMPRIMIR_LAUDO = "//a[contains(text(), 'Imprimir laudo')]"
XPATH_FECHAR_POPUP = "//input[@value='X']"


def get_usuario():
    return os.getenv("LABIMED_USUARIO", "")


def get_senha():
    return os.getenv("LABIMED_SENHA", "")
