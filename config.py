"""
Configurações globais — URL, seletores, credenciais e diretório padrão via .env
"""
import os
import sys
from pathlib import Path
from dotenv import load_dotenv, set_key


def obter_diretorio_base() -> Path:
    """Retorna o diretório onde está o executável (.exe) ou o script Python."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def carregar_ambiente() -> Path:
    """Carrega variáveis do arquivo .env na pasta base ou no diretório atual."""
    caminho_env_base = obter_diretorio_base() / ".env"
    caminho_env_cwd = Path.cwd() / ".env"

    if caminho_env_base.exists():
        load_dotenv(dotenv_path=caminho_env_base, override=False)
        return caminho_env_base
    if caminho_env_cwd.exists():
        load_dotenv(dotenv_path=caminho_env_cwd, override=False)
        return caminho_env_cwd

    load_dotenv(override=False)
    return caminho_env_base


CAMINHO_ENV = carregar_ambiente()

# Seletores do site (se o Labimed atualizar, mudar aqui)
URL_SITE = "https://lis.labimed.com.br/shift/lis/labimed/elis/s01.iu.web.Login.cls?config=UNICO#dadosOs"
ID_CAMPO_LOGIN = "control_40"
ID_CAMPO_SENHA = "control_42"
ID_BOTAO_ENTRAR = "control_51"
NOME_SELETOR_LISTA = "listaOS"
CLASSE_DATA = "lbDataOs"
XPATH_IMPRIMIR_LAUDO = "//a[contains(text(), 'Imprimir laudo')]"
XPATH_FECHAR_POPUP = "//input[@value='X']"


def get_usuario() -> str:
    return os.getenv("LABIMED_USUARIO", "").strip()


def get_senha() -> str:
    return os.getenv("LABIMED_SENHA", "").strip()


def get_pasta_destino_padrao() -> str:
    pasta_env = os.getenv("LABIMED_PASTA_DESTINO", "").strip()
    if pasta_env:
        return pasta_env
    return str(obter_diretorio_base() / "Laudos_Baixados")


def salvar_config_env(usuario: str, senha: str, pasta_destino: str = "") -> None:
    """Salva ou atualiza as credenciais e pasta padrão no arquivo .env local."""
    env_file = str(CAMINHO_ENV)
    if not CAMINHO_ENV.exists():
        CAMINHO_ENV.write_text("", encoding="utf-8")
    set_key(env_file, "LABIMED_USUARIO", usuario.strip())
    set_key(env_file, "LABIMED_SENHA", senha.strip())
    if pasta_destino:
        set_key(env_file, "LABIMED_PASTA_DESTINO", pasta_destino.strip())
    os.environ["LABIMED_USUARIO"] = usuario.strip()
    os.environ["LABIMED_SENHA"] = senha.strip()
    if pasta_destino:
        os.environ["LABIMED_PASTA_DESTINO"] = pasta_destino.strip()
