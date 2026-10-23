"""
Funções utilitárias — pastas, datas, máscara de entrada, download e manipulação de PDFs
"""
from collections.abc import Callable
from datetime import datetime
import os
import re
import shutil
import time
from pypdf import PdfReader, PdfWriter

EXAMES_CONHECIDOS = [
    "HEMOGRAMA",
    "GASOMETRIA",
    "CREATININA",
    "UREIA",
    "SODIO",
    "POTASSIO",
    "MAGNESIO",
    "CALCIO",
    "GLICOSE",
    "PCR",
    "PROTEINA C REATIVA",
    "TROPONINA",
    "TGO",
    "TGP",
    "BILIRRUBINA",
    "FOSFATASE",
    "GAMA GT",
    "AMILASE",
    "LIPASE",
    "COAGULOGRAMA",
    "TEMPO DE PROTROMBINA",
    "TTPA",
    "EQU",
    "PARCIAL DE URINA",
    "UROCULTURA",
    "HEMOCULTURA",
    "CULTURA",
    "LACTATO",
    "ALBUMINA",
    "TSH",
    "T4 LIVRE",
]

NOMES_RESERVADOS_WIN = {
    "CON", "PRN", "AUX", "NUL",
    *(f"COM{i}" for i in range(1, 10)),
    *(f"LPT{i}" for i in range(1, 10)),
}


def sanitizar_nome_arquivo(nome: str, max_len: int = 80) -> str:
    """Remove caracteres inválidos, nomes reservados e limita tamanho no Windows."""
    limpo = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "", nome or "")
    limpo = re.sub(r"\s+", " ", limpo).strip(" .")[:max_len].strip(" .")
    if not limpo or limpo.upper() in NOMES_RESERVADOS_WIN:
        return "Paciente_Sem_Nome"
    return limpo


def converter_data_site(texto: str) -> datetime | None:
    """Converte '28/05/2025 - 13:50' ou '28/05/2025 13:50' para datetime."""
    if not texto:
        return None
    limpo = texto.strip()
    for formato in ("%d/%m/%Y - %H:%M", "%d/%m/%Y %H:%M", "%d/%m/%Y - %H:%M:%S"):
        try:
            return datetime.strptime(limpo, formato)
        except ValueError:
            continue
    return None


def extrair_codigo_os_texto(texto_item: str) -> str:
    """
    Extrai o código da Ordem de Serviço do texto do item na lista (ex: '210-67720-157').
    """
    if not texto_item:
        return ""
    m = re.search(r"\b(\d{3}-\d+-\d+)\b", texto_item)
    return m.group(1) if m else ""


def formatar_mascara_data_hora(texto: str) -> str:
    """
    Aplica máscara automática DD/MM/AAAA HH:MM a partir dos dígitos informados.
    Ex: '280520251350' -> '28/05/2025 13:50'
    """
    digitos = "".join(c for c in (texto or "") if c.isdigit())[:12]
    partes = []
    if len(digitos) <= 2:
        return digitos
    partes.append(digitos[:2] + "/")
    if len(digitos) <= 4:
        return "".join(partes) + digitos[2:]
    partes.append(digitos[2:4] + "/")
    if len(digitos) <= 8:
        return "".join(partes) + digitos[4:]
    partes.append(digitos[4:8] + " ")
    if len(digitos) <= 10:
        return "".join(partes) + digitos[8:]
    partes.append(digitos[8:10] + ":" + digitos[10:12])
    return "".join(partes)


def preparar_pastas(nome_paciente: str, pasta_base: str | None = None) -> tuple[str, str]:
    """Cria a pasta do paciente e a pasta temporária de download."""
    base_dir = os.path.abspath(pasta_base) if pasta_base else os.getcwd()
    os.makedirs(base_dir, exist_ok=True)

    nome_seguro = sanitizar_nome_arquivo(nome_paciente)
    pasta_paciente = os.path.join(base_dir, nome_seguro)
    os.makedirs(pasta_paciente, exist_ok=True)

    pasta_temp = os.path.join(base_dir, "temp_download")
    if os.path.exists(pasta_temp):
        shutil.rmtree(pasta_temp, ignore_errors=True)
    os.makedirs(pasta_temp, exist_ok=True)

    return pasta_paciente, pasta_temp


def montar_caminho_internacao(
    pasta_paciente: str, inicio: datetime, usados: set[str] | None = None
) -> str:
    """Gera e cria a subpasta da internação dentro da pasta do paciente sem colisões."""
    nome_base = f"{inicio.strftime('%d-%m')} - internacao"
    nome_final = nome_base
    if usados is not None:
        sufixo = 2
        while nome_final in usados:
            nome_final = f"{nome_base} ({sufixo})"
            sufixo += 1
        usados.add(nome_final)

    caminho = os.path.join(pasta_paciente, nome_final)
    os.makedirs(caminho, exist_ok=True)
    return caminho


def limpar_pasta_temp(pasta: str) -> None:
    """Remove todos os arquivos da pasta temp."""
    if os.path.exists(pasta):
        for arq in os.listdir(pasta):
            caminho = os.path.join(pasta, arq)
            if os.path.isfile(caminho):
                try:
                    os.remove(caminho)
                except OSError:
                    pass


def pdf_contem_codigo_os(caminho_pdf: str, codigo_os: str) -> bool:
    """
    Verifica se o PDF baixado pertence realmente à Ordem de Serviço esperada (codigo_os).
    Se codigo_os for vazio, apenas valida que o PDF possui páginas legíveis.
    """
    try:
        with open(caminho_pdf, "rb") as f_in:
            leitor = PdfReader(f_in)
            if len(leitor.pages) == 0:
                return False
            if not codigo_os:
                return True
            texto_pag1 = leitor.pages[0].extract_text() or ""
            if not texto_pag1.strip():
                return True
            return codigo_os in texto_pag1
    except Exception:
        return False


def esperar_download(
    pasta: str,
    timeout: int = 45,
    intervalo: float = 0.5,
    deve_parar: Callable[[], bool] | None = None,
    codigo_os_esperado: str = "",
) -> str | None:
    """
    Espera até timeout segundos por um arquivo .pdf finalizado (sem .crdownload/.tmp ativo)
    e que pertença ao codigo_os_esperado (descartando PDFs atrasados de outra O.S.).
    """
    tentativas = max(1, int(timeout / intervalo))
    for _ in range(tentativas):
        if deve_parar and deve_parar():
            return None
        if not os.path.exists(pasta):
            time.sleep(intervalo)
            continue

        arquivos = os.listdir(pasta)
        em_andamento = any(
            a.endswith(".crdownload") or a.endswith(".tmp") for a in arquivos
        )
        pdfs = [a for a in arquivos if a.lower().endswith(".pdf")]

        if pdfs and not em_andamento:
            caminho_pdf = os.path.join(pasta, pdfs[0])
            try:
                if os.path.getsize(caminho_pdf) > 0:
                    if pdf_contem_codigo_os(caminho_pdf, codigo_os_esperado):
                        return pdfs[0]
                    # PDF residual de uma O.S. anterior que terminou atrasado: remove e continua esperando
                    os.remove(caminho_pdf)
            except OSError:
                pass

        time.sleep(intervalo)
    return None


def extrair_resumo_laudo_pdf(caminho_pdf: str, max_exames: int = 3) -> str:
    """
    Lê o texto do PDF baixado fechando o descritor imediatamente via context manager.
    Retorna string vazia caso o PDF não possua camada de texto ou nenhum exame seja identificado.
    """
    try:
        with open(caminho_pdf, "rb") as f_in:
            leitor = PdfReader(f_in)
            paginas_txt = [(p.extract_text() or "") for p in leitor.pages[:5]]
    except Exception:
        return ""

    texto_total = "\n".join(paginas_txt).upper()
    if not texto_total.strip():
        return ""

    # 1. Padrão nativo Labimed: primeira linha não vazia da página antes de "RESULTADO:" / "MATERIAL:"
    titulos_labimed: list[str] = []
    for txt_pag in paginas_txt:
        upper_pag = txt_pag.upper()
        if "RESULTADO:" in upper_pag or "MATERIAL:" in upper_pag:
            linhas = [ln.strip() for ln in txt_pag.splitlines() if ln.strip()]
            if linhas:
                primeira = linhas[0]
                if not primeira.upper().startswith(("REVISOR", "GRUPO -", "PACIENTE")):
                    titulo = primeira.split(" - ")[0].strip()
                    titulo_limpo = sanitizar_nome_arquivo(titulo, max_len=35)
                    if (
                        titulo_limpo
                        and titulo_limpo != "Paciente_Sem_Nome"
                        and titulo_limpo not in titulos_labimed
                    ):
                        titulos_labimed.append(titulo_limpo)
                        if len(titulos_labimed) >= max_exames:
                            break
    if titulos_labimed:
        return sanitizar_nome_arquivo("_".join(titulos_labimed), max_len=50)

    # 2. Fallback por palavras-chave de exames conhecidos
    encontrados = []
    for exame in EXAMES_CONHECIDOS:
        if re.search(rf"\b{re.escape(exame)}\b", texto_total):
            nome_curto = (
                exame.replace("PROTEINA C REATIVA", "PCR")
                .replace("TEMPO DE PROTROMBINA", "TP")
                .replace("PARCIAL DE URINA", "EQU")
            )
            if nome_curto not in encontrados:
                encontrados.append(nome_curto)
            if len(encontrados) >= max_exames:
                break

    if encontrados:
        return sanitizar_nome_arquivo("_".join(encontrados), max_len=50)

    match_exame = re.search(r"EXAME[ \t]*[:\-][ \t]*([A-Z0-9À-Ú ]{3,30})", texto_total)
    if match_exame:
        return sanitizar_nome_arquivo(match_exame.group(1).strip(), max_len=40)

    return ""


def unificar_pdfs_internacao(
    pasta_internacao: str, nome_saida: str = "LAUDOS_UNIFICADOS.pdf"
) -> str | None:
    """
    Junta todos os PDFs individuais (LAB*.pdf) de uma internação em um único arquivo consolidado.
    Retorna o caminho do PDF unificado ou None se não houver arquivos.
    """
    if not os.path.exists(pasta_internacao):
        return None

    arquivos_pdf = [
        arq
        for arq in os.listdir(pasta_internacao)
        if arq.lower().endswith(".pdf") and arq != nome_saida
    ]
    if not arquivos_pdf:
        return None

    def chave_ordenacao(nome: str) -> tuple[int, str]:
        m = re.match(r"^LAB(\d+)", nome, re.IGNORECASE)
        return (int(m.group(1)) if m else 999999, nome)

    arquivos_pdf.sort(key=chave_ordenacao)

    escritor = PdfWriter()
    paginas_adicionadas = 0
    for arq in arquivos_pdf:
        caminho = os.path.join(pasta_internacao, arq)
        try:
            with open(caminho, "rb") as f_in:
                leitor = PdfReader(f_in)
                for pagina in leitor.pages:
                    escritor.add_page(pagina)
                    paginas_adicionadas += 1
        except Exception:
            continue

    if paginas_adicionadas == 0:
        return None

    caminho_saida = os.path.join(pasta_internacao, nome_saida)
    with open(caminho_saida, "wb") as f_out:
        escritor.write(f_out)

    return caminho_saida
