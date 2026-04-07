"""
Funções utilitárias — pastas, datas, download
"""
import os
import shutil
import time
from datetime import datetime


def converter_data_site(texto):
    """Converte '28/05/2025 - 13:50' para datetime."""
    try:
        return datetime.strptime(texto.strip(), "%d/%m/%Y - %H:%M")
    except ValueError:
        return None


def preparar_pastas(nome_paciente):
    """Cria pastas temp e retorna (pasta_paciente, pasta_temp)."""
    base_dir = os.getcwd()
    pasta_paciente = os.path.join(base_dir, nome_paciente)
    pasta_temp = os.path.join(base_dir, "temp_download")

    if os.path.exists(pasta_temp):
        shutil.rmtree(pasta_temp)
    os.makedirs(pasta_temp)

    return pasta_paciente, pasta_temp


def limpar_pasta_temp(pasta):
    """Remove todos os arquivos da pasta temp."""
    if os.path.exists(pasta):
        for arq in os.listdir(pasta):
            caminho = os.path.join(pasta, arq)
            if os.path.isfile(caminho):
                os.remove(caminho)


def esperar_download(pasta, timeout=40):
    """Espera até timeout segundos por um .pdf completo. Retorna nome ou None."""
    for _ in range(timeout):
        if not os.path.exists(pasta):
            time.sleep(1)
            continue
        pdfs = [a for a in os.listdir(pasta) if a.endswith(".pdf")]
        if pdfs:
            return pdfs[0]
        time.sleep(1)
    return None
