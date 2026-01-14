import sys
import time
import os
import shutil
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# --- 1. CONFIGURAÇÕES E DADOS ---
NOME_PACIENTE = "Miguel Paniz Bastos"

# SUAS DATAS DE INTERNAÇÃO (Dia/Mês/Ano Hora:Minuto)
INTERNACOES = [
    ("25/05/2025 08:00", "01/06/2025 18:00"),
    ("05/06/2025 10:30", "10/06/2025 09:00")
]

# CREDENCIAIS (Edite aqui se for rodar direto no VS Code)
if len(sys.argv) < 3:
    usuario = "SEU_USUARIO_AQUI"
    senha = "SUA_SENHA_AQUI"
else:
    usuario = sys.argv[1]
    senha = sys.argv[2]

url_site = "https://lis.labimed.com.br/shift/lis/labimed/elis/s01.iu.web.Login.cls?config=UNICO#dadosOs" # SEU LINK REAL

# --- FUNÇÕES DE DATA ---
def converter_data_site(texto):
    # O site mostra: "28/05/2025 - 13:50"
    try:
        texto = texto.strip()
        return datetime.strptime(texto, "%d/%m/%Y - %H:%M") 
    except ValueError:
        return None 

def converter_data_config(texto):
    return datetime.strptime(texto, "%d/%m/%Y %H:%M")

# --- PREPARAÇÃO DE PASTAS ---
base_dir = os.getcwd()
pasta_paciente = os.path.join(base_dir, NOME_PACIENTE)
pasta_temp = os.path.join(base_dir, "temp_download")

if os.path.exists(pasta_temp): shutil.rmtree(pasta_temp)
os.makedirs(pasta_temp)

if not os.path.exists(pasta_paciente): os.makedirs(pasta_paciente)

# --- CONFIGURAÇÃO DO CHROME ---
chrome_options = Options()
prefs = {
    "download.default_directory": pasta_temp,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "plugins.always_open_pdf_externally": True
}
chrome_options.add_experimental_option("prefs", prefs)

def esperar_download_acabar(pasta):
    print("   Baixando...", end="")
    for _ in range(40): # Espera até 40 segundos
        if not os.path.exists(pasta): return None
        arquivos = os.listdir(pasta)
        for arq in arquivos:
            if not arq.endswith(".crdownload") and not arq.endswith(".tmp"):
                print(" Concluído!")
                return arq
        time.sleep(1)
        print(".", end="")
    return None

# --- INÍCIO DO ROBÔ ---
print(f"--- INICIANDO ROBÔ PARA: {NOME_PACIENTE} ---")
servico = Service(ChromeDriverManager().install())
navegador = webdriver.Chrome(service=servico, options=chrome_options)
navegador.maximize_window()
wait = WebDriverWait(navegador, 20)

try:
    # --- LOGIN ---
    navegador.get(url_site)
    wait.until(EC.presence_of_element_located((By.ID, "control_40"))).send_keys(usuario)
    navegador.find_element(By.ID, "control_42").send_keys(senha)
    navegador.find_element(By.ID, "control_51").click()
    
    print("Login OK. Carregando lista...")
    wait.until(EC.presence_of_element_located((By.NAME, "listaOS")))
    
    # Pega a lista inicial
    itens_exames = navegador.find_elements(By.NAME, "listaOS")
    print(f"Total de exames na lista: {len(itens_exames)}")

    for periodo in INTERNACOES:
        inicio = converter_data_config(periodo[0])
        fim = converter_data_config(periodo[1])
        
        nome_pasta = f"{inicio.strftime('%d-%m')} - internacao"
        caminho_final = os.path.join(pasta_paciente, nome_pasta)
        
        if not os.path.exists(caminho_final): os.makedirs(caminho_final)
            
        print(f"\n>>> Buscando entre: {inicio} e {fim}")
        contador_lab = 1 
        
        # Percorre a lista pelo índice para evitar perder a referência
        for i in range(len(itens_exames)):
            try:
                # Recarrega a lista para garantir que o elemento existe
                itens_atuais = navegador.find_elements(By.NAME, "listaOS")
                item = itens_atuais[i]
                
                # Scroll para o item ficar visível
                navegador.execute_script("arguments[0].scrollIntoView({block: 'center'});", item)
                
                # Lê a data
                span_data = item.find_element(By.CLASS_NAME, "lbDataOs")
                texto_data_site = span_data.text 
                data_exame = converter_data_site(texto_data_site)
                
                if data_exame is None: continue 

                # Verifica se está no prazo
                if inicio <= data_exame <= fim:
                    print(f"--> Exame encontrado: {texto_data_site}")
                    
                    # 1. Abre o Pop-up
                    item.click()
                    
                    # 2. Clica em IMPRIMIR LAUDO (Usando seu print)
                    # Procura um link (a) que contenha o texto "Imprimir laudo"
                    botao_imprimir = wait.until(EC.element_to_be_clickable(
                        (By.XPATH, "//a[contains(text(), 'Imprimir laudo')]")
                    ))
                    botao_imprimir.click()
                    
                    # 3. Gerencia o Download
                    nome_arquivo = esperar_download_acabar(pasta_temp)
                    
                    if nome_arquivo:
                        novo_nome = f"LAB{contador_lab}.pdf"
                        destino = os.path.join(caminho_final, novo_nome)
                        shutil.move(os.path.join(pasta_temp, nome_arquivo), destino)
                        print(f"   [OK] Salvo em: {novo_nome}")
                        contador_lab += 1
                    else:
                        print("   [ERRO] Download falhou.")

                    # 4. FECHA O POP-UP (Usando seu print do X)
                    # Procura um input que tenha value="X"
                    print("   Fechando janela...")
                    botao_fechar = wait.until(EC.element_to_be_clickable(
                        (By.XPATH, "//input[@value='X']")
                    ))
                    botao_fechar.click()
                    
                    # Espera um pouquinho para o modal sumir totalmente
                    time.sleep(2)

            except Exception as e:
                # Se der erro em um, tenta o próximo
                # print(f"Erro leve ao ler item: {e}")
                continue

    print("\n--- Processo finalizado com sucesso! ---")

except Exception as e:
    print(f"ERRO CRÍTICO: {e}")