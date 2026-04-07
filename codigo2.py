import tkinter as tk
from tkinter import ttk, messagebox
import threading
import time
import os
import shutil
import re
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# ============================================================
# CONFIGURAÇÕES
# ============================================================
USUARIO = "P197937"
SENHA = "264Y9"
URL_SITE = "https://lis.labimed.com.br/shift/lis/labimed/elis/s01.iu.web.Login.cls?config=UNICO#dadosOs"

ID_CAMPO_LOGIN = "control_40"
ID_CAMPO_SENHA = "control_42"
ID_BOTAO_ENTRAR = "control_51"
NOME_SELETOR_LISTA = "listaOS"
CLASSE_DATA = "lbDataOs"
XPATH_IMPRIMIR_LAUDO = "//a[contains(text(), 'Imprimir laudo')]"
XPATH_FECHAR_POPUP = "//input[@value='X']"

# Estado global
navegador = None
wait = None
pasta_temp_global = None


# ============================================================
# FUNÇÕES DE PASTA
# ============================================================
def limpar_pasta_temp(pasta):
    if os.path.exists(pasta):
        for arq in os.listdir(pasta):
            caminho = os.path.join(pasta, arq)
            if os.path.isfile(caminho):
                os.remove(caminho)


def preparar_pastas(nome_paciente):
    base_dir = os.getcwd()
    pasta_paciente = os.path.join(base_dir, nome_paciente)
    pasta_temp = os.path.join(base_dir, "temp_download")

    if os.path.exists(pasta_temp):
        shutil.rmtree(pasta_temp)
    os.makedirs(pasta_temp)

    return pasta_paciente, pasta_temp


# ============================================================
# BROWSER
# ============================================================
def iniciar_browser(pasta_temp):
    global navegador, wait

    chrome_options = Options()
    prefs = {
        "download.default_directory": pasta_temp,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True,
    }
    chrome_options.add_experimental_option("prefs", prefs)

    servico = Service(ChromeDriverManager().install())
    navegador = webdriver.Chrome(service=servico, options=chrome_options)
    navegador.maximize_window()
    wait = WebDriverWait(navegador, 20)


def fechar_popup():
    global wait, navegador
    try:
        botao_fechar = wait.until(EC.element_to_be_clickable(
            (By.XPATH, XPATH_FECHAR_POPUP)
        ))
        botao_fechar.click()
        time.sleep(2)
    except:
        pass


# ============================================================
# LOGIN / RELOGIN
# ============================================================
def fazer_login(log_func):
    global navegador, wait
    log_func("[Relogin] Abrindo página de login...")
    navegador.get(URL_SITE)
    time.sleep(2)

    try:
        wait.until(EC.presence_of_element_located(
            (By.ID, ID_CAMPO_LOGIN)
        )).send_keys(USUARIO)
        navegador.find_element(By.ID, ID_CAMPO_SENHA).send_keys(SENHA)
        navegador.find_element(By.ID, ID_BOTAO_ENTRAR).click()
        log_func("[Relogin] Login realizado, aguardando lista carregar...")
        wait.until(EC.presence_of_element_located(
            (By.NAME, NOME_SELETOR_LISTA)
        ))
        time.sleep(3)
        log_func("[Relogin] Login OK.")
        return True
    except Exception as e:
        log_func(f"[ERRO] Falha ao fazer login: {e}")
        return False


def verificar_sessao():
    global navegador
    try:
        campo_login = navegador.find_elements(By.ID, ID_CAMPO_LOGIN)
        if campo_login and campo_login[0].is_displayed():
            url_atual = navegador.current_url.lower()
            if "login" in url_atual:
                return False
        return True
    except:
        return True


def garantir_sessao(log_func):
    if not verificar_sessao():
        log_func("\n[Sessão expirada detectada!]")
        log_func("[Relogin] Tentando reconectar...")
        return fazer_login(log_func)
    return True


def fechar_navegador():
    global navegador
    if navegador is not None:
        navegador.quit()


# ============================================================
# DOWNLOAD
# ============================================================
def esperar_download_acabar(pasta, timeout=40):
    for _ in range(timeout):
        if not os.path.exists(pasta):
            time.sleep(1)
            continue
        pdfs = [a for a in os.listdir(pasta) if a.endswith(".pdf")]
        if pdfs:
            return pdfs[0]
        time.sleep(1)
    if os.path.exists(pasta):
        arquivos = os.listdir(pasta)
        if arquivos:
            return None  # Retorna None mas arquivos existem
    return None


def baixar_exame(item, caminho_final, texto_data, contador_lab, indice_item, log_func):
    global navegador, wait, pasta_temp_global

    try:
        if not garantir_sessao(log_func):
            log_func(f"[Reencontrando item após relogin...]")
            itens_atuais = navegador.find_elements(By.NAME, NOME_SELETOR_LISTA)
            if indice_item < len(itens_atuais):
                item = itens_atuais[indice_item]
            else:
                log_func(f"   [ERRO] Item não encontrado após relogin.")
                return False, contador_lab

        navegador.execute_script(
            "arguments[0].scrollIntoView({block: 'center'});", item
        )
        item.click()
        time.sleep(1)

        limpar_pasta_temp(pasta_temp_global)

        botao_imprimir = wait.until(EC.element_to_be_clickable(
            (By.XPATH, XPATH_IMPRIMIR_LAUDO)
        ))
        botao_imprimir.click()

        nome_arquivo = esperar_download_acabar(pasta_temp_global)

        if nome_arquivo:
            for _ in range(10):
                novo_nome = f"LAB{contador_lab}.pdf"
                destino = os.path.join(caminho_final, novo_nome)
                if not os.path.exists(destino):
                    break
                contador_lab += 1

            shutil.move(os.path.join(pasta_temp_global, nome_arquivo), destino)
            log_func(f"   [OK] Salvo: {novo_nome} (data: {texto_data})")
            contador_lab += 1
            sucesso = True
        else:
            log_func(f"   [ERRO] Download falhou.")
            sucesso = False

        fechar_popup()
        return sucesso, contador_lab

    except Exception as e:
        log_func(f"   Erro ao baixar exame {texto_data}: {e}")
        try:
            fechar_popup()
        except:
            pass
        return False, contador_lab


def converter_data_site(texto):
    try:
        return datetime.strptime(texto.strip(), "%d/%m/%Y - %H:%M")
    except ValueError:
        return None


# ============================================================
# FUNÇÃO PRINCIPAL (roda em thread separada)
# ============================================================
def executar_roda(nome_paciente, internacoes_dict, log_callback, resultado_callback):
    global pasta_temp_global

    def log(msg):
        log_callback(msg)

    try:
        pasta_paciente, pasta_temp_global = preparar_pastas(nome_paciente)
        iniciar_browser(pasta_temp_global)

        # Login inicial
        if not fazer_login(log):
            resultado_callback("ERRO", "Não foi possível fazer login.")
            return

        itens_exames = navegador.find_elements(By.NAME, NOME_SELETOR_LISTA)
        log(f"Total de exames na lista: {len(itens_exames)}")

        total_baixado = 0
        total_falhou = 0

        for idx, (nome_pasta, (inicio, fim)) in enumerate(internacoes_dict.items()):
            caminho_final = os.path.join(pasta_paciente, nome_pasta)
            if not os.path.exists(caminho_final):
                os.makedirs(caminho_final)

            log(f"\n{'='*40}")
            log(f">>> Buscando: {inicio.strftime('%d/%m/%Y %H:%M')} ate {fim.strftime('%d/%m/%Y %H:%M')}")

            exames_encontrados = []
            for i in range(len(itens_exames)):
                if not garantir_sessao(log):
                    itens_exames = navegador.find_elements(By.NAME, NOME_SELETOR_LISTA)

                try:
                    itens_atuais = navegador.find_elements(By.NAME, NOME_SELETOR_LISTA)
                    if i >= len(itens_atuais):
                        break
                    item = itens_atuais[i]

                    navegador.execute_script(
                        "arguments[0].scrollIntoView({block: 'center'});", item
                    )
                    span_data = item.find_element(By.CLASS_NAME, CLASSE_DATA)
                    texto_data = span_data.text
                    data_exame = converter_data_site(texto_data)

                    if data_exame is None:
                        continue

                    if inicio <= data_exame <= fim:
                        log(f"  --> Encontrado: {texto_data}")
                        exames_encontrados.append((data_exame, item, i, texto_data))

                except Exception as e:
                    log(f"  Erro ao ler item {i}: {e}")
                    continue

            exames_encontrados.sort(key=lambda x: x[0], reverse=True)
            log(f"  {len(exames_encontrados)} exame(s) encontrado(s). Baixando...")

            contador_lab = 1
            for _, item, indice, texto_data in exames_encontrados:
                try:
                    if not garantir_sessao(log):
                        itens_exames = navegador.find_elements(By.NAME, NOME_SELETOR_LISTA)
                        itens_atuais = navegador.find_elements(By.NAME, NOME_SELETOR_LISTA)
                        if indice < len(itens_atuais):
                            item = itens_atuais[indice]

                    sucesso, contador_lab = baixar_exame(
                        item, caminho_final, texto_data,
                        contador_lab, indice, log
                    )
                    if sucesso:
                        total_baixado += 1
                    else:
                        total_falhou += 1

                    time.sleep(2)
                except Exception as e:
                    log(f"   Erro inesperado: {e}")
                    try:
                        fechar_popup()
                    except:
                        pass
                    total_falhou += 1

        log(f"\n{'='*40}")
        log(f"--- FINALIZADO ---")
        log(f"Baixados: {total_baixado} | Falharam: {total_falhou}")
        log(f"Pasta: {pasta_paciente}")

        resultado_callback("SUCESSO", f"{total_baixado} laudo(s) salvo(s) em:\n{pasta_paciente}")

    except Exception as e:
        import traceback
        log(f"\nERRO CRÍTICO: {e}")
        log(traceback.format_exc())
        resultado_callback("ERRO", str(e))

    finally:
        fechar_navegador()


# ============================================================
# INTERFACE GRÁFICA
# ============================================================
class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Labimed - Download de Laudos")
        self.root.geometry("550x520")
        self.root.resizable(False, False)

        # Frame principal com margin
        main = ttk.Frame(root, padding=20)
        main.pack(fill=tk.BOTH, expand=True)

        # --- Nome do paciente ---
        ttk.Label(main, text="Nome do Paciente:", font=("Segoe UI", 10, "bold")).pack(anchor="w")
        self.nome_entry = ttk.Entry(main, font=("Segoe UI", 10), width=50)
        self.nome_entry.pack(fill=tk.X, pady=(0, 10))

        # --- Internações ---
        ttk.Label(main, text="Internações (início → fim):", font=("Segoe UI", 10, "bold")).pack(anchor="w")
        self.internacoes_frame = ttk.Frame(main)
        self.internacoes_frame.pack(fill=tk.X, pady=(0, 5))

        self.internacoes_entries = []
        self.adicionar_internacao()  # Primeira linha

        self.btn_add = ttk.Button(main, text="+ Adicionar internação",
                                  command=self.adicionar_internacao)
        self.btn_add.pack(anchor="w", pady=(0, 10))

        self.btn_rem = ttk.Button(main, text="- Remover última",
                                  command=self.remover_internacao)
        self.btn_rem.pack(anchor="w", pady=(0, 15))

        # --- Botão iniciar ---
        self.btn_iniciar = ttk.Button(main, text="INICIAR DOWNLOAD",
                                      command=self.iniciar, style="Accent.TButton")
        self.btn_iniciar.pack(fill=tk.X, pady=(0, 10))

        # --- Log ---
        ttk.Label(main, text="Log:", font=("Segoe UI", 10, "bold")).pack(anchor="w")
        self.log_text = tk.Text(main, height=10, font=("Consolas", 9), state=tk.DISABLED)
        self.log_text.pack(fill=tk.BOTH, expand=True)

        # Scrollbar do log
        scrollbar = ttk.Scrollbar(self.log_text, command=self.log_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.configure(yscrollcommand=scrollbar.set)

    def adicionar_internacao(self):
        frame = ttk.Frame(self.internacoes_frame)
        frame.pack(fill=tk.X)

        ttk.Label(frame, text="De:", width=5).pack(side=tk.LEFT)
        inicio = ttk.Entry(frame, font=("Segoe UI", 9), width=18)
        inicio.insert(0, "DD/MM/AAAA HH:MM")
        inicio.pack(side=tk.LEFT, padx=(0, 5))

        ttk.Label(frame, text="Até:", width=5).pack(side=tk.LEFT)
        fim = ttk.Entry(frame, font=("Segoe UI", 9), width=18)
        fim.insert(0, "DD/MM/AAAA HH:MM")
        fim.pack(side=tk.LEFT, padx=(0, 5))

        self.internacoes_entries.append((inicio, fim))

    def remover_internacao(self):
        if len(self.internacoes_entries) > 1:
            entry = self.internacoes_entries.pop()
            entry[0].master.destroy()
        else:
            messagebox.showwarning("Atenção", "É necessário pelo menos uma internação.")

    def log(self, msg):
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, msg + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)

    def iniciar(self):
        nome = self.nome_entry.get().strip()
        if not nome:
            messagebox.showerror("Erro", "Informe o nome do paciente.")
            return

        internacoes_dict = {}
        for inicio_entry, fim_entry in self.internacoes_entries:
            inicio_texto = inicio_entry.get().strip()
            fim_texto = fim_entry.get().strip()

            if "DD/" in inicio_texto or "DD/" in fim_texto:
                messagebox.showerror("Erro", "Preencha todas as datas no formato DD/MM/AAAA HH:MM")
                return

            try:
                inicio_dt = datetime.strptime(inicio_texto, "%d/%m/%Y %H:%M")
                fim_dt = datetime.strptime(fim_texto, "%d/%m/%Y %H:%M")
                chave = f"{inicio_dt.strftime('%d-%m')} - internacao"
                internacoes_dict[chave] = (inicio_dt, fim_dt)
            except ValueError:
                messagebox.showerror("Erro", f"Formato de data inválido: {inicio_texto} / {fim_texto}")
                return

        # Limpa log
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete("1.0", tk.END)
        self.log_text.config(state=tk.DISABLED)

        self.btn_iniciar.config(state=tk.DISABLED)
        self.root.config(cursor="watch")

        thread = threading.Thread(target=executar_roda, args=(
            nome,
            internacoes_dict,
            self.log,
            self.resultado,
        ), daemon=True)
        thread.start()

    def resultado(self, tipo, msg):
        self.btn_iniciar.config(state=tk.NORMAL)
        self.root.config(cursor="")

        if tipo == "SUCESSO":
            messagebox.showinfo("Concluído", msg)
        else:
            messagebox.showerror("Erro", msg)


if __name__ == "__main__":
    root = tk.Tk()
    App(root)
    root.mainloop()
