"""
Bot Selenium — login, sessão, navegação e download de laudos
"""
import time
import os
import shutil
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

import config as cfg
from utilidades import (
    converter_data_site,
    limpar_pasta_temp,
    esperar_download,
)


class LabimedBot:
    def __init__(self, pasta_temp, log_func):
        self.pasta_temp = pasta_temp
        self.log = log_func
        self.navegador = None
        self.wait = None
        self.parar = False  # flag para botão de parar

    # ---- browser ----
    def iniciar_browser(self):
        chrome_options = Options()
        prefs = {
            "download.default_directory": self.pasta_temp,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "plugins.always_open_pdf_externally": True,
        }
        chrome_options.add_experimental_option("prefs", prefs)

        servico = Service(ChromeDriverManager().install())
        self.navegador = webdriver.Chrome(service=servico, options=chrome_options)
        self.navegador.maximize_window()
        self.wait = WebDriverWait(self.navegador, 20)
        self.log("[Browser] Chrome iniciado.")

    def fechar_browser(self):
        if self.navegador is not None:
            self.navegador.quit()
            self.log("[Browser] Fechado.")

    # ---- login ----
    def fazer_login(self):
        self.log("[Relogin] Abrindo página de login...")
        self.navegador.get(cfg.URL_SITE)
        time.sleep(2)

        try:
            self.wait.until(EC.presence_of_element_located(
                (By.ID, cfg.ID_CAMPO_LOGIN)
            )).send_keys(cfg.USUARIO)
            self.navegador.find_element(By.ID, cfg.ID_CAMPO_SENHA).send_keys(cfg.SENHA)
            self.navegador.find_element(By.ID, cfg.ID_BOTAO_ENTRAR).click()

            self.log("[Relogin] Aguardando lista carregar...")
            self.wait.until(EC.presence_of_element_located(
                (By.NAME, cfg.NOME_SELETOR_LISTA)
            ))
            time.sleep(3)
            self.log("[Relogin] Login OK.")
            return True
        except Exception as e:
            self.log(f"[ERRO] Falha ao fazer login: {e}")
            return False

    def verificar_sessao(self):
        try:
            campos = self.navegador.find_elements(By.ID, cfg.ID_CAMPO_LOGIN)
            if campos and campos[0].is_displayed():
                if "login" in self.navegador.current_url.lower():
                    return False
            return True
        except:
            return True

    def garantir_sessao(self):
        if not self.verificar_sessao():
            self.log("\n[Sessão expirada! Reconectando...]")
            return self.fazer_login()
        return True

    # ---- popup ----
    def fechar_popup(self):
        try:
            btn = self.wait.until(EC.element_to_be_clickable(
                (By.XPATH, cfg.XPATH_FECHAR_POPUP)
            ))
            btn.click()
            time.sleep(2)
        except:
            pass

    # ---- download ----
    def clicar_e_baixar(self, item, caminho_final, texto_data, contador_lab, indice):
        if self.parar:
            return False, contador_lab

        try:
            # Verifica sessão antes de cada ação
            if not self.garantir_sessao():
                itens = self.navegador.find_elements(By.NAME, cfg.NOME_SELETOR_LISTA)
                if indice < len(itens):
                    item = itens[indice]

            self.navegador.execute_script(
                "arguments[0].scrollIntoView({block: 'center'});", item
            )
            item.click()
            time.sleep(1)

            # Limpa temp e clica em imprimir
            limpar_pasta_temp(self.pasta_temp)
            btn = self.wait.until(EC.element_to_be_clickable(
                (By.XPATH, cfg.XPATH_IMPRIMIR_LAUDO)
            ))
            btn.click()

            nome_arq = esperar_download(self.pasta_temp)

            if nome_arq:
                # Evita conflito de nome
                for _ in range(10):
                    novo = f"LAB{contador_lab}.pdf"
                    dest = os.path.join(caminho_final, novo)
                    if not os.path.exists(dest):
                        break
                    contador_lab += 1

                shutil.move(os.path.join(self.pasta_temp, nome_arq), dest)
                self.log(f"   [OK] {novo} (data: {texto_data})")
                contador_lab += 1
                return True, contador_lab
            else:
                self.log(f"   [ERRO] Falha no download.")
                return False, contador_lab

        except Exception as e:
            self.log(f"   Erro em {texto_data}: {e}")
            try:
                self.fechar_popup()
            except:
                pass
            return False, contador_lab

    # ---- fluxo principal ----
    def executar(self, internacoes_dict):
        """
        internacoes_dict = { nome_pasta: (inicio, fim), ... }
        retorna (total_sucesso, total_falhas)
        """
        if not self.fazer_login():
            return 0, 0

        itens = self.navegador.find_elements(By.NAME, cfg.NOME_SELETOR_LISTA)
        self.log(f"Total de exames na lista: {len(itens)}")

        total_ok = 0
        total_err = 0

        for nome_pasta, (inicio, fim) in internacoes_dict.items():
            if self.parar:
                self.log("\n[Parado pelo usuário]")
                break

            self.log(f"\n{'='*40}")
            self.log(f">>> {inicio.strftime('%d/%m/%Y %H:%M')} ate "
                     f"{fim.strftime('%d/%m/%Y %H:%M')}")

            encontrados = []
            for i in range(len(itens)):
                if self.parar:
                    break
                if not self.garantir_sessao():
                    itens = self.navegador.find_elements(By.NAME, cfg.NOME_SELETOR_LISTA)

                try:
                    atuais = self.navegador.find_elements(By.NAME, cfg.NOME_SELETOR_LISTA)
                    if i >= len(atuais):
                        break
                    item = atuais[i]

                    self.navegador.execute_script(
                        "arguments[0].scrollIntoView({block: 'center'});", item
                    )
                    txt = item.find_element(By.CLASS_NAME, cfg.CLASSE_DATA).text
                    dt = converter_data_site(txt)

                    if dt and inicio <= dt <= fim:
                        self.log(f"  --> {txt}")
                        encontrados.append((dt, item, i, txt))
                except Exception as e:
                    self.log(f"  Erro ao ler item {i}: {e}")

            encontrados.sort(key=lambda x: x[0], reverse=True)
            self.log(f"  {len(encontrados)} exame(s) encontrado(s). Baixando...")

            contador = 1
            for _, item, indice, txt in encontrados:
                if self.parar:
                    self.log("\n[Parado pelo usuário]")
                    break

                ok, contador = self.clicar_e_baixar(
                    item, nome_pasta, txt, contador, indice
                )
                if ok:
                    total_ok += 1
                else:
                    total_err += 1

                if not self.parar:
                    time.sleep(2)  # pausa entre downloads

            self.log(f"  Baixados nesta internação: {contador - 1}")

        return total_ok, total_err
