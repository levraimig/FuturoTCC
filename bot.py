"""
Bot Selenium — login, sessão, navegação, download e organização de laudos
"""
from collections.abc import Callable
from datetime import datetime
import os
import shutil
import time
from selenium import webdriver
from selenium.common.exceptions import WebDriverException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
import selenium.webdriver.chrome.webdriver  # noqa: F401 (necessário para PyInstaller)
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager

import config as cfg
from utilidades import (
    converter_data_site,
    esperar_download,
    extrair_codigo_os_texto,
    extrair_resumo_laudo_pdf,
    limpar_pasta_temp,
    unificar_pdfs_internacao,
)


class LabimedBot:
    def __init__(
        self,
        usuario: str,
        senha: str,
        pasta_temp: str,
        log_func: Callable[[str], None],
        identificar_exames: bool = True,
        unificar_pdfs: bool = True,
    ):
        self.usuario = usuario
        self.senha = senha
        self.pasta_temp = os.path.abspath(pasta_temp)
        self.log = log_func
        self.identificar_exames = identificar_exames
        self.unificar_pdfs = unificar_pdfs
        self.navegador: webdriver.Chrome | None = None
        self.wait: WebDriverWait | None = None
        self.aba_principal: str | None = None
        self.parar = False

    # ---- browser ----
    def iniciar_browser(self) -> None:
        chrome_options = Options()
        prefs = {
            "download.default_directory": self.pasta_temp,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "plugins.always_open_pdf_externally": True,
        }
        chrome_options.add_experimental_option("prefs", prefs)

        try:
            servico = Service(ChromeDriverManager().install())
            self.navegador = webdriver.Chrome(service=servico, options=chrome_options)
        except Exception:
            self.navegador = webdriver.Chrome(options=chrome_options)

        self.navegador.maximize_window()
        self.wait = WebDriverWait(self.navegador, 25)
        self.log("[Browser] Google Chrome iniciado.")

    def fechar_browser(self) -> None:
        if self.navegador is not None:
            try:
                self.navegador.quit()
            except Exception:
                pass
            finally:
                self.navegador = None
                self.log("[Browser] Fechado.")
        if os.path.exists(self.pasta_temp):
            shutil.rmtree(self.pasta_temp, ignore_errors=True)

    # ---- abas secundárias (Laudo.cls) ----
    def _fechar_abas_extras(self) -> None:
        if self.navegador is None:
            return
        try:
            handles = list(self.navegador.window_handles)
            if not handles:
                return
            principal = (
                self.aba_principal if self.aba_principal in handles else handles[0]
            )
            for h in handles:
                if h != principal:
                    try:
                        self.navegador.switch_to.window(h)
                        self.navegador.close()
                    except Exception:
                        pass
            self.navegador.switch_to.window(principal)
            self.aba_principal = principal
        except Exception:
            pass

    # ---- login ----
    def fazer_login(self) -> bool:
        if self.parar or self.navegador is None or self.wait is None:
            return False
        self.log("[Login] Abrindo portal Labimed...")
        self.navegador.get(cfg.URL_SITE)
        self.aba_principal = self.navegador.current_window_handle
        time.sleep(2)

        try:
            campo_login = self.wait.until(
                EC.presence_of_element_located((By.ID, cfg.ID_CAMPO_LOGIN))
            )
            campo_login.clear()
            campo_login.send_keys(self.usuario)

            campo_senha = self.navegador.find_element(By.ID, cfg.ID_CAMPO_SENHA)
            campo_senha.clear()
            campo_senha.send_keys(self.senha)

            self.navegador.find_element(By.ID, cfg.ID_BOTAO_ENTRAR).click()

            self.log("[Login] Aguardando lista de exames carregar...")
            self.wait.until(
                EC.presence_of_element_located((By.NAME, cfg.NOME_SELETOR_LISTA))
            )
            time.sleep(2)
            self.log("[Login] Autenticado com sucesso.")
            return True
        except Exception as e:
            self.log(f"[ERRO] Falha ao fazer login: {e}")
            return False

    def verificar_sessao(self) -> bool:
        if self.navegador is None:
            return False
        try:
            self._fechar_abas_extras()
            _ = self.navegador.current_url
            campos = self.navegador.find_elements(By.ID, cfg.ID_CAMPO_LOGIN)
            if campos and campos[0].is_displayed():
                if "login" in self.navegador.current_url.lower():
                    return False
            return True
        except WebDriverException:
            return False

    def garantir_sessao(self) -> bool:
        if not self.verificar_sessao():
            self.log("\n[Sessão expirada! Reconectando...]")
            return self.fazer_login()
        return True

    # ---- popup ----
    def fechar_popup(self) -> None:
        if self.navegador is None:
            return
        self._fechar_abas_extras()
        try:
            # Esconde o botão de imprimir e fecha o modal via JS/clique para evitar clique fantasma no próximo exame
            self.navegador.execute_script(
                """
                try {
                    if (typeof zenSetProp === 'function') zenSetProp('imprimir', 'hidden', true);
                    if (typeof zenPage !== 'undefined') zenPage.MostandoResultado = false;
                } catch(e) {}
                """
            )
            botoes = self.navegador.find_elements(By.XPATH, cfg.XPATH_FECHAR_POPUP)
            if botoes and botoes[0].is_displayed():
                botoes[0].click()
                time.sleep(0.8)
        except Exception:
            pass

    def _aguardar_modal_os(self, codigo_os: str, timeout: float = 20.0) -> bool:
        """
        Aguarda até que o modal do Labimed tenha carregado os procedimentos da O.S. correta
        e o botão 'Imprimir laudo' esteja visível.
        """
        if self.navegador is None:
            return False
        if not codigo_os:
            return True

        inicio_t = time.time()
        while time.time() - inicio_t < timeout:
            if self.parar:
                return False
            try:
                pronto = self.navegador.execute_script(
                    """
                    try {
                        var lbl = zenPage.getComponentById('labelOs');
                        var osAtual = lbl ? String(lbl.getValue() || '').trim() : '';
                        var btn = document.querySelector("a[onclick*='imprimirLaudo']");
                        var btnVisivel = Boolean(btn && btn.offsetParent !== null);
                        var resultadosProntos = Boolean(zenPage.MostandoResultado);
                        return {os: osAtual, visivel: btnVisivel, pronto: resultadosProntos};
                    } catch (e) {
                        return {os: '', visivel: true, pronto: true};
                    }
                    """
                )
                if isinstance(pronto, dict):
                    os_modal = pronto.get("os", "")
                    if (not os_modal or os_modal == codigo_os) and pronto.get("visivel") and pronto.get("pronto"):
                        return True
            except Exception:
                return True
            time.sleep(0.4)
        return False

    # ---- download ----
    def clicar_e_baixar(
        self,
        caminho_final: str,
        texto_data: str,
        contador_lab: int,
        indice: int,
        codigo_os: str = "",
    ) -> tuple[bool, int]:
        if self.parar or self.navegador is None or self.wait is None:
            return False, contador_lab

        try:
            if not self.garantir_sessao():
                self.log("   [ERRO] Sessão perdida e falha ao reconectar.")
                return False, contador_lab

            self.fechar_popup()
            limpar_pasta_temp(self.pasta_temp)

            itens = self.navegador.find_elements(By.NAME, cfg.NOME_SELETOR_LISTA)
            if indice >= len(itens):
                self.log(f"   [ERRO] Item de índice {indice} não encontrado na lista.")
                return False, contador_lab

            item = itens[indice]
            self.navegador.execute_script(
                "arguments[0].scrollIntoView({block: 'center'});", item
            )
            item.click()

            if not self._aguardar_modal_os(codigo_os, timeout=20.0):
                self.log(f"   [AVISO] Modal demorou para confirmar O.S. {codigo_os}.")

            btn = self.wait.until(
                EC.element_to_be_clickable((By.XPATH, cfg.XPATH_IMPRIMIR_LAUDO))
            )
            btn.click()

            nome_arq = esperar_download(
                self.pasta_temp,
                timeout=45,
                deve_parar=lambda: self.parar,
                codigo_os_esperado=codigo_os,
            )
            self.fechar_popup()

            if self.parar:
                return False, contador_lab

            if nome_arq:
                os.makedirs(caminho_final, exist_ok=True)
                caminho_temp_pdf = os.path.join(self.pasta_temp, nome_arq)

                resumo_exames = ""
                if self.identificar_exames:
                    resumo_exames = extrair_resumo_laudo_pdf(caminho_temp_pdf)

                sufixo = f" - {resumo_exames}" if resumo_exames else ""
                existentes = os.listdir(caminho_final)
                while any(
                    arq.upper().startswith(f"LAB{contador_lab}.")
                    or arq.upper().startswith(f"LAB{contador_lab} - ")
                    for arq in existentes
                ):
                    contador_lab += 1

                novo = f"LAB{contador_lab}{sufixo}.pdf"
                dest = os.path.join(caminho_final, novo)
                shutil.move(caminho_temp_pdf, dest)
                self.log(f"   [OK] {novo} (data: {texto_data})")
                contador_lab += 1
                return True, contador_lab

            self.log(f"   [ERRO] Tempo esgotado aguardando download ({texto_data}).")
            return False, contador_lab

        except Exception as e:
            self.log(f"   [ERRO] Falha em {texto_data}: {e}")
            self.fechar_popup()
            return False, contador_lab

    # ---- leitura única do catálogo de exames ----
    def _coletar_catalogo_exames(self) -> list[tuple[datetime, int, str, str]]:
        """
        Lê de uma só vez todas as O.S. da lista do paciente: (dt, indice, texto_data, codigo_os).
        Evita reler o DOM 30 vezes quando há 30 internações.
        """
        assert self.navegador is not None
        itens = self.navegador.find_elements(By.NAME, cfg.NOME_SELETOR_LISTA)
        catalogo: list[tuple[datetime, int, str, str]] = []
        for i, item in enumerate(itens):
            try:
                txt_data = item.find_element(By.CLASS_NAME, cfg.CLASSE_DATA).text
                dt = converter_data_site(txt_data)
                if dt is not None:
                    codigo_os = extrair_codigo_os_texto(getattr(item, "text", "") or "")
                    catalogo.append((dt, i, txt_data, codigo_os))
            except Exception as e:
                self.log(f"  Aviso ao catalogar item {i}: {e}")
        return catalogo

    # ---- fluxo principal ----
    def executar(
        self, internacoes_dict: dict[str, tuple[datetime, datetime]]
    ) -> tuple[int, int]:
        """
        internacoes_dict = { caminho_pasta_internacao: (inicio_dt, fim_dt), ... }
        Retorna (total_sucesso, total_falhas)
        """
        if self.parar:
            return 0, 0
        if not self.fazer_login():
            if self.parar:
                return 0, 0
            raise RuntimeError(
                "Não foi possível autenticar no portal Labimed. Verifique seu usuário, senha e conexão."
            )

        catalogo = self._coletar_catalogo_exames()
        self.log(f"Total de exames listados no portal: {len(catalogo)}")

        total_ok = 0
        total_err = 0

        for caminho_pasta, (inicio, fim) in internacoes_dict.items():
            if self.parar:
                self.log("\n[Parado pelo usuário]")
                break

            os.makedirs(caminho_pasta, exist_ok=True)
            # Limpa PDFs de execuções anteriores na mesma pasta de internação para não misturar arquivos
            for arq_antigo in os.listdir(caminho_pasta):
                if arq_antigo.lower().endswith(".pdf"):
                    try:
                        os.remove(os.path.join(caminho_pasta, arq_antigo))
                    except OSError:
                        pass

            nome_exibicao = os.path.basename(caminho_pasta)
            self.log(f"\n{'=' * 45}")
            self.log(
                f">>> [{nome_exibicao}] {inicio.strftime('%d/%m/%Y %H:%M')} "
                f"até {fim.strftime('%d/%m/%Y %H:%M')}"
            )

            encontrados = [
                (dt, idx, txt, cod_os)
                for dt, idx, txt, cod_os in catalogo
                if inicio <= dt <= fim
            ]
            for _, _, txt, cod_os in encontrados:
                sufixo_os = f" (O.S. {cod_os})" if cod_os else ""
                self.log(f"  --> Exame no período: {txt}{sufixo_os}")

            encontrados.sort(key=lambda x: x[0], reverse=True)
            self.log(f"  {len(encontrados)} exame(s) encontrado(s). Baixando...")

            contador = 1
            baixados_internacao = 0
            for _, indice, txt, cod_os in encontrados:
                if self.parar:
                    self.log("\n[Parado pelo usuário]")
                    break

                ok, contador = self.clicar_e_baixar(
                    caminho_pasta, txt, contador, indice, codigo_os=cod_os
                )
                if ok:
                    total_ok += 1
                    baixados_internacao += 1
                else:
                    total_err += 1

                if not self.parar:
                    time.sleep(1.0)

            self.log(f"  Baixados nesta internação: {baixados_internacao}")

            if self.unificar_pdfs and baixados_internacao > 1 and not self.parar:
                caminho_unificado = unificar_pdfs_internacao(caminho_pasta)
                if caminho_unificado:
                    self.log(
                        f"  [PDF Unificado] Gerado: {os.path.basename(caminho_unificado)}"
                    )

        return total_ok, total_err
