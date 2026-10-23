"""
Interface gráfica em CustomTkinter — Labimed Download (CHAA)
"""
import os
import threading
import traceback
from datetime import datetime
from tkinter import filedialog, messagebox
import customtkinter as ctk

import config as cfg
from bot import LabimedBot
from utilidades import (
    formatar_mascara_data_hora,
    montar_caminho_internacao,
    preparar_pastas,
)


class App:
    def __init__(self, root: ctk.CTk):
        self.root = root
        self.root.title("Labimed CHAA - Automação de Laudos Médicos")
        self.root.geometry("720x740")
        self.root.minsize(680, 680)

        self.bot: LabimedBot | None = None
        self.rodando = False
        self.fechando = False
        self.internacoes_entries: list[tuple[ctk.CTkEntry, ctk.CTkEntry, ctk.CTkFrame]] = []
        self.root.protocol("WM_DELETE_WINDOW", self._ao_fechar_janela)
        self._construir_interface()
        self.root.lift()
        self.root.attributes("-topmost", True)
        self.root.after(400, lambda: self.root.attributes("-topmost", False))

    def _construir_interface(self) -> None:
        container = ctk.CTkFrame(self.root, corner_radius=10)
        container.pack(fill="both", expand=True, padx=16, pady=16)

        ctk.CTkLabel(
            container,
            text="Download Automatizado de Laudos — Labimed",
            font=ctk.CTkFont(family="Segoe UI", size=18, weight="bold"),
        ).pack(anchor="w", padx=16, pady=(14, 10))

        # --- Bloco 1: Credenciais ---
        cred = ctk.CTkFrame(container)
        cred.pack(fill="x", padx=16, pady=(0, 10))
        ctk.CTkLabel(cred, text="Usuário Labimed:", font=ctk.CTkFont(weight="bold")).grid(
            row=0, column=0, padx=(12, 6), pady=10, sticky="w"
        )
        self.usuario_entry = ctk.CTkEntry(cred, width=180, placeholder_text="Digite seu usuário")
        self.usuario_entry.grid(row=0, column=1, padx=(0, 14), pady=10, sticky="w")
        if cfg.get_usuario():
            self.usuario_entry.insert(0, cfg.get_usuario())

        ctk.CTkLabel(cred, text="Senha:", font=ctk.CTkFont(weight="bold")).grid(
            row=0, column=2, padx=(0, 6), pady=10, sticky="w"
        )
        self.senha_entry = ctk.CTkEntry(cred, width=160, show="•", placeholder_text="Senha de acesso")
        self.senha_entry.grid(row=0, column=3, padx=(0, 10), pady=10, sticky="w")
        if cfg.get_senha():
            self.senha_entry.insert(0, cfg.get_senha())

        self.mostrar_senha_var = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(
            cred, text="Mostrar", variable=self.mostrar_senha_var, command=self._alternar_senha, width=70
        ).grid(row=0, column=4, padx=(0, 12), pady=10, sticky="w")

        # --- Bloco 2: Paciente e Pasta Destino ---
        dados = ctk.CTkFrame(container)
        dados.pack(fill="x", padx=16, pady=(0, 10))
        dados.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(dados, text="Nome do Paciente:", font=ctk.CTkFont(weight="bold")).grid(
            row=0, column=0, padx=12, pady=(10, 6), sticky="w"
        )
        self.nome_entry = ctk.CTkEntry(dados, placeholder_text="Digite o nome completo do paciente")
        self.nome_entry.grid(row=0, column=1, columnspan=2, padx=(0, 12), pady=(10, 6), sticky="ew")

        ctk.CTkLabel(dados, text="Salvar na Pasta:", font=ctk.CTkFont(weight="bold")).grid(
            row=1, column=0, padx=12, pady=(6, 10), sticky="w"
        )
        self.pasta_entry = ctk.CTkEntry(dados)
        self.pasta_entry.grid(row=1, column=1, padx=(0, 8), pady=(6, 10), sticky="ew")
        self.pasta_entry.insert(0, cfg.get_pasta_destino_padrao())

        botoes_pasta = ctk.CTkFrame(dados, fg_color="transparent")
        botoes_pasta.grid(row=1, column=2, padx=(0, 12), pady=(6, 10))
        ctk.CTkButton(botoes_pasta, text="Escolher...", width=90, command=self._escolher_pasta).pack(
            side="left", padx=(0, 6)
        )
        ctk.CTkButton(
            botoes_pasta,
            text="Abrir Pasta",
            width=90,
            fg_color="#45475a",
            hover_color="#585b70",
            command=self._abrir_pasta_destino,
        ).pack(side="left")

        # --- Bloco 3: Internações ---
        ctk.CTkLabel(
            container,
            text="Períodos de Internação (digite apenas números para máscara DD/MM/AAAA HH:MM):",
            font=ctk.CTkFont(weight="bold"),
        ).pack(anchor="w", padx=16, pady=(2, 4))

        self.internacoes_scroll = ctk.CTkScrollableFrame(container, height=130)
        self.internacoes_scroll.pack(fill="x", padx=16, pady=(0, 6))
        self.adicionar_linha()

        acoes_int = ctk.CTkFrame(container, fg_color="transparent")
        acoes_int.pack(fill="x", padx=16, pady=(0, 8))
        ctk.CTkButton(acoes_int, text="+ Adicionar Internação", width=160, command=self.adicionar_linha).pack(
            side="left", padx=(0, 8)
        )
        ctk.CTkButton(
            acoes_int,
            text="- Remover Última",
            width=140,
            fg_color="#45475a",
            hover_color="#585b70",
            command=self.remover_linha,
        ).pack(side="left")

        # --- Bloco 4: Opções de PDF e .env ---
        opcoes = ctk.CTkFrame(container)
        opcoes.pack(fill="x", padx=16, pady=(0, 10))
        self.identificar_exames_var = ctk.BooleanVar(value=True)
        self.unificar_pdfs_var = ctk.BooleanVar(value=True)
        self.salvar_env_var = ctk.BooleanVar(value=False)
        for texto, var in (
            ("Identificar exames no nome do PDF", self.identificar_exames_var),
            ("Gerar PDF unificado por internação", self.unificar_pdfs_var),
            ("Lembrar login/pasta (.env)", self.salvar_env_var),
        ):
            ctk.CTkCheckBox(opcoes, text=texto, variable=var).pack(side="left", padx=12, pady=8)

        # --- Bloco 5: Botões Iniciar / Parar ---
        btn_frame = ctk.CTkFrame(container, fg_color="transparent")
        btn_frame.pack(fill="x", padx=16, pady=(0, 10))
        fonte_btn = ctk.CTkFont(weight="bold", size=14)
        self.btn_iniciar = ctk.CTkButton(
            btn_frame, text="INICIAR DOWNLOAD", height=38, font=fonte_btn,
            fg_color="#2ea043", hover_color="#238636", command=self.iniciar,
        )
        self.btn_iniciar.pack(side="left", expand=True, fill="x", padx=(0, 8))

        self.btn_parar = ctk.CTkButton(
            btn_frame, text="PARAR", height=38, font=fonte_btn,
            fg_color="#da3633", hover_color="#b62324", state="disabled", command=self.parar,
        )
        self.btn_parar.pack(side="left", expand=True, fill="x")

        # --- Bloco 6: Log ---
        ctk.CTkLabel(container, text="Log de Execução:", font=ctk.CTkFont(weight="bold")).pack(
            anchor="w", padx=16, pady=(0, 4)
        )
        self.log_text = ctk.CTkTextbox(
            container, height=160, font=ctk.CTkFont(family="Consolas", size=12), state="disabled"
        )
        self.log_text.pack(fill="both", expand=True, padx=16, pady=(0, 14))

    def _alternar_senha(self) -> None:
        self.senha_entry.configure(show="" if self.mostrar_senha_var.get() else "•")

    def _escolher_pasta(self) -> None:
        pasta = filedialog.askdirectory(
            initialdir=self.pasta_entry.get().strip() or os.getcwd(),
            title="Selecione a pasta onde os laudos serão salvos",
        )
        if pasta:
            self.pasta_entry.delete(0, "end")
            self.pasta_entry.insert(0, os.path.normpath(pasta))

    def _abrir_pasta_destino(self) -> None:
        pasta = self.pasta_entry.get().strip() or cfg.get_pasta_destino_padrao()
        try:
            os.makedirs(pasta, exist_ok=True)
            if hasattr(os, "startfile"):
                os.startfile(pasta)
        except OSError as e:
            messagebox.showerror("Erro ao abrir pasta", f"Não foi possível acessar '{pasta}':\n{e}")

    def _aplicar_mascara(self, event, entry: ctk.CTkEntry) -> None:
        if getattr(event, "keysym", "") in ("BackSpace", "Delete", "Left", "Right", "Tab"):
            return
        atual = entry.get()
        formatado = formatar_mascara_data_hora(atual)
        if formatado != atual:
            entry.delete(0, "end")
            entry.insert(0, formatado)

    def adicionar_linha(self) -> None:
        num = len(self.internacoes_entries) + 1
        frame = ctk.CTkFrame(self.internacoes_scroll, fg_color="transparent")
        frame.pack(fill="x", pady=3)

        ctk.CTkLabel(frame, text=f"#{num}  De:", width=55, anchor="w").pack(side="left")
        ini = ctk.CTkEntry(frame, width=180, placeholder_text="DD/MM/AAAA HH:MM")
        ini.pack(side="left", padx=(0, 12))
        ini.bind("<KeyRelease>", lambda e, ent=ini: self._aplicar_mascara(e, ent))

        ctk.CTkLabel(frame, text="Até:", width=35, anchor="w").pack(side="left")
        fim = ctk.CTkEntry(frame, width=180, placeholder_text="DD/MM/AAAA HH:MM")
        fim.pack(side="left", padx=(0, 8))
        fim.bind("<KeyRelease>", lambda e, ent=fim: self._aplicar_mascara(e, ent))

        self.internacoes_entries.append((ini, fim, frame))

    def remover_linha(self) -> None:
        if len(self.internacoes_entries) > 1:
            _, _, frame = self.internacoes_entries.pop()
            frame.destroy()
        else:
            messagebox.showwarning("Atenção", "É necessário manter ao menos uma internação.")

    def _atualizar_log(self, msg: str) -> None:
        if self.fechando:
            return
        self.log_text.configure(state="normal")
        self.log_text.insert("end", msg + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def log(self, msg: str) -> None:
        if not self.fechando:
            self.root.after(0, self._atualizar_log, msg)

    def iniciar(self) -> None:
        usuario = self.usuario_entry.get().strip()
        senha = self.senha_entry.get().strip()
        nome = self.nome_entry.get().strip()
        pasta_base = self.pasta_entry.get().strip() or cfg.get_pasta_destino_padrao()

        if not usuario or not senha:
            messagebox.showerror("Erro", "Informe o usuário e a senha do Labimed.")
            return
        if not nome:
            messagebox.showerror("Erro", "Informe o nome do paciente.")
            return

        periodos: list[tuple[datetime, datetime]] = []
        for idx, (ini_e, fim_e, _) in enumerate(self.internacoes_entries, start=1):
            it, ft = ini_e.get().strip(), fim_e.get().strip()
            try:
                ini_dt = datetime.strptime(it, "%d/%m/%Y %H:%M")
                fim_dt = datetime.strptime(ft, "%d/%m/%Y %H:%M")
            except ValueError:
                messagebox.showerror(
                    "Erro", f"Data inválida na internação #{idx}: '{it}' até '{ft}'.\nUse DD/MM/AAAA HH:MM."
                )
                return
            if ini_dt > fim_dt:
                messagebox.showerror("Erro", f"Internação #{idx}: data inicial ({it}) maior que a final ({ft}).")
                return
            periodos.append((ini_dt, fim_dt))

        try:
            if self.salvar_env_var.get():
                cfg.salvar_config_env(usuario, senha, pasta_base)
            pasta_paciente, pasta_temp = preparar_pastas(nome, pasta_base)
            usados: set[str] = set()
            internacoes_dict = {
                montar_caminho_internacao(pasta_paciente, ini_dt, usados): (ini_dt, fim_dt)
                for ini_dt, fim_dt in periodos
            }
        except OSError as e:
            messagebox.showerror("Erro de Permissão/Caminho", f"Não foi possível preparar as pastas:\n{e}")
            return

        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")
        self.btn_iniciar.configure(state="disabled")
        self.btn_parar.configure(state="normal")
        self.rodando = True

        bot = LabimedBot(
            usuario=usuario,
            senha=senha,
            pasta_temp=pasta_temp,
            log_func=self.log,
            identificar_exames=self.identificar_exames_var.get(),
            unificar_pdfs=self.unificar_pdfs_var.get(),
        )
        self.bot = bot
        threading.Thread(target=self._rodar, args=(bot, internacoes_dict, pasta_paciente), daemon=True).start()

    def _rodar(
        self, bot: LabimedBot, internacoes_dict: dict[str, tuple[datetime, datetime]], pasta_paciente: str
    ) -> None:
        try:
            self.log(f"Pasta de destino do paciente: {pasta_paciente}")
            bot.iniciar_browser()
            if bot.parar:
                self.log("\n=== DOWNLOAD INTERROMPIDO PELO USUÁRIO ===")
                return
            ok, err = bot.executar(internacoes_dict)

            if self.fechando:
                return
            if bot.parar:
                self.log("\n=== DOWNLOAD INTERROMPIDO PELO USUÁRIO ===")
                self.root.after(
                    0,
                    lambda: messagebox.showinfo(
                        "Interrompido", f"{ok} laudo(s) salvo(s) antes de parar.\nPasta: {pasta_paciente}"
                    ),
                )
            else:
                self.log(f"\n{'=' * 45}\n=== PROCESSAMENTO CONCLUÍDO ===\nSucesso: {ok} | Falhas: {err}")
                self.log(f"Arquivos salvos em: {pasta_paciente}")
                self.root.after(
                    0,
                    lambda: messagebox.showinfo(
                        "Concluído",
                        f"Download finalizado!\n\nSucesso: {ok}\nFalhas: {err}\n\nPasta:\n{pasta_paciente}",
                    ),
                )
        except Exception as e:
            erro_msg = str(e)
            self.log(f"\n[ERRO CRÍTICO] {erro_msg}")
            self.log(traceback.format_exc())
            if not self.fechando:
                self.root.after(0, lambda m=erro_msg: messagebox.showerror("Erro", m))
        finally:
            bot.fechar_browser()
            if not self.fechando:
                self.root.after(0, self._limpar_botoes)

    def parar(self) -> None:
        if self.bot and self.rodando:
            self.bot.parar = True
            self.btn_parar.configure(state="disabled")
            self.log("\n>>> Solicitação de parada recebida. Aguardando exame atual... <<<")

    def _limpar_botoes(self) -> None:
        if self.fechando:
            return
        self.rodando = False
        self.btn_iniciar.configure(state="normal")
        self.btn_parar.configure(state="disabled")

    def _ao_fechar_janela(self) -> None:
        self.fechando = True
        if self.bot and self.rodando:
            self.bot.parar = True
            threading.Thread(target=self.bot.fechar_browser, daemon=True).start()
        self.root.destroy()
