"""
Interface gráfica — Labimed Download
Com botão PARAR para interromper a qualquer momento
"""
import tkinter as tk
from tkinter import ttk, messagebox
import threading
import os

from bot import LabimedBot
from utilidades import preparar_pastas


class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Labimed - Download de Laudos")
        self.root.geometry("580x540")
        self.root.resizable(False, False)

        self.bot = None
        self.rodando = False

        main = ttk.Frame(root, padding=15)
        main.pack(fill=tk.BOTH, expand=True)

        # --- Nome do paciente ---
        ttk.Label(main, text="Nome do Paciente:",
                  font=("Segoe UI", 10, "bold")).pack(anchor="w")
        self.nome_entry = ttk.Entry(main, font=("Segoe UI", 10), width=50)
        self.nome_entry.pack(fill=tk.X, pady=(0, 12))

        # --- Internações ---
        ttk.Label(main, text="Internações (início → fim):",
                  font=("Segoe UI", 10, "bold")).pack(anchor="w")
        self.internacoes_frame = ttk.Frame(main)
        self.internacoes_frame.pack(fill=tk.X, pady=(0, 5))

        self.internacoes_entries = []
        self.adicionar_linha()

        ttk.Button(main, text="+ Adicionar internação",
                   command=self.adicionar_linha).pack(anchor="w", pady=(0, 5))
        ttk.Button(main, text="- Remover última",
                   command=self.remover_linha).pack(anchor="w", pady=(0, 12))

        # --- Botões ---
        btn_frame = ttk.Frame(main)
        btn_frame.pack(fill=tk.X, pady=(0, 10))

        self.btn_iniciar = ttk.Button(btn_frame, text="INICIAR",
                                      command=self.iniciar)
        self.btn_iniciar.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(0, 5))

        self.btn_parar = ttk.Button(btn_frame, text="PARAR",
                                    command=self.parar,
                                    state=tk.DISABLED)
        self.btn_parar.pack(side=tk.LEFT, expand=True, fill=tk.X)

        # --- Log ---
        ttk.Label(main, text="Log:", font=("Segoe UI", 10, "bold")).pack(anchor="w")

        log_frame = ttk.Frame(main)
        log_frame.pack(fill=tk.BOTH, expand=True)

        self.log_text = tk.Text(log_frame, height=10, font=("Consolas", 8),
                                state=tk.DISABLED, bg="#1e1e2e", fg="#cdd6f4")
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(log_frame, command=self.log_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.configure(yscrollcommand=scrollbar.set)

    # ---- helpers (thread-safe) ----
    def _atualizar_log(self, msg):
        """Atualiza o log na thread principal (seguro)."""
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, msg + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)

    def log(self, msg):
        """Chamada de qualquer thread — despacha para a thread principal."""
        self.root.after(0, self._atualizar_log, msg)

    def adicionar_linha(self):
        frame = ttk.Frame(self.internacoes_frame)
        frame.pack(fill=tk.X)

        ttk.Label(frame, text="De:", width=5).pack(side=tk.LEFT)
        ini = ttk.Entry(frame, font=("Segoe UI", 9), width=18)
        ini.insert(0, "DD/MM/AAAA HH:MM")
        ini.pack(side=tk.LEFT, padx=(0, 5))

        ttk.Label(frame, text="Até:", width=5).pack(side=tk.LEFT)
        fim = ttk.Entry(frame, font=("Segoe UI", 9), width=18)
        fim.insert(0, "DD/MM/AAAA HH:MM")
        fim.pack(side=tk.LEFT, padx=(0, 5))

        self.internacoes_entries.append((ini, fim))

    def remover_linha(self):
        if len(self.internacoes_entries) > 1:
            self.internacoes_entries.pop()[0].master.destroy()
        else:
            messagebox.showwarning("Atenção", "Mínimo de uma internação.")

    # ---- ações ----
    def iniciar(self):
        from datetime import datetime

        nome = self.nome_entry.get().strip()
        if not nome:
            messagebox.showerror("Erro", "Informe o nome do paciente.")
            return

        internacoes_dict = {}
        for ini_e, fim_e in self.internacoes_entries:
            it = ini_e.get().strip()
            ft = fim_e.get().strip()

            if "DD/" in it or "DD/" in ft:
                messagebox.showerror("Erro", "Preencha todas as datas corretamente.")
                return

            try:
                ini_dt = datetime.strptime(it, "%d/%m/%Y %H:%M")
                fim_dt = datetime.strptime(ft, "%d/%m/%Y %H:%M")
                chave = f"{ini_dt.strftime('%d-%m')} - internacao"
                internacoes_dict[chave] = (ini_dt, fim_dt)
            except ValueError:
                messagebox.showerror("Erro", f"Data inválida: {it} / {ft}")
                return

        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete("1.0", tk.END)
        self.log_text.config(state=tk.DISABLED)

        self.btn_iniciar.config(state=tk.DISABLED)
        self.btn_parar.config(state=tk.NORMAL)
        self.root.config(cursor="watch")
        self.rodando = True

        pasta_paciente, pasta_temp = preparar_pastas(nome)
        self.bot = LabimedBot(usuario, senha, pasta_temp, self.log)
        self.bot.iniciar_browser()

        threading.Thread(
            target=self._rodar,
            args=(internacoes_dict, pasta_paciente),
            daemon=True,
        ).start()

    def _rodar(self, internacoes_dict, pasta_paciente):
        try:
            ok, err = self.bot.executar(internacoes_dict)

            if self.bot.parar:
                self.log("\n=== DOWNLOAD INTERROMPIDO ===")
            else:
                self.log(f"\n{'='*40}")
                self.log(f"=== CONCLUÍDO ===")
                self.log(f"Sucesso: {ok} | Falhas: {err}")
                self.log(f"Pasta: {pasta_paciente}")

            if ok > 0 and not self.bot.parar:
                self.root.after(0, lambda: messagebox.showinfo(
                    "Concluído",
                    f"{ok} laudo(s) salvo(s).\nFalhas: {err}\n\n{pasta_paciente}"
                ))
            elif self.bot.parar:
                messagebox.showinfo(
                    "Interrompido",
                    f"{ok} laudo(s) salvo(s) antes de parar."
                )
        except Exception as e:
            import traceback
            self.log(f"\nERRO CRÍTICO: {e}")
            self.log(traceback.format_exc())
            self.root.after(0, lambda: messagebox.showerror("Erro", str(e)))
        finally:
            self.bot.fechar_browser()
            self._limpar_botoes()

    def parar(self):
        if self.bot and self.rodando:
            self.bot.parar = True
            self.log("\n>>> PARANDO... aguarde o exame atual finalizar <<<")

    def _limpar_botoes(self):
        self.rodando = False
        self.btn_iniciar.config(state=tk.NORMAL)
        self.btn_parar.config(state=tk.DISABLED)
        self.root.config(cursor="")


if __name__ == "__main__":
    root = tk.Tk()
    App(root)
    root.mainloop()
