import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
import threading
import queue
import sys
import io
import os
import traceback


class _QueueStream(io.StringIO):
    """Redireciona print() para a fila da GUI."""
    def __init__(self, q):
        super().__init__()
        self.q = q

    def write(self, text):
        if text and text.strip():
            self.q.put(("log", text.rstrip()))

    def flush(self):
        pass


class App(tk.Tk):
    def __init__(self):
        super().__init__()

        from comparativo_de_caixa import VERSION
        self.version = VERSION

        self.title(f"Comparativo de Extratos - v{VERSION}")
        self.resizable(False, False)
        self.geometry("500x440")
        self.configure(bg="#f0f0f0")

        self.q = queue.Queue()
        self._update_event = threading.Event()
        self._update_choice = False

        self._build_ui()
        self.protocol("WM_DELETE_WINDOW", self.destroy)
        self.after(100, self._poll_queue)
        self.after(400, lambda: threading.Thread(target=self._run_update_check, daemon=True).start())

    def _build_ui(self):
        # Cabeçalho azul
        header = tk.Frame(self, bg="#1a4a7a", pady=14)
        header.pack(fill="x")
        tk.Label(
            header, text="Comparativo de Extratos",
            fg="white", bg="#1a4a7a", font=("Segoe UI", 15, "bold")
        ).pack()
        tk.Label(
            header, text=f"v{self.version}",
            fg="#90b8d8", bg="#1a4a7a", font=("Segoe UI", 9)
        ).pack()

        # Conteúdo principal
        body = tk.Frame(self, bg="#f0f0f0", padx=24, pady=16)
        body.pack(fill="both", expand=True)

        # Status atual
        self.status_var = tk.StringVar(value="Pronto para processar.")
        self.status_label = tk.Label(
            body, textvariable=self.status_var,
            bg="#f0f0f0", font=("Segoe UI", 10), anchor="w", fg="#333"
        )
        self.status_label.pack(fill="x", pady=(0, 6))

        # Barra de progresso
        style = ttk.Style()
        style.theme_use("default")
        style.configure(
            "green.Horizontal.TProgressbar",
            troughcolor="#ddd", background="#2e7d32", thickness=18
        )
        self.progress = ttk.Progressbar(
            body, orient="horizontal", length=452,
            mode="determinate", maximum=100,
            style="green.Horizontal.TProgressbar"
        )
        self.progress.pack(fill="x", pady=(0, 14))

        # Área de log
        tk.Label(
            body, text="Log de execução:", bg="#f0f0f0",
            font=("Segoe UI", 9), fg="#666", anchor="w"
        ).pack(fill="x")
        self.log = scrolledtext.ScrolledText(
            body, height=11, font=("Consolas", 9),
            bg="white", fg="#222", state="disabled",
            relief="solid", bd=1, wrap="word"
        )
        self.log.pack(fill="both", expand=True, pady=(2, 0))

        # Rodapé com botão
        footer = tk.Frame(self, bg="#f0f0f0", pady=14)
        footer.pack(fill="x", padx=24)
        self.btn = ttk.Button(
            footer, text="▶  Processar", command=self._on_process, width=22
        )
        self.btn.pack()

    # ──────────────────────── helpers de UI ────────────────────────

    def _append_log(self, msg, color=None):
        self.log.config(state="normal")
        tag = None
        if color:
            tag = color
            self.log.tag_config(tag, foreground=color)
            self.log.insert("end", msg + "\n", tag)
        else:
            self.log.insert("end", msg + "\n")
        self.log.see("end")
        self.log.config(state="disabled")

    def _set_progress(self, value, status=None, color="#333"):
        self.progress["value"] = value
        if status:
            self.status_var.set(status)
            self.status_label.config(fg=color)

    # ──────────────────────── processamento ────────────────────────

    def _run_update_check(self):
        from comparativo_de_caixa import VERSION
        from updater import check_for_update, download_and_apply
        self.q.put(("log", "Verificando atualizações..."))
        try:
            update_info = check_for_update(VERSION)
            if update_info:
                self.q.put(("update_prompt", update_info))
                self._update_event.wait()
                self._update_event.clear()
                if self._update_choice:
                    tag, url = update_info
                    self.q.put(("log", f"Baixando v{tag}..."))
                    download_and_apply(url, tag)
            else:
                self.q.put(("log", "Versão atualizada."))
        except Exception as e:
            self.q.put(("log", f"Aviso: {e}"))

    def _on_process(self):
        self.btn.config(state="disabled")
        self.log.config(state="normal")
        self.log.delete("1.0", "end")
        self.log.config(state="disabled")
        self._set_progress(0, "Iniciando...", "#333")
        threading.Thread(target=self._run, daemon=True).start()

    def _run(self):
        old_stdout = sys.stdout
        sys.stdout = _QueueStream(self.q)
        try:
            from comparativo_de_caixa import CashFlowComparative

            def progress_cb(pct, msg):
                self.q.put(("progress", (pct, msg, "#333")))

            CashFlowComparative(progress_callback=progress_cb)
            self.q.put(("done", None))

        except Exception as e:
            self.q.put(("error", (str(e), traceback.format_exc())))
        finally:
            sys.stdout = old_stdout

    def _poll_queue(self):
        try:
            while True:
                kind, data = self.q.get_nowait()

                if kind == "log":
                    self._append_log(f"  {data}")

                elif kind == "progress":
                    pct, msg, color = data
                    self._set_progress(pct, msg, color)
                    self._append_log(f"• {msg}")

                elif kind == "update_prompt":
                    tag, url = data
                    choice = messagebox.askyesno(
                        "Atualização disponível",
                        f"Nova versão {tag} disponível.\n\nDeseja atualizar agora?\n"
                        "(O programa será reiniciado automaticamente)"
                    )
                    self._update_choice = choice
                    self._update_event.set()

                elif kind == "done":
                    self._on_done()

                elif kind == "error":
                    self._on_error(*data)

        except queue.Empty:
            pass
        self.after(100, self._poll_queue)

    def _on_done(self):
        self._set_progress(100, "✓  Concluído com sucesso!", "#2e7d32")
        self._append_log("─" * 48)
        self._append_log("Arquivos gerados na mesma pasta do programa:")
        self._append_log("  • comparativo_de_caixa.xlsx")
        self._append_log("  • fluxo_de_caixa.xlsx")
        self.btn.config(state="normal", text="▶  Processar novamente")

    def _on_error(self, msg, tb):
        self._set_progress(self.progress["value"], "✗  Erro ao processar", "#c62828")
        self._append_log("─" * 48)

        err = msg.lower()
        err_type = tb.split("Error:")[0].split("\n")[-1].strip() if "Error:" in tb else ""

        if "permission" in err:
            friendly = "O arquivo está aberto no Excel. Feche-o e tente novamente."
        elif "no such file" in err or "filenotfound" in err.replace(" ", ""):
            friendly = "Arquivo não encontrado. Verifique as pastas Extratos/ e Sophia/."
        elif "no sheet named" in err or "worksheet" in err:
            friendly = f"Aba da planilha não encontrada:\n  {msg}"
        elif "engine" in err or "format cannot be determined" in err:
            friendly = f"Formato de arquivo não reconhecido:\n  {msg}"
        elif "keyerror" in err_type.lower():
            friendly = f"Coluna não encontrada na planilha:\n  {msg}"
        else:
            friendly = msg

        self._append_log(f"ERRO: {friendly}", color="#c62828")

        log_path = os.path.join(os.getcwd(), "erro_log.txt")
        with open(log_path, "w", encoding="utf-8") as f:
            f.write(tb)
        self._append_log(f"\nLog técnico salvo em: erro_log.txt", color="#888")
        self.btn.config(state="normal")


if __name__ == "__main__":
    app = App()
    app.mainloop()
