import queue
import threading
import tkinter as tk
from tkinter import ttk

from classes import (
    FaturamentoCompleto,
    FaturamentoAtipico,
    FaturamentoDeAcordo,
    FaturamentoSaoSebastiao,
    GerarRelatorio,
    ProgramaCopiarPeriodo,
    ProgramaRemoverPeriodo,
)


class DesktopApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Yuta - Central de Processos")
        self.geometry("860x620")
        self.minsize(760, 520)

        self._log_queue = queue.Queue()
        self._running = False

        self._build_style()
        self._build_layout()
        self._poll_log()

    def _build_style(self):
        style = ttk.Style()
        style.theme_use("clam")

        style.configure("TFrame", background="#eef2f7")
        style.configure("Header.TLabel", background="#eef2f7", foreground="#0f172a", font=("Segoe UI", 18, "bold"))
        style.configure("Sub.TLabel", background="#eef2f7", foreground="#334155", font=("Segoe UI", 10))

        style.configure("Card.TFrame", background="#ffffff")
        style.configure("CardTitle.TLabel", background="#ffffff", foreground="#0f172a", font=("Segoe UI", 12, "bold"))

        style.configure("Action.TButton", font=("Segoe UI", 10), padding=(12, 8))
        style.map("Action.TButton", background=[("active", "#dbeafe")])

        style.configure("Status.TLabel", background="#ffffff", foreground="#0f172a", font=("Segoe UI", 10))

    def _build_layout(self):
        root = ttk.Frame(self)
        root.pack(fill="both", expand=True)

        header = ttk.Frame(root)
        header.pack(fill="x", padx=18, pady=(16, 8))

        ttk.Label(header, text="Central Yuta", style="Header.TLabel").pack(anchor="w")
        ttk.Label(
            header,
            text="Selecione uma acao. O resultado aparece no painel abaixo.",
            style="Sub.TLabel",
        ).pack(anchor="w")

        body = ttk.Frame(root)
        body.pack(fill="both", expand=True, padx=18, pady=8)

        left = ttk.Frame(body, style="Card.TFrame")
        left.pack(side="left", fill="y", padx=(0, 12), pady=4)

        ttk.Label(left, text="Menu", style="CardTitle.TLabel").pack(anchor="w", padx=12, pady=(10, 6))

        self._buttons = []
        for label, action in self._menu_items():
            btn = ttk.Button(left, text=label, style="Action.TButton", command=lambda a=action: self._run_action(a))
            btn.pack(fill="x", padx=12, pady=6)
            self._buttons.append(btn)

        ttk.Button(left, text="Limpar log", command=self._clear_log).pack(fill="x", padx=12, pady=(14, 12))

        right = ttk.Frame(body, style="Card.TFrame")
        right.pack(side="left", fill="both", expand=True, pady=4)

        ttk.Label(right, text="Saida", style="CardTitle.TLabel").pack(anchor="w", padx=12, pady=(10, 6))

        log_frame = ttk.Frame(right, style="Card.TFrame")
        log_frame.pack(fill="both", expand=True, padx=12, pady=(0, 8))

        self._log_text = tk.Text(
            log_frame,
            wrap="word",
            height=18,
            bg="#0b1220",
            fg="#e2e8f0",
            insertbackground="#e2e8f0",
            font=("Consolas", 10),
        )
        self._log_text.pack(side="left", fill="both", expand=True)

        scrollbar = ttk.Scrollbar(log_frame, command=self._log_text.yview)
        scrollbar.pack(side="right", fill="y")
        self._log_text.configure(yscrollcommand=scrollbar.set)

        footer = ttk.Frame(root, style="Card.TFrame")
        footer.pack(fill="x", padx=18, pady=(0, 16))

        self._status = ttk.Label(footer, text="Pronto", style="Status.TLabel")
        self._status.pack(anchor="w", padx=12, pady=8)

    def _menu_items(self):
        return [
            ("Faturamento (Normal)", lambda: FaturamentoCompleto().executar()),
            ("Faturamento (Atipico)", lambda: FaturamentoAtipico().executar()),
            ("Faturamento Sao Sebastiao", lambda: FaturamentoSaoSebastiao().executar()),
            ("De Acordo", lambda: FaturamentoDeAcordo().executar()),
            ("Fazer Ponto", lambda: ProgramaCopiarPeriodo(debug=True).executar()),
            ("Desfazer Ponto", lambda: ProgramaRemoverPeriodo(debug=True).executar()),
            ("Relatorio", self._relatorio_safe),
        ]

    def _relatorio_safe(self):
        if hasattr(GerarRelatorio, "executar"):
            GerarRelatorio().executar()
        else:
            self._write_log("Relatorio nao implementado\n")

    def _run_action(self, action):
        if self._running:
            return
        self._running = True
        self._set_buttons_state("disabled")
        self._status.configure(text="Executando...")

        def job():
            try:
                action()
                self._write_log("\n✅ Concluido.\n")
            except Exception as exc:
                self._write_log(f"\n❌ Erro: {exc}\n")
            finally:
                self._log_queue.put("__DONE__")

        threading.Thread(target=job, daemon=True).start()

    def _set_buttons_state(self, state):
        for btn in self._buttons:
            btn.configure(state=state)

    def _clear_log(self):
        self._log_text.delete("1.0", "end")

    def _set_idle(self):
        self._running = False
        self._set_buttons_state("normal")
        self._status.configure(text="Pronto")

    def _write_log(self, text):
        self._log_queue.put(text)

    def _poll_log(self):
        try:
            while True:
                msg = self._log_queue.get_nowait()
                if msg == "__DONE__":
                    self._set_idle()
                    continue
                self._log_text.insert("end", msg)
                self._log_text.see("end")
        except queue.Empty:
            pass
        self.after(100, self._poll_log)


def run_desktop():
    app = DesktopApp()
    app.mainloop()
