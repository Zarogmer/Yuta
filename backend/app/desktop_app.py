import queue
import threading
import tkinter as tk
from tkinter import ttk
from pathlib import Path

from pdf2image import convert_from_path
from PIL import Image, ImageTk

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
        self.geometry("980x640")
        self.minsize(860, 560)

        self._log_queue = queue.Queue()
        self._running = False
        self._buttons = []
        self._pending_action = None
        self._pending_label = None
        self._pending_selection = None
        self._preview_pil = None
        self._preview_image = None
        self._preview_pdf_path = None
        self._preview_pages = []
        self._preview_zoom = 1.0
        self._preview_page_index = 0

        self._build_style()
        self._build_layout()
        self._configure_tags()
        self._poll_log()

    # ---------------------------
    # UI / Style
    # ---------------------------
    def _build_style(self):
        style = ttk.Style(self)
        style.theme_use("clam")

        # Base
        self.configure(bg="#0b1220")

        style.configure("App.TFrame", background="#0b1220")
        style.configure("Surface.TFrame", background="#0f172a")
        style.configure("Card.TFrame", background="#0f172a")

        style.configure(
            "Header.TLabel",
            background="#0b1220",
            foreground="#e5e7eb",
            font=("Segoe UI", 18, "bold"),
        )
        style.configure(
            "Sub.TLabel",
            background="#0b1220",
            foreground="#94a3b8",
            font=("Segoe UI", 10),
        )

        style.configure(
            "Section.TLabel",
            background="#0f172a",
            foreground="#cbd5e1",
            font=("Segoe UI", 11, "bold"),
        )

        style.configure(
            "Action.TButton",
            font=("Segoe UI", 10, "bold"),
            padding=(12, 10),
            background="#111827",
            foreground="#e5e7eb",
        )
        style.map(
            "Action.TButton",
            background=[("active", "#1f2937"), ("disabled", "#0b1220")],
            foreground=[("disabled", "#64748b")],
        )

        style.configure(
            "Ghost.TButton",
            font=("Segoe UI", 9),
            padding=(10, 8),
            background="#0f172a",
            foreground="#cbd5e1",
        )
        style.map("Ghost.TButton", background=[("active", "#111827")])

        style.configure(
            "Status.TLabel",
            background="#0f172a",
            foreground="#cbd5e1",
            font=("Segoe UI", 10),
        )

        style.configure(
            "Thin.TSeparator",
            background="#111827",
        )

    def _build_layout(self):
        root = ttk.Frame(self, style="App.TFrame")
        root.pack(fill="both", expand=True)

        # Header (top)
        header = ttk.Frame(root, style="App.TFrame")
        header.pack(fill="x", padx=18, pady=(16, 10))

        ttk.Label(header, text="Central Yuta", style="Header.TLabel").pack(anchor="w")
        ttk.Label(
            header,
            text="Escolha um processo no menu. O log aparece em tempo real aqui do lado.",
            style="Sub.TLabel",
        ).pack(anchor="w", pady=(2, 0))

        # Body
        body = ttk.Frame(root, style="App.TFrame")
        body.pack(fill="both", expand=True, padx=18, pady=(0, 16))

        # Sidebar
        sidebar = ttk.Frame(body, style="Card.TFrame")
        sidebar.pack(side="left", fill="y", ipadx=8, ipady=8)

        ttk.Label(sidebar, text="Menu", style="Section.TLabel").pack(
            anchor="w", padx=14, pady=(14, 10)
        )

        # Buttons
        for item in self._menu_items():
            btn = ttk.Button(
                sidebar,
                text=item["label"],
                style="Action.TButton",
                command=lambda i=item: self._handle_menu_action(i),
            )
            btn.pack(fill="x", padx=14, pady=6)
            self._buttons.append(btn)

        ttk.Separator(sidebar, orient="horizontal", style="Thin.TSeparator").pack(
            fill="x", padx=14, pady=(14, 10)
        )

        tools = ttk.Frame(sidebar, style="Card.TFrame")
        tools.pack(fill="x", padx=14, pady=(0, 12))

        self._btn_generate_excel = ttk.Button(
            tools,
            text="Gerar Excel",
            style="Ghost.TButton",
            command=self._run_pending_action,
            state="disabled",
        )
        self._btn_generate_excel.pack(fill="x", pady=4)

        ttk.Button(tools, text="Limpar log", style="Ghost.TButton", command=self._clear_log).pack(
            fill="x", pady=4
        )
        ttk.Button(tools, text="Copiar log", style="Ghost.TButton", command=self._copy_log).pack(
            fill="x", pady=4
        )

        # Main (log)
        main = ttk.Frame(body, style="Card.TFrame")
        main.pack(side="left", fill="both", expand=True, padx=(12, 0), ipadx=10, ipady=10)

        top_row = ttk.Frame(main, style="Card.TFrame")
        top_row.pack(fill="x", padx=12, pady=(12, 8))

        ttk.Label(top_row, text="Saída / Log", style="Section.TLabel").pack(side="left")

        self._status = ttk.Label(top_row, text="Pronto", style="Status.TLabel")
        self._status.pack(side="right")

        # Log / Preview area
        log_wrap = ttk.Frame(main, style="Card.TFrame")
        log_wrap.pack(fill="both", expand=True, padx=12, pady=(0, 10))

        self._notebook = ttk.Notebook(log_wrap)
        self._notebook.pack(fill="both", expand=True)

        log_tab = ttk.Frame(self._notebook, style="Card.TFrame")
        self._notebook.add(log_tab, text="Saida / Log")

        self._log_text = tk.Text(
            log_tab,
            wrap="word",
            bg="#0b1220",
            fg="#e5e7eb",
            insertbackground="#e5e7eb",
            relief="flat",
            font=("Consolas", 10),
            height=18,
        )
        self._log_text.pack(side="left", fill="both", expand=True)

        sb = ttk.Scrollbar(log_tab, command=self._log_text.yview)
        sb.pack(side="right", fill="y")
        self._log_text.configure(yscrollcommand=sb.set)

        preview_tab = ttk.Frame(self._notebook, style="Card.TFrame")
        self._notebook.add(preview_tab, text="Preview PDF")

        preview_toolbar = ttk.Frame(preview_tab, style="Card.TFrame")
        preview_toolbar.pack(fill="x", padx=8, pady=(8, 4))

        self._btn_prev_page = ttk.Button(
            preview_toolbar,
            text="◀",
            style="Ghost.TButton",
            command=self._prev_preview_page,
        )
        self._btn_prev_page.pack(side="left", padx=(0, 6))

        self._btn_next_page = ttk.Button(
            preview_toolbar,
            text="▶",
            style="Ghost.TButton",
            command=self._next_preview_page,
        )
        self._btn_next_page.pack(side="left", padx=(0, 12))

        self._preview_page_label = ttk.Label(
            preview_toolbar,
            text="0/0",
            style="Status.TLabel",
        )
        self._preview_page_label.pack(side="left", padx=(0, 12))

        ttk.Button(
            preview_toolbar,
            text="Zoom +",
            style="Ghost.TButton",
            command=lambda: self._set_preview_zoom(self._preview_zoom + 0.1),
        ).pack(side="left", padx=(0, 6))
        ttk.Button(
            preview_toolbar,
            text="Zoom -",
            style="Ghost.TButton",
            command=lambda: self._set_preview_zoom(max(0.3, self._preview_zoom - 0.1)),
        ).pack(side="left", padx=(0, 6))
        ttk.Button(
            preview_toolbar,
            text="Reset",
            style="Ghost.TButton",
            command=lambda: self._set_preview_zoom(1.0),
        ).pack(side="left")

        preview_wrap = ttk.Frame(preview_tab, style="Card.TFrame")
        preview_wrap.pack(fill="both", expand=True, padx=8, pady=(0, 8))

        self._preview_canvas = tk.Canvas(
            preview_wrap,
            bg="#0b1220",
            highlightthickness=0,
        )
        self._preview_canvas.pack(side="left", fill="both", expand=True)
        self._preview_canvas.bind("<Configure>", self._on_preview_resize)
        self._preview_canvas.bind("<MouseWheel>", self._on_preview_mousewheel)

        preview_scroll = ttk.Scrollbar(
            preview_wrap,
            orient="vertical",
            command=self._preview_canvas.yview,
        )
        preview_scroll.pack(side="right", fill="y")
        self._preview_canvas.configure(yscrollcommand=preview_scroll.set)

        self._preview_placeholder = self._preview_canvas.create_text(
            8,
            8,
            anchor="nw",
            text="Sem pre-visualizacao",
            fill="#94a3b8",
            font=("Segoe UI", 10),
        )

        # Status bar bottom (progress)
        bottom = ttk.Frame(main, style="Card.TFrame")
        bottom.pack(fill="x", padx=12, pady=(0, 12))

        self._progress = ttk.Progressbar(bottom, mode="indeterminate")
        self._progress.pack(fill="x")

    def _configure_tags(self):
        # tags no Text para colorir
        self._log_text.tag_configure("info", foreground="#cbd5e1")
        self._log_text.tag_configure("ok", foreground="#34d399")     # verde
        self._log_text.tag_configure("err", foreground="#fb7185")    # vermelho
        self._log_text.tag_configure("warn", foreground="#fbbf24")   # amarelo

    # ---------------------------
    # Actions
    # ---------------------------
    def _menu_items(self):
        return [
            {
                "label": "🧾 Faturamento (Normal)",
                "preview": lambda: FaturamentoCompleto().executar(preview=True),
                "action": lambda selection=None: FaturamentoCompleto().executar(
                    preview=False,
                    selection=selection,
                ),
            },
            {
                "label": "🧩 Faturamento (Atípico)",
                "preview": lambda: FaturamentoAtipico().executar(preview=True),
                "action": lambda selection=None: FaturamentoAtipico().executar(
                    preview=False,
                    selection=selection,
                ),
            },
            {
                "label": "⚓ Faturamento São Sebastião",
                "preview": lambda: FaturamentoSaoSebastiao().executar(preview=True),
                "action": lambda selection=None: FaturamentoSaoSebastiao().executar(
                    preview=False,
                    selection=selection,
                ),
            },
            {
                "label": "✅ De Acordo",
                "preview": lambda: FaturamentoDeAcordo().executar(preview=True),
                "action": lambda selection=None: FaturamentoDeAcordo().executar(
                    preview=False,
                    selection=selection,
                ),
            },
            {"label": "🕒 Fazer Ponto", "action": lambda: ProgramaCopiarPeriodo(debug=True).executar()},
            {"label": "↩️ Desfazer Ponto", "action": lambda: ProgramaRemoverPeriodo(debug=True).executar()},
            {"label": "📊 Relatório", "action": self._relatorio_safe},
        ]

    def _relatorio_safe(self):
        if hasattr(GerarRelatorio, "executar"):
            GerarRelatorio().executar()
        else:
            self._write_log("Relatório não implementado.\n", tag="warn")

    def _handle_menu_action(self, item):
        if item.get("preview"):
            self._run_preview(item["preview"], item["action"], item["label"])
        else:
            self._clear_pending_action()
            self._run_action(item["action"])

    def _run_action(self, action):
        if self._running:
            return

        self._running = True
        self._set_buttons_state("disabled")
        self._status.configure(text="Executando...")
        self._progress.start(12)

        self._write_log("\n▶ Iniciando...\n", tag="info")

        def job():
            try:
                action()
                self._write_log("\n✅ Concluído.\n", tag="ok")
            except Exception as exc:
                self._write_log(f"\n❌ Erro: {exc}\n", tag="err")
            finally:
                self._log_queue.put(("__DONE__", None))

        threading.Thread(target=job, daemon=True).start()

    def _run_preview(self, preview_action, final_action, label):
        if self._running:
            return

        self._running = True
        self._set_buttons_state("disabled")
        self._status.configure(text="Gerando pre-visualizacao...")
        self._progress.start(12)

        self._write_log("\n[PREVIEW] Iniciando...\n", tag="info")

        def job():
            try:
                result = preview_action() or ""
                preview_text = ""
                preview_pdf = None
                selection = None

                if isinstance(result, dict):
                    preview_text = result.get("text", "")
                    preview_pdf = result.get("preview_pdf")
                    selection = result.get("selection")
                else:
                    preview_text = str(result)

                if preview_text:
                    self._write_log(preview_text + "\n", tag="info")
                if preview_pdf:
                    self._log_queue.put(("__PREVIEW_PDF__", preview_pdf))
                self._set_pending_action(label, final_action, selection)
                self._write_log("[PREVIEW] Pronto. Clique em 'Gerar Excel'.\n", tag="ok")
            except Exception as exc:
                self._write_log(f"\n[PREVIEW] Erro: {exc}\n", tag="err")
            finally:
                self._log_queue.put(("__DONE__", None))

        threading.Thread(target=job, daemon=True).start()

    def _set_buttons_state(self, state):
        for btn in self._buttons:
            btn.configure(state=state)

    # ---------------------------
    # Log helpers
    # ---------------------------
    def _write_log(self, text, tag="info"):
        self._log_queue.put((text, tag))

    def _clear_log(self):
        self._log_text.delete("1.0", "end")
        self._status.configure(text="Pronto")
        self._clear_pending_action()
        self._clear_preview()

    def _copy_log(self):
        txt = self._log_text.get("1.0", "end-1c")
        self.clipboard_clear()
        self.clipboard_append(txt)
        self._status.configure(text="Log copiado ✅")

    def _set_idle(self):
        self._running = False
        self._set_buttons_state("normal")
        self._progress.stop()
        self._status.configure(text="Pronto")

    def _set_pending_action(self, label, action, selection=None):
        self._pending_action = action
        self._pending_label = label
        self._pending_selection = selection
        self._btn_generate_excel.configure(state="normal")
        self._status.configure(text="Pre-visualizacao pronta")

    def _clear_pending_action(self):
        self._pending_action = None
        self._pending_label = None
        self._pending_selection = None
        self._btn_generate_excel.configure(state="disabled")

    def _run_pending_action(self):
        if not self._pending_action:
            self._write_log("Nenhuma pre-visualizacao pendente.\n", tag="warn")
            return

        action = self._pending_action
        selection = self._pending_selection
        self._clear_pending_action()
        self._run_action(lambda: action(selection))

    def _poll_log(self):
        try:
            while True:
                msg, tag = self._log_queue.get_nowait()
                if msg == "__DONE__":
                    self._set_idle()
                    continue

                if msg == "__PREVIEW_PDF__":
                    self._show_pdf_preview(tag)
                    continue

                self._log_text.insert("end", msg, tag)
                self._log_text.see("end")
        except queue.Empty:
            pass

        self.after(80, self._poll_log)

    def _on_preview_resize(self, _event):
        if self._preview_pages:
            self._render_preview_page()

    def _on_preview_mousewheel(self, event):
        if not self._preview_pages:
            return
        delta = int(-1 * (event.delta / 120))
        self._preview_canvas.yview_scroll(delta, "units")

    def _render_preview_page(self):
        if not self._preview_pages:
            return

        width = max(self._preview_canvas.winfo_width(), 1)
        page = self._preview_pages[self._preview_page_index]
        img_width, img_height = page.size
        scale = (width / img_width) * self._preview_zoom

        new_size = (
            max(int(img_width * scale), 1),
            max(int(img_height * scale), 1),
        )
        resized = page.resize(new_size, Image.LANCZOS)

        self._preview_image = ImageTk.PhotoImage(resized)
        self._preview_canvas.delete("preview_img")
        self._preview_canvas.create_image(
            0,
            0,
            image=self._preview_image,
            anchor="nw",
            tags="preview_img",
        )

        self._preview_canvas.configure(
            scrollregion=(0, 0, resized.size[0], resized.size[1])
        )
        self._preview_canvas.yview_moveto(0)

        if self._preview_placeholder:
            self._preview_canvas.itemconfigure(self._preview_placeholder, state="hidden")

        self._update_preview_nav()

    def _set_preview_zoom(self, zoom):
        self._preview_zoom = max(0.3, min(2.5, zoom))
        self._render_preview_page()

    def _update_preview_nav(self):
        total = len(self._preview_pages)
        current = self._preview_page_index + 1 if total else 0
        self._preview_page_label.configure(text=f"{current}/{total}")

        if total <= 1:
            self._btn_prev_page.configure(state="disabled")
            self._btn_next_page.configure(state="disabled")
            return

        self._btn_prev_page.configure(
            state="normal" if self._preview_page_index > 0 else "disabled"
        )
        self._btn_next_page.configure(
            state="normal" if self._preview_page_index < total - 1 else "disabled"
        )

    def _prev_preview_page(self):
        if self._preview_page_index <= 0:
            return
        self._preview_page_index -= 1
        self._render_preview_page()

    def _next_preview_page(self):
        if self._preview_page_index >= len(self._preview_pages) - 1:
            return
        self._preview_page_index += 1
        self._render_preview_page()

    def _show_pdf_preview(self, pdf_path):
        try:
            path = Path(str(pdf_path))
            if not path.exists():
                self._write_log(f"Preview PDF nao encontrado: {path}\n", tag="warn")
                return

            poppler_path = Path(r"C:\poppler-25.12.0\Library\bin")
            kwargs = {}
            if poppler_path.exists():
                kwargs["poppler_path"] = str(poppler_path)

            pages = convert_from_path(str(path), **kwargs)
            if not pages:
                self._write_log("Nao foi possivel renderizar o PDF.\n", tag="warn")
                return

            self._preview_pages = pages
            self._preview_page_index = 0
            self._preview_pdf_path = path
            self._render_preview_page()
            self._notebook.select(1)
        except Exception as exc:
            self._write_log(f"Falha ao renderizar PDF: {exc}\n", tag="warn")

    def _clear_preview(self):
        self._preview_pil = None
        self._preview_image = None
        self._preview_pdf_path = None
        self._preview_pages = []
        self._preview_zoom = 1.0
        self._preview_page_index = 0
        self._preview_canvas.delete("preview_img")
        if self._preview_placeholder:
            self._preview_canvas.itemconfigure(self._preview_placeholder, state="normal")


def run_desktop():
    app = DesktopApp()
    app.mainloop()


if __name__ == "__main__":
    run_desktop()
