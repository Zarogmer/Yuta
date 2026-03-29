"""
Desktop Yuta - Central de Processos.
Interface Tkinter para faturamentos, ponto e relatorios.
"""
import queue
import subprocess
import sys
import threading
import tkinter as tk
import os
from datetime import datetime
from pathlib import Path
from tkinter import messagebox, ttk

from pdf2image import convert_from_path
from PIL import Image, ImageTk
from backend.app.utils.path_utils import poppler_paths_candidatos
from backend.app.yuta_helpers import fechar_workbooks

from backend.app.modules import (
    CriarPasta,
    FaturamentoAtipico,
    FaturamentoCompleto,
    FaturamentoDeAcordo,
    FaturamentoSaoSebastiao,
    GerarRelatorio,
    FazerPonto,
    ProgramaRemoverPeriodo,
)
# ---- Tema (cores e fontes) ----
THEME = {
    "bg_root": "#0b1220",
    "bg_surface": "#0f172a",
    "bg_card": "#0f172a",
    "bg_button": "#111827",
    "bg_button_hover": "#1f2937",
    "bg_disabled": "#0b1220",
    "fg_primary": "#e5e7eb",
    "fg_secondary": "#94a3b8",
    "fg_section": "#cbd5e1",
    "fg_status": "#cbd5e1",
    "fg_disabled": "#64748b",
    "accent_ok": "#34d399",
    "accent_err": "#fb7185",
    "accent_warn": "#fbbf24",
    "separator": "#111827",
    "font_header": ("Segoe UI", 18, "bold"),
    "font_sub": ("Segoe UI", 10),
    "font_section": ("Segoe UI", 11, "bold"),
    "font_button": ("Segoe UI", 10, "bold"),
    "font_ghost": ("Segoe UI", 9),
    "font_log": ("Consolas", 10),
}


class DesktopApp(tk.Tk):
    _LOCKED_MENU_KEYWORDS = (
        "RELATORIO",
        "RELATORIO",
        "GERADOR NFS-E",
    )

    def __init__(self):
        super().__init__()
        self.title("Yuta - Central de Processos")
        self._center_window(self, 1280, 900)
        self.minsize(980, 640)
        self.protocol("WM_DELETE_WINDOW", self._on_close_app)

        self._log_queue = queue.Queue()
        self._running = False
        self._buttons = []
        self._menu_modo = "principal"
        self._menu_buttons_frame = None
        self._fazer_ponto_ctx = None
        self._fazer_ponto_form_vars = None
        self._desfazer_ponto_ctx = None
        self._desfazer_ponto_form_vars = None
        self._pending_action = None
        self._pending_label = None
        self._pending_selection = None
        self._preview_pil = None
        self._preview_image = None
        self._preview_pdf_path = None
        self._preview_pages = []
        self._preview_zoom = 1.0
        self._preview_zoom_max = 2.5
        self._preview_page_index = 0
        self._preview_mode_hint = ""
        self._preview_reset_scroll = False
        self._startup_cancelled = False
        self._progress_job = None
        self._progress_value = 0.0

        self._build_style()
        self._usuario_nome = self._selecionar_usuario_inicial()
        if self._startup_cancelled:
            self.after(0, self.destroy)
            return
        self._build_layout()
        self._configure_tags()
        self._poll_log()

    def _on_close_app(self):
        self.destroy()

    @staticmethod
    def _center_window(window: tk.Misc, largura: int, altura: int):
        screen_w = window.winfo_screenwidth()
        screen_h = window.winfo_screenheight()
        x = max((screen_w // 2) - (largura // 2), 0)
        y = max((screen_h // 2) - (altura // 2) - 30, 0)
        window.geometry(f"{largura}x{altura}+{x}+{y}")

    def _selecionar_usuario_inicial(self) -> str:
        self.configure(bg="#060b1a")

        bg = tk.Frame(self, bg="#060b1a")
        bg.pack(fill="both", expand=True)

        # Card central
        card = tk.Frame(bg, bg="#0f172a", highlightthickness=1, highlightbackground="#1f2a44")
        card.place(relx=0.5, rely=0.5, anchor="center", width=520, height=420)

        header = tk.Frame(card, bg="#0f172a")
        header.pack(fill="x", padx=24, pady=(20, 8))

        tk.Label(
            header,
            text="Yuta",
            bg="#0f172a",
            fg="#e5e7eb",
            font=("Segoe UI", 28, "bold"),
        ).pack(anchor="center")
        tk.Label(
            header,
            text="Central de Processos",
            bg="#0f172a",
            fg="#9ca3af",
            font=("Segoe UI", 12),
        ).pack(anchor="center", pady=(2, 0))

        tk.Frame(card, bg="#1f2a44", height=1).pack(fill="x", pady=(8, 0))

        content = tk.Frame(card, bg="#0f172a")
        content.pack(fill="both", expand=True, padx=24, pady=(14, 24))

        tk.Label(
            content,
            text="Iniciar sessão",
            bg="#0f172a",
            fg="#f3f4f6",
            font=("Segoe UI", 15, "bold"),
        ).pack(anchor="w")
        tk.Label(
            content,
            text="Escolha o usuário",
            bg="#0f172a",
            fg="#9ca3af",
            font=("Segoe UI", 11),
        ).pack(anchor="w", pady=(2, 12))

        usuario_var = tk.StringVar(value="Carol Carmo")

        opcoes_frame = tk.Frame(content, bg="#0f172a")
        opcoes_frame.pack(fill="x")

        estilo_opcoes = {
            "normal_border": "#25324d",
            "normal_bg": "#111a2e",
            "selected_border": "#3b82f6",
            "selected_bg": "#142544",
        }

        opcoes = []

        def selecionar(nome: str):
            usuario_var.set(nome)
            atualizar_estilo_opcoes()

        def criar_opcao(nome: str, iniciais: str):
            wrap = tk.Frame(
                opcoes_frame,
                bg=estilo_opcoes["normal_border"],
                highlightthickness=0,
            )
            wrap.pack(fill="x", pady=(0, 8))

            row = tk.Frame(wrap, bg=estilo_opcoes["normal_bg"], height=48)
            row.pack(fill="x", padx=1, pady=1)
            row.pack_propagate(False)

            indicador = tk.Label(
                row,
                text="o",
                bg=estilo_opcoes["normal_bg"],
                fg="#93a4c2",
                font=("Segoe UI", 15),
                width=2,
            )
            indicador.pack(side="left", padx=(10, 4), pady=8)

            avatar = tk.Label(
                row,
                text=iniciais,
                bg="#d1d5db",
                fg="#111827",
                font=("Segoe UI", 10, "bold"),
                width=3,
                pady=2,
            )
            avatar.pack(side="left", padx=(0, 8), pady=8)

            nome_lbl = tk.Label(
                row,
                text=nome,
                bg=estilo_opcoes["normal_bg"],
                fg="#e5e7eb",
                font=("Segoe UI", 13, "bold"),
            )
            nome_lbl.pack(side="left")

            def on_click(_e=None, n=nome):
                selecionar(n)

            for widget in (wrap, row, indicador, avatar, nome_lbl):
                widget.bind("<Button-1>", on_click)

            opcoes.append({
                "nome": nome,
                "wrap": wrap,
                "row": row,
                "indicador": indicador,
                "nome_lbl": nome_lbl,
                "avatar": avatar,
            })

        def atualizar_estilo_opcoes():
            selecionado = usuario_var.get()
            for opt in opcoes:
                ativo = opt["nome"] == selecionado
                border = estilo_opcoes["selected_border"] if ativo else estilo_opcoes["normal_border"]
                fundo = estilo_opcoes["selected_bg"] if ativo else estilo_opcoes["normal_bg"]
                opt["wrap"].configure(bg=border)
                opt["row"].configure(bg=fundo)
                opt["indicador"].configure(
                    text="*" if ativo else "o",
                    bg=fundo,
                    fg="#3b82f6" if ativo else "#93a4c2",
                )
                opt["nome_lbl"].configure(bg=fundo)
                opt["avatar"].configure(bg="#f3f4f6" if ativo else "#d1d5db")

        criar_opcao("Carol Carmo", "CC")
        criar_opcao("Diogo Barros", "DB")
        atualizar_estilo_opcoes()

        escolhido = {"nome": "Carol Carmo"}
        concluido = tk.BooleanVar(value=False)

        def confirmar():
            escolhido["nome"] = usuario_var.get().strip() or "Carol Carmo"
            concluido.set(True)

        def fechar_app():
            self._startup_cancelled = True
            concluido.set(True)

        btn_entrar = tk.Button(
            content,
            text="Entrar",
            bg="#2563eb",
            fg="#ffffff",
            activebackground="#1d4ed8",
            activeforeground="#ffffff",
            relief="flat",
            bd=0,
            font=("Segoe UI", 14, "bold"),
            cursor="hand2",
            command=confirmar,
        )
        btn_entrar.pack(fill="x", pady=(12, 0), ipady=8)

        self.bind("<Return>", lambda _e: confirmar())
        self.protocol("WM_DELETE_WINDOW", fechar_app)
        self.focus_force()
        self.wait_variable(concluido)

        try:
            self.unbind("<Return>")
        except Exception:
            pass

        try:
            # Em alguns ambientes/instalações do Tk no Windows, destruir a árvore
            # inteira desta tela de boas-vindas pode travar a inicialização.
            # Como ela é usada apenas uma vez por execução, basta removê-la do layout.
            bg.pack_forget()
        except Exception:
            pass

        self.protocol("WM_DELETE_WINDOW", self._on_close_app)
        if self._startup_cancelled:
            return ""
        return escolhido["nome"] or "Carol Carmo"

    # ---------------------------
    # UI / Style
    # ---------------------------
    def _build_style(self):
        t = THEME
        style = ttk.Style(self)
        style.theme_use("clam")

        self.configure(bg=t["bg_root"])
        style.configure("App.TFrame", background=t["bg_root"])
        style.configure("Surface.TFrame", background=t["bg_surface"])
        style.configure("Card.TFrame", background=t["bg_card"])

        style.configure(
            "Header.TLabel",
            background=t["bg_root"],
            foreground=t["fg_primary"],
            font=t["font_header"],
        )
        style.configure(
            "Sub.TLabel",
            background=t["bg_root"],
            foreground=t["fg_secondary"],
            font=t["font_sub"],
        )
        style.configure(
            "Section.TLabel",
            background=t["bg_surface"],
            foreground=t["fg_section"],
            font=t["font_section"],
        )
        style.configure(
            "Action.TButton",
            font=t["font_button"],
            padding=(12, 10),
            background=t["bg_button"],
            foreground=t["fg_primary"],
        )
        style.map(
            "Action.TButton",
            background=[("active", t["bg_button_hover"]), ("disabled", t["bg_disabled"])],
            foreground=[("disabled", t["fg_disabled"])],
        )
        style.configure(
            "Ghost.TButton",
            font=t["font_ghost"],
            padding=(10, 8),
            background=t["bg_surface"],
            foreground=t["fg_section"],
        )
        style.map("Ghost.TButton", background=[("active", t["bg_button"])])
        style.configure(
            "Status.TLabel",
            background=t["bg_surface"],
            foreground=t["fg_status"],
            font=t["font_sub"],
        )
        style.configure(
            "StatusReady.TLabel",
            background=t["bg_surface"],
            foreground=t["accent_ok"],
            font=("Segoe UI", 11, "bold"),
        )
        style.configure(
            "StatusBusy.TLabel",
            background=t["bg_surface"],
            foreground=t["accent_warn"],
            font=("Segoe UI", 11, "bold"),
        )
        style.configure("Thin.TSeparator", background=t["separator"])
        style.configure(
            "Accent.Horizontal.TProgressbar",
            troughcolor=t["bg_button"],
            background="#38bdf8",
            bordercolor=t["bg_button"],
            lightcolor="#38bdf8",
            darkcolor="#38bdf8",
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
        ttk.Label(
            header,
            text=f"Usuário ativo: {self._usuario_nome}",
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

        ttk.Label(sidebar, text=f"Usuário: {self._usuario_nome}", style="Sub.TLabel").pack(
            anchor="w", padx=14, pady=(0, 8)
        )

        self._menu_buttons_frame = ttk.Frame(sidebar, style="Card.TFrame")
        self._menu_buttons_frame.pack(fill="both", expand=True)
        self._render_menu_buttons()

        tools = ttk.Frame(sidebar, style="Card.TFrame")
        tools.pack(side="bottom", fill="x", padx=14, pady=(0, 12))

        ttk.Separator(sidebar, orient="horizontal", style="Thin.TSeparator").pack(
            side="bottom", fill="x", padx=14, pady=(10, 10)
        )

        self._btn_generate_excel = ttk.Button(
            tools,
            text="Executar",
            style="Ghost.TButton",
            command=self._run_pending_action,
            state="disabled",
        )
        self._btn_generate_excel.pack(fill="x", pady=4)

        # Main (log)
        main = ttk.Frame(body, style="Card.TFrame")
        main.pack(side="left", fill="both", expand=True, padx=(12, 0), ipadx=10, ipady=10)

        top_row = ttk.Frame(main, style="Card.TFrame")
        top_row.pack(fill="x", padx=12, pady=(12, 8))

        ttk.Label(top_row, text="Saída / Log", style="Section.TLabel").pack(side="left")

        self._status = ttk.Label(top_row, text="Pronto", style="StatusReady.TLabel")
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
            bg=THEME["bg_root"],
            fg=THEME["fg_primary"],
            insertbackground=THEME["fg_primary"],
            relief="flat",
            font=THEME["font_log"],
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
            text="<",
            style="Ghost.TButton",
            command=self._prev_preview_page,
        )
        self._btn_prev_page.pack(side="left", padx=(0, 6))

        self._btn_next_page = ttk.Button(
            preview_toolbar,
            text=">",
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
            bg="#081630",
            highlightthickness=0,
        )
        self._preview_canvas.pack(side="left", fill="both", expand=True)
        self._preview_canvas.bind("<Configure>", self._on_preview_resize)
        self._preview_canvas.bind("<MouseWheel>", self._on_preview_mousewheel)
        self._preview_canvas.bind("<Shift-MouseWheel>", self._on_preview_mousewheel_horizontal)
        self._preview_canvas.bind("<ButtonPress-1>", self._on_preview_drag_start)
        self._preview_canvas.bind("<B1-Motion>", self._on_preview_drag_move)

        preview_scroll = ttk.Scrollbar(
            preview_wrap,
            orient="vertical",
            command=self._preview_canvas.yview,
        )
        preview_scroll.pack(side="right", fill="y")
        self._preview_canvas.configure(yscrollcommand=preview_scroll.set)

        self._preview_hscroll = ttk.Scrollbar(
            preview_wrap,
            orient="horizontal",
            command=self._preview_canvas.xview,
        )
        self._preview_canvas.configure(xscrollcommand=self._preview_hscroll.set)

        self._preview_placeholder = self._preview_canvas.create_text(
            8, 8, anchor="nw", text="Sem pre-visualizacao",
            fill=THEME["fg_secondary"], font=THEME["font_sub"],
        )

        # Status bar bottom (progress)
        bottom = ttk.Frame(main, style="Card.TFrame")
        bottom.pack(fill="x", padx=12, pady=(0, 12))

        self._progress = ttk.Progressbar(
            bottom,
            mode="determinate",
            maximum=100,
            style="Accent.Horizontal.TProgressbar",
        )
        self._progress.configure(value=0)
        self._progress.pack(fill="x")

    def _render_menu_buttons(self):
        if not self._menu_buttons_frame:
            return

        for child in self._menu_buttons_frame.winfo_children():
            child.destroy()

        self._buttons = []

        if self._menu_modo == "criar_pasta":
            self._render_form_criar_pasta()
            return

        if self._menu_modo == "fazer_ponto":
            self._render_form_fazer_ponto()
            return

        if self._menu_modo == "desfazer_ponto":
            self._render_form_desfazer_ponto()
            return

        for item in self._menu_items():
            label = item["label"]
            label_norm = label.upper()
            locked = any(chave in label_norm for chave in self._LOCKED_MENU_KEYWORDS)

            texto = f"{label}  🔒" if locked else label
            if item.get("submenu"):
                texto = f"   {texto}"

            btn = ttk.Button(
                self._menu_buttons_frame,
                text=texto,
                style="Action.TButton",
                command=lambda i=item: self._handle_menu_action(i),
                state="disabled" if locked else "normal",
            )
            btn.pack(fill="x", padx=14, pady=6)

            if not locked:
                self._buttons.append(btn)

        if self._running:
            self._set_buttons_state("disabled")

    def _render_form_criar_pasta(self):
        frame = self._menu_buttons_frame

        ttk.Label(frame, text="Criar Pastas", style="Section.TLabel").pack(anchor="w", padx=14, pady=(2, 8))

        cp = CriarPasta()
        clientes = []
        try:
            clientes = cp.listar_clientes_modelos()
        except Exception as exc:
            self._write_log(f"Aviso ao carregar clientes: {exc}\n", tag="warn")

        ttk.Label(frame, text="Cliente", style="Sub.TLabel").pack(anchor="w", padx=14)
        cliente_var = tk.StringVar(value=clientes[0] if clientes else "")
        cliente_cb = ttk.Combobox(frame, textvariable=cliente_var, values=clientes, state="normal")
        cliente_cb.pack(fill="x", padx=14, pady=(2, 8))

        ttk.Label(frame, text="Navio", style="Sub.TLabel").pack(anchor="w", padx=14)
        navio_var = tk.StringVar()
        navio_entry = ttk.Entry(frame, textvariable=navio_var)
        navio_entry.pack(fill="x", padx=14, pady=(2, 8))

        ttk.Label(frame, text="DN", style="Sub.TLabel").pack(anchor="w", padx=14)
        dn_var = tk.StringVar()
        dn_entry = ttk.Entry(frame, textvariable=dn_var)
        dn_entry.pack(fill="x", padx=14, pady=(2, 12))

        def on_criar():
            cliente = cliente_var.get().strip()
            navio = navio_var.get().strip()
            dn = dn_var.get().strip()
            if not cliente or not navio or not dn:
                messagebox.showwarning("Dados incompletos", "Preencha cliente, navio e DN.", parent=self)
                return
            self._executar_criar_pasta(cliente, navio, dn)

        def _on_enter_criar(_event=None):
            on_criar()
            return "break"

        for widget in (cliente_cb, navio_entry, dn_entry):
            widget.bind("<Return>", _on_enter_criar)
            widget.bind("<KP_Enter>", _on_enter_criar)

        btn_criar = ttk.Button(frame, text="Criar Pasta", style="Action.TButton", command=on_criar)
        btn_criar.pack(fill="x", padx=14, pady=4)
        self._buttons.append(btn_criar)

        btn_voltar = ttk.Button(frame, text="Voltar", style="Ghost.TButton", command=self._voltar_menu_principal)
        btn_voltar.pack(fill="x", padx=14, pady=(4, 6))
        self._buttons.append(btn_voltar)

        cliente_cb.focus_set()

        if self._running:
            self._set_buttons_state("disabled")

    def _render_form_fazer_ponto(self):
        frame = self._menu_buttons_frame

        ttk.Label(frame, text="Adicionar Ponto", style="Section.TLabel").pack(anchor="w", padx=14, pady=(2, 8))

        if not self._fazer_ponto_ctx:
            self._fazer_ponto_form_vars = None
            ttk.Label(
                frame,
                text="Selecione um arquivo para continuar.",
                style="Sub.TLabel",
            ).pack(anchor="w", padx=14, pady=(0, 8))

            btn_arquivo = ttk.Button(frame, text="Selecionar Arquivo", style="Action.TButton", command=self._abrir_foco_fazer_ponto)
            btn_arquivo.pack(fill="x", padx=14, pady=4)
            self._buttons.append(btn_arquivo)

            btn_voltar = ttk.Button(frame, text="Voltar", style="Ghost.TButton", command=self._voltar_menu_principal)
            btn_voltar.pack(fill="x", padx=14, pady=(4, 6))
            self._buttons.append(btn_voltar)
            return

        datas = list(self._fazer_ponto_ctx.get("datas") or [])
        caminho_navio = str(self._fazer_ponto_ctx.get("caminho_navio") or "")

        ttk.Label(frame, text="Data", style="Sub.TLabel").pack(anchor="w", padx=14)
        data_var = tk.StringVar(value=datas[0] if datas else "")
        data_cb = ttk.Combobox(frame, textvariable=data_var, values=datas, state="normal")
        data_cb.pack(fill="x", padx=14, pady=(2, 2))

        ttk.Label(frame, text="Adicione mais datas caso necessário (DD/MM/AAAA).", style="Sub.TLabel").pack(
            anchor="w", padx=14, pady=(0, 8)
        )

        ttk.Label(frame, text="Período", style="Sub.TLabel").pack(anchor="w", padx=14)
        periodo_var = tk.StringVar(value="06h")
        periodo_cb = ttk.Combobox(frame, textvariable=periodo_var, values=["06h", "12h", "18h", "00h"], state="readonly")
        periodo_cb.pack(fill="x", padx=14, pady=(2, 12))

        self._fazer_ponto_form_vars = {
            "data_var": data_var,
            "periodo_var": periodo_var,
            "datas": datas,
            "caminho_navio": caminho_navio,
        }

        def on_preview(auto=False):
            data = data_var.get().strip()
            periodo = periodo_var.get().strip()
            if not data or not periodo:
                if not auto:
                    messagebox.showwarning("Dados incompletos", "Selecione data e período.", parent=self)
                return
            try:
                data = datetime.strptime(data, "%d/%m/%Y").strftime("%d/%m/%Y")
            except ValueError:
                if not auto:
                    messagebox.showwarning("Data inválida", "Informe a data no formato DD/MM/AAAA.", parent=self)
                return

            selecao = {
                "caminho_navio": caminho_navio,
                "datas": datas,
                "data": data,
                "periodo": periodo,
            }
            self._write_log(f"Adicionar Ponto: data={data} | periodo={periodo}\n", tag="info")
            self._clear_pending_action()

            def preview_action_inline(s=selecao):
                result = FazerPonto(debug=True).executar_preview(selection=s)
                if isinstance(result, dict):
                    selection_preview = result.get("selection") or {}
                    result["selection"] = {**selection_preview, **s}
                return result

            self._run_preview(
                preview_action=preview_action_inline,
                final_action=lambda selection=None: FazerPonto(debug=True).executar(selection=selection),
                label="Adicionar Ponto",
            )

        # Nao regenera preview a cada clique em data/periodo para evitar sobrecarga.
        # O preview e atualizado somente ao selecionar/trocar arquivo.

        btn_trocar = ttk.Button(frame, text="Trocar Arquivo", style="Ghost.TButton", command=self._abrir_foco_fazer_ponto)
        btn_trocar.pack(fill="x", padx=14, pady=(0, 0))
        self._buttons.append(btn_trocar)

        btn_voltar = ttk.Button(frame, text="Voltar", style="Ghost.TButton", command=self._voltar_menu_principal)
        btn_voltar.pack(fill="x", padx=14, pady=(4, 6))
        self._buttons.append(btn_voltar)

        data_cb.focus_set()

        # Sem botão de pré-visualização: ao entrar na tela já gera preview automaticamente.
        on_preview(auto=True)

        if self._running:
            self._set_buttons_state("disabled")

    def _render_form_desfazer_ponto(self):
        frame = self._menu_buttons_frame

        ttk.Label(frame, text="Remover Ponto", style="Section.TLabel").pack(anchor="w", padx=14, pady=(2, 8))

        if not self._desfazer_ponto_ctx:
            self._desfazer_ponto_form_vars = None
            ttk.Label(
                frame,
                text="Selecione um arquivo para continuar.",
                style="Sub.TLabel",
            ).pack(anchor="w", padx=14, pady=(0, 8))

            btn_arquivo = ttk.Button(frame, text="Selecionar Arquivo", style="Action.TButton", command=self._abrir_foco_desfazer_ponto)
            btn_arquivo.pack(fill="x", padx=14, pady=4)
            self._buttons.append(btn_arquivo)

            btn_voltar = ttk.Button(frame, text="Voltar", style="Ghost.TButton", command=self._voltar_menu_principal)
            btn_voltar.pack(fill="x", padx=14, pady=(4, 6))
            self._buttons.append(btn_voltar)
            return

        datas = list(self._desfazer_ponto_ctx.get("datas") or [])
        caminho_navio = str(self._desfazer_ponto_ctx.get("caminho_navio") or "")

        ttk.Label(frame, text="Data", style="Sub.TLabel").pack(anchor="w", padx=14)
        data_var = tk.StringVar(value=datas[0] if datas else "")
        data_cb = ttk.Combobox(frame, textvariable=data_var, values=datas, state="readonly")
        data_cb.pack(fill="x", padx=14, pady=(2, 8))

        ttk.Label(frame, text="Periodo", style="Sub.TLabel").pack(anchor="w", padx=14)
        periodo_var = tk.StringVar(value="06h")
        periodo_cb = ttk.Combobox(frame, textvariable=periodo_var, values=["06h", "12h", "18h", "00h"], state="readonly")
        periodo_cb.pack(fill="x", padx=14, pady=(2, 12))

        self._desfazer_ponto_form_vars = {
            "data_var": data_var,
            "periodo_var": periodo_var,
            "datas": datas,
            "caminho_navio": caminho_navio,
        }

        def on_preview(auto=False):
            data = data_var.get().strip()
            periodo = periodo_var.get().strip()
            if not data or not periodo:
                if not auto:
                    messagebox.showwarning("Dados incompletos", "Selecione data e periodo.", parent=self)
                return

            selecao = {
                "caminho_navio": caminho_navio,
                "datas": datas,
                "data": data,
                "periodo": periodo,
            }
            self._write_log(f"Remover Ponto: data={data} | periodo={periodo}\n", tag="info")
            self._clear_pending_action()

            def preview_action_inline(s=selecao):
                result = ProgramaRemoverPeriodo(debug=True).executar_preview(selection=s)
                if isinstance(result, dict):
                    selection_preview = result.get("selection") or {}
                    result["selection"] = {**selection_preview, **s}
                return result

            self._run_preview(
                preview_action=preview_action_inline,
                final_action=lambda selection=None: ProgramaRemoverPeriodo(debug=True).executar(selection=selection),
                label="Remover Ponto",
            )

        btn_trocar = ttk.Button(frame, text="Trocar Arquivo", style="Ghost.TButton", command=self._abrir_foco_desfazer_ponto)
        btn_trocar.pack(fill="x", padx=14, pady=(0, 0))
        self._buttons.append(btn_trocar)

        btn_voltar = ttk.Button(frame, text="Voltar", style="Ghost.TButton", command=self._voltar_menu_principal)
        btn_voltar.pack(fill="x", padx=14, pady=(4, 6))
        self._buttons.append(btn_voltar)

        data_cb.focus_set()

        on_preview(auto=True)

        if self._running:
            self._set_buttons_state("disabled")

    def _configure_tags(self):
        t = THEME
        self._log_text.tag_configure("info", foreground=t["fg_section"])
        self._log_text.tag_configure("ok", foreground=t["accent_ok"])
        self._log_text.tag_configure("err", foreground=t["accent_err"])
        self._log_text.tag_configure("warn", foreground=t["accent_warn"])

    # ---------------------------
    # Actions
    # ---------------------------
    def _menu_items(self):
        if self._menu_modo == "faturamento":
            return [
                self._item_faturamento_santos(),
                self._item_faturamento_atipico(),
                self._item_faturamento_sao_sebastiao(),
                {"label": "Voltar", "action": self._voltar_menu_principal},
            ]

        itens = [
            {"label": "Criar Pastas", "action": self._abrir_foco_criar_pasta, "menu_only": True},
            {
                "label": "Adicionar Ponto",
                "action": self._abrir_foco_fazer_ponto,
                "menu_only": True,
            },
            {
                "label": "Remover Ponto",
                "action": self._abrir_foco_desfazer_ponto,
                "menu_only": True,
            },
            {"label": "Faturamentos", "action": self._abrir_foco_faturamento, "menu_only": True},
            {
                "label": "De Acordo",
                "preview": lambda: FaturamentoDeAcordo(usuario_nome=self._usuario_nome).executar(preview=True),
                "action": lambda selection=None: FaturamentoDeAcordo(usuario_nome=self._usuario_nome).executar(
                    preview=False,
                    selection=selection,
                ),
            },
            {"label": "Relatorio", "action": self._relatorio_safe},
            {"label": "Gerador NFS-e", "action": self._abrir_gerador_nf},
        ]

        return itens

    def _abrir_foco_faturamento(self):
        self._menu_modo = "faturamento"
        self._render_menu_buttons()

    def _voltar_menu_principal(self):
        self._menu_modo = "principal"
        self._fazer_ponto_form_vars = None
        self._desfazer_ponto_form_vars = None
        self._render_menu_buttons()

    def _abrir_foco_criar_pasta(self):
        self._menu_modo = "criar_pasta"
        self._render_menu_buttons()

    def _obter_arquivo_e_datas(self, programa_cls):
        programa = programa_cls(debug=False)
        try:
            programa.abrir_arquivo_navio()
            if not programa.ws:
                return None

            programa.carregar_datas()
            datas = programa.datas or []
            if not datas:
                messagebox.showwarning("Sem dados", "Nenhuma data encontrada na planilha.", parent=self)
                return None

            caminho_base = getattr(programa, "_caminho_navio_destino", None)
            caminho_navio = str(caminho_base or programa.caminho_navio or "")
            return {
                "caminho_navio": caminho_navio,
                "datas": datas,
            }
        finally:
            try:
                fechar_workbooks(
                    app=programa.app,
                    wb_navio=programa.wb_navio,
                    wb_cliente=programa.wb_cliente,
                )
            except Exception:
                pass

    def _abrir_foco_fazer_ponto(self):
        ctx = self._obter_arquivo_e_datas(FazerPonto)
        if not ctx:
            if self._menu_modo not in ("fazer_ponto",):
                self._voltar_menu_principal()
            return
        self._fazer_ponto_ctx = ctx
        self._menu_modo = "fazer_ponto"
        self._render_menu_buttons()

    def _abrir_foco_desfazer_ponto(self):
        ctx = self._obter_arquivo_e_datas(ProgramaRemoverPeriodo)
        if not ctx:
            if self._menu_modo not in ("desfazer_ponto",):
                self._voltar_menu_principal()
            return
        self._desfazer_ponto_ctx = ctx
        self._menu_modo = "desfazer_ponto"
        self._render_menu_buttons()

    def _item_faturamento_santos(self):
        return {
            "label": "Santos",
            "preview": lambda: FaturamentoCompleto(usuario_nome=self._usuario_nome).executar(preview=True),
            "action": lambda selection=None: FaturamentoCompleto(usuario_nome=self._usuario_nome).executar(
                preview=False,
                selection=selection,
            ),
        }

    def _item_faturamento_atipico(self):
        return {
            "label": "Atipico",
            "preview": lambda: FaturamentoAtipico(usuario_nome=self._usuario_nome).executar(preview=True),
            "action": lambda selection=None: FaturamentoAtipico(usuario_nome=self._usuario_nome).executar(
                preview=False,
                selection=selection,
            ),
        }

    def _item_faturamento_sao_sebastiao(self):
        return {
            "label": "Sao Sebastiao",
            "preview": lambda: FaturamentoSaoSebastiao(usuario_nome=self._usuario_nome).executar(preview=True),
            "action": lambda selection=None: FaturamentoSaoSebastiao(usuario_nome=self._usuario_nome).executar(
                preview=False,
                selection=selection,
            ),
        }

    def _relatorio_safe(self):
        if hasattr(GerarRelatorio, "executar"):
            GerarRelatorio().executar()
        else:
            self._write_log("Relatório não implementado.\n", tag="warn")

    def _abrir_gerador_nf(self):
        """Abre o Gerador de NFS-e (GINFES Santos) em uma janela separada (PyQt6)."""
        path_gerador = project_root_path() / "backend" / "app" / "modules" / "gerador_nf.py"
        if not path_gerador.exists():
            self._write_log(f"Gerador NFS-e não encontrado: {path_gerador}\n", tag="err")
            messagebox.showerror("Erro", "Arquivo gerador_nf.py não encontrado.", parent=self)
            return
        try:
            subprocess.Popen(
                [sys.executable, "-m", "backend.app.modules.gerador_nf"],
                cwd=str(project_root_path()),
            )
            self._write_log("Gerador NFS-e aberto em nova janela.\n", tag="ok")
        except Exception as e:
            self._write_log(f"Erro ao abrir Gerador NFS-e: {e}\n", tag="err")
            messagebox.showerror(
                "Erro",
                f"Não foi possível abrir o Gerador NFS-e.\n\nVerifique se o PyQt6 está instalado:\npip install PyQt6",
                parent=self,
            )

    def _criar_pasta_ui(self):
        """Mantido por compatibilidade; redireciona para o modo de foco na lateral."""
        self._abrir_foco_criar_pasta()

    def _executar_criar_pasta(self, cliente, navio, dn):

        self._write_log(
            f"Criar pasta: cliente={cliente} | navio={navio} | DN={dn}\n",
            tag="info",
        )

        # Define a ação que será executada na thread
        def action():
            cp = CriarPasta()
            info = cp.executar(
                cliente=cliente,
                navio=navio,
                dn=dn,
                return_info=True,
                log_callback=self._write_log,
            )
            
            destino = info["destino"]
            
            # Verificação final
            if destino.exists():
                self._write_log("Confirmado: pasta existe no sistema\n", tag="ok")
            else:
                self._write_log("AVISO: Pasta nao encontrada apos criacao!\n", tag="err")

        # Executa a ação na thread
        self._run_action(action)

    def _pedir_dados_criar_pasta(self):
        dialog = tk.Toplevel(self)
        dialog.title("Criar Pasta")
        dialog.configure(bg=THEME["bg_root"])
        dialog.resizable(False, False)
        dialog.grab_set()

        frame = ttk.Frame(dialog, style="Card.TFrame")
        frame.pack(padx=16, pady=14, fill="both", expand=True)

        cp = CriarPasta()
        clientes = []
        try:
            clientes = cp.listar_clientes_modelos()
        except Exception as exc:
            messagebox.showwarning(
                "Aviso",
                f"Nao foi possivel carregar lista de clientes dos modelos:\n{exc}",
                parent=dialog,
            )

        # Cliente (lista de modelos FATURAMENTOS)
        ttk.Label(frame, text="Cliente", style="Section.TLabel").grid(row=0, column=0, sticky="w", pady=(0, 4))
        cliente_var = tk.StringVar(value=clientes[0] if clientes else "")
        cliente_cb = ttk.Combobox(frame, textvariable=cliente_var, values=clientes, state="normal", width=38)
        cliente_cb.grid(row=1, column=0, sticky="ew", pady=(0, 6))

        ttk.Label(
            frame,
            text=f"Lista rapida carregada: {len(clientes)} clientes.",
            style="Sub.TLabel",
        ).grid(row=2, column=0, sticky="w", pady=(0, 8))

        # Nome do navio
        ttk.Label(frame, text="Nome do navio", style="Section.TLabel").grid(row=3, column=0, sticky="w", pady=(0, 4))
        navio_var = tk.StringVar()
        navio_entry = ttk.Entry(frame, textvariable=navio_var, width=40)
        navio_entry.grid(row=4, column=0, sticky="ew", pady=(0, 10))

        # DN (obrigatorio)
        ttk.Label(frame, text="DN (obrigatorio)", style="Section.TLabel").grid(row=5, column=0, sticky="w", pady=(0, 4))
        dn_var = tk.StringVar()
        dn_entry = ttk.Entry(frame, textvariable=dn_var, width=40)
        dn_entry.grid(row=6, column=0, sticky="ew", pady=(0, 4))

        ttk.Label(
            frame,
            text="A pasta sera criada na rede dentro do cliente selecionado.",
            style="Sub.TLabel",
        ).grid(row=7, column=0, sticky="w", pady=(0, 12))

        buttons = ttk.Frame(frame, style="Card.TFrame")
        buttons.grid(row=8, column=0, sticky="e")

        resultado = {"valor": None}

        def on_ok():
            cliente = cliente_var.get().strip()
            navio = navio_var.get().strip()
            dn = dn_var.get().strip()

            if not cliente or not navio or not dn:
                messagebox.showwarning("Dados incompletos", "Preencha cliente, navio e DN.")
                return

            resultado["valor"] = (cliente, navio, dn)
            dialog.destroy()

        def on_cancel():
            dialog.destroy()

        # Bind da tecla Enter para acionar OK
        dialog.bind('<Return>', lambda event: on_ok())

        ttk.Button(buttons, text="Cancelar", style="Ghost.TButton", command=on_cancel).pack(side="right", padx=(8, 0))
        ttk.Button(buttons, text="OK", style="Action.TButton", command=on_ok).pack(side="right")

        cliente_cb.focus_set()
        dialog.wait_window()
        return resultado["valor"]

    def _pedir_dados_periodo(self, programa_cls, titulo, apenas_arquivo=False):
        programa = programa_cls(debug=False)

        try:
            programa.abrir_arquivo_navio()
            if not programa.ws:
                return None

            programa.carregar_datas()
            datas = programa.datas or []
            if not datas:
                messagebox.showwarning("Sem dados", "Nenhuma data encontrada na planilha.", parent=self)
                return None

            caminho_base = getattr(programa, "_caminho_navio_destino", None)
            caminho_navio = str(caminho_base or programa.caminho_navio or "")
            if apenas_arquivo:
                return {
                    "caminho_navio": caminho_navio,
                    "datas": datas,
                }
        finally:
            try:
                fechar_workbooks(
                    app=programa.app,
                    wb_navio=programa.wb_navio,
                    wb_cliente=programa.wb_cliente,
                )
            except Exception:
                pass

        dialog = tk.Toplevel(self)
        dialog.title(titulo)
        dialog.configure(bg=THEME["bg_root"])
        dialog.resizable(False, False)
        dialog.grab_set()

        frame = ttk.Frame(dialog, style="Card.TFrame")
        frame.pack(padx=16, pady=14, fill="both", expand=True)

        ttk.Label(frame, text="Data", style="Section.TLabel").grid(row=0, column=0, sticky="w", pady=(0, 4))
        data_var = tk.StringVar(value=datas[0])
        data_cb = ttk.Combobox(frame, textvariable=data_var, values=datas, state="readonly", width=36)
        data_cb.grid(row=1, column=0, sticky="ew", pady=(0, 12))

        ttk.Label(frame, text="Período", style="Section.TLabel").grid(row=2, column=0, sticky="w", pady=(0, 4))
        periodos = ["06h", "12h", "18h", "00h"]
        periodo_var = tk.StringVar(value=periodos[0])
        periodo_cb = ttk.Combobox(frame, textvariable=periodo_var, values=periodos, state="readonly", width=36)
        periodo_cb.grid(row=3, column=0, sticky="ew", pady=(0, 12))

        buttons = ttk.Frame(frame, style="Card.TFrame")
        buttons.grid(row=4, column=0, sticky="e")

        resultado = {"valor": None}

        def on_ok():
            data = data_var.get().strip()
            periodo = periodo_var.get().strip()
            if not data or not periodo:
                messagebox.showwarning("Dados incompletos", "Selecione data e período.", parent=dialog)
                return

            resultado["valor"] = {
                "data": data,
                "periodo": periodo,
                "caminho_navio": caminho_navio,
            }
            dialog.destroy()

        def on_cancel():
            dialog.destroy()

        dialog.bind('<Return>', lambda _event: on_ok())

        ttk.Button(buttons, text="Cancelar", style="Ghost.TButton", command=on_cancel).pack(side="right", padx=(8, 0))
        ttk.Button(buttons, text="OK", style="Action.TButton", command=on_ok).pack(side="right")

        data_cb.focus_set()
        dialog.wait_window()
        return resultado["valor"]

    def _pedir_data_periodo_com_datas(self, datas, titulo):
        if not datas:
            return None

        dialog = tk.Toplevel(self)
        dialog.title(titulo)
        dialog.configure(bg=THEME["bg_root"])
        dialog.resizable(False, False)
        dialog.grab_set()

        frame = ttk.Frame(dialog, style="Card.TFrame")
        frame.pack(padx=16, pady=14, fill="both", expand=True)

        ttk.Label(frame, text="Data", style="Section.TLabel").grid(row=0, column=0, sticky="w", pady=(0, 4))
        data_var = tk.StringVar(value=datas[0])
        data_cb = ttk.Combobox(frame, textvariable=data_var, values=datas, state="normal", width=36)
        data_cb.grid(row=1, column=0, sticky="ew", pady=(0, 2))

        ttk.Label(
            frame,
            text="Adicione mais datas caso necessário (DD/MM/AAAA).",
            style="Sub.TLabel",
        ).grid(row=2, column=0, sticky="w", pady=(0, 10))

        ttk.Label(frame, text="Período", style="Section.TLabel").grid(row=3, column=0, sticky="w", pady=(0, 4))
        periodos = ["06h", "12h", "18h", "00h"]
        periodo_var = tk.StringVar(value=periodos[0])
        periodo_cb = ttk.Combobox(frame, textvariable=periodo_var, values=periodos, state="readonly", width=36)
        periodo_cb.grid(row=4, column=0, sticky="ew", pady=(0, 12))

        buttons = ttk.Frame(frame, style="Card.TFrame")
        buttons.grid(row=5, column=0, sticky="e")

        resultado = {"valor": None}

        def on_ok():
            data = data_var.get().strip()
            periodo = periodo_var.get().strip()
            if not data or not periodo:
                messagebox.showwarning("Dados incompletos", "Selecione data e período.", parent=dialog)
                return

            try:
                data = datetime.strptime(data, "%d/%m/%Y").strftime("%d/%m/%Y")
            except ValueError:
                messagebox.showwarning(
                    "Data invalida",
                    "Informe a data no formato DD/MM/AAAA.",
                    parent=dialog,
                )
                return

            resultado["valor"] = {
                "data": data,
                "periodo": periodo,
            }
            dialog.destroy()

        def on_cancel():
            dialog.destroy()

        dialog.bind('<Return>', lambda _event: on_ok())

        ttk.Button(buttons, text="Cancelar", style="Ghost.TButton", command=on_cancel).pack(side="right", padx=(8, 0))
        ttk.Button(buttons, text="OK", style="Action.TButton", command=on_ok).pack(side="right")

        data_cb.focus_set()
        dialog.wait_window()
        return resultado["valor"]

    def _handle_menu_action(self, item):
        # Casos especiais que precisam coletar dados antes de executar
        if item.get("menu_only"):
            item["action"]()
            return

        if item["label"] == "Adicionar Ponto":
            selecao = self._pedir_dados_periodo(FazerPonto, item["label"], apenas_arquivo=True)
            if not selecao:
                self._set_status("Operacao cancelada", busy=False)
                return

            self._write_log(
                f"{item['label']}: arquivo={Path(selecao['caminho_navio']).name}\n",
                tag="info",
            )
            self._clear_pending_action()

            self._run_preview(
                preview_action=lambda s=selecao: FazerPonto(debug=True).executar_preview(selection=s),
                final_action=item["action"],
                label=item["label"],
            )
            return

        if item["label"] == "Remover Ponto":
            selecao = self._pedir_dados_periodo(ProgramaRemoverPeriodo, item["label"])
            if not selecao:
                self._set_status("Operacao cancelada", busy=False)
                return

            self._write_log(
                f"{item['label']}: data={selecao['data']} | período={selecao['periodo']}\n",
                tag="info",
            )
            self._clear_pending_action()
            self._run_action(lambda: item["action"](selecao))
            return
        
        if item.get("preview"):
            self._run_preview(item["preview"], item["action"], item["label"])
        else:
            self._clear_pending_action()
            self._run_action(item["action"])

    def _run_action(self, action, pending_context=None):
        if self._running:
            return

        self._running = True
        self._set_buttons_state("disabled")
        self._set_status("Executando...", busy=True)
        self._start_loading_progress()

        self._write_log("\nIniciando...\n", tag="info")

        def job():
            rearm_pending = False
            try:
                result = action()
                if isinstance(result, dict):
                    info_msg = str(result.get("info") or "").strip()
                    if info_msg:
                        self._write_log(info_msg + "\n", tag="info")

                    msg = str(result.get("message") or "").strip()
                    if msg:
                        tag = "ok" if bool(result.get("changed", True)) else "warn"
                        self._write_log(msg + "\n", tag=tag)
                        if (not bool(result.get("changed", True))) and ("ja existe" in msg.lower()):
                            self._log_queue.put(("__POPUP_WARN__", msg))
                            rearm_pending = True
                elif isinstance(result, str) and result.strip():
                    self._write_log(result.strip() + "\n", tag="info")
                self._write_log("\nConcluido.\n", tag="ok")
            except Exception as exc:
                import traceback
                self._write_log(f"\nERRO: {exc}\n", tag="err")
                self._write_log(f"Traceback:\n{traceback.format_exc()}\n", tag="err")
                rearm_pending = True
            finally:
                self._log_queue.put(("__DONE__", None))
                if rearm_pending and pending_context:
                    self._log_queue.put(("__REARM_PENDING__", pending_context))

        threading.Thread(target=job, daemon=True).start()

    def _run_preview(self, preview_action, final_action, label):
        if self._running:
            return

        self._preview_mode_hint = str(label or "")
        self._running = True
        self._set_buttons_state("disabled")
        self._set_status("Gerando pre-visualizacao...", busy=True)
        self._start_loading_progress()

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
                self._write_log("[PREVIEW] Pronto. Clique em 'Executar'.\n", tag="ok")
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
        self._set_status("Pronto", busy=False)
        self._clear_pending_action()
        self._clear_preview()

    def _copy_log(self):
        txt = self._log_text.get("1.0", "end-1c")
        self.clipboard_clear()
        self.clipboard_append(txt)
        self._set_status("Log copiado", busy=False)

    def _set_idle(self):
        self._running = False
        self._set_buttons_state("normal")
        self._finish_loading_progress()
        self._set_status("Pronto", busy=False)

    def _set_status(self, text, busy=False):
        style = "StatusBusy.TLabel" if busy else "StatusReady.TLabel"
        self._status.configure(text=text, style=style)

    def _start_loading_progress(self):
        self._cancel_loading_progress_job()
        self._progress_value = 0.0
        self._progress.configure(value=0)
        self._schedule_loading_progress_tick()

    def _schedule_loading_progress_tick(self):
        if not self._running:
            return

        # Avanco mais lento para processos longos (faturamento).
        faltante = max(92.0 - self._progress_value, 0.0)
        passo = max(0.03, faltante * 0.008)
        self._progress_value = min(92.0, self._progress_value + passo)
        self._progress.configure(value=self._progress_value)
        self._progress_job = self.after(220, self._schedule_loading_progress_tick)

    def _finish_loading_progress(self):
        self._cancel_loading_progress_job()
        self._progress_value = 100.0
        self._progress.configure(value=100)

        def _reset_bar():
            self._progress_value = 0.0
            self._progress.configure(value=0)

        self.after(220, _reset_bar)

    def _cancel_loading_progress_job(self):
        if self._progress_job is not None:
            try:
                self.after_cancel(self._progress_job)
            except Exception:
                pass
            self._progress_job = None

    def _set_pending_action(self, label, action, selection=None):
        self._pending_action = action
        self._pending_label = label
        self._pending_selection = selection
        self._btn_generate_excel.configure(state="normal")
        self._set_status("Pre-visualizacao pronta", busy=False)

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

        if self._pending_label == "Adicionar Ponto":
            form = self._fazer_ponto_form_vars or {}
            datas = list(form.get("datas") or (selection or {}).get("datas") or [])

            domingos = []
            for d in datas:
                try:
                    if datetime.strptime(str(d).strip(), "%d/%m/%Y").weekday() == 6:
                        domingos.append(d)
                except Exception:
                    continue

            if domingos:
                if len(domingos) <= 4:
                    msg = f"Este arquivo possui domingo(s): {', '.join(domingos)}."
                else:
                    msg = (
                        f"Este arquivo possui {len(domingos)} domingo(s), "
                        f"incluindo {', '.join(domingos[:4])}."
                    )
                messagebox.showwarning(
                    "Atenção: domingos no arquivo",
                    msg,
                    parent=self,
                )

            data_val = str((form.get("data_var").get() if form.get("data_var") else (selection or {}).get("data") or "")).strip()
            periodo_val = str((form.get("periodo_var").get() if form.get("periodo_var") else (selection or {}).get("periodo") or "")).strip()

            if not data_val or not periodo_val:
                messagebox.showwarning("Dados incompletos", "Selecione data e período.", parent=self)
                self._set_status("Operacao cancelada", busy=False)
                return

            try:
                data_val = datetime.strptime(data_val, "%d/%m/%Y").strftime("%d/%m/%Y")
            except ValueError:
                messagebox.showwarning("Data inválida", "Informe a data no formato DD/MM/AAAA.", parent=self)
                self._set_status("Operacao cancelada", busy=False)
                return

            selection = {
                **(selection or {}),
                "caminho_navio": str(form.get("caminho_navio") or (selection or {}).get("caminho_navio") or ""),
                "datas": datas,
                "data": data_val,
                "periodo": periodo_val,
            }
            self._write_log(
                f"Adicionar Ponto: data={selection['data']} | periodo={selection['periodo']}\n",
                tag="info",
            )

        if self._pending_label == "Remover Ponto":
            form = self._desfazer_ponto_form_vars or {}

            data_val = str((form.get("data_var").get() if form.get("data_var") else (selection or {}).get("data") or "")).strip()
            periodo_val = str((form.get("periodo_var").get() if form.get("periodo_var") else (selection or {}).get("periodo") or "")).strip()

            if not data_val or not periodo_val:
                messagebox.showwarning("Dados incompletos", "Selecione data e periodo.", parent=self)
                self._set_status("Operacao cancelada", busy=False)
                return

            selection = {
                **(selection or {}),
                "caminho_navio": str(form.get("caminho_navio") or (selection or {}).get("caminho_navio") or ""),
                "datas": list(form.get("datas") or (selection or {}).get("datas") or []),
                "data": data_val,
                "periodo": periodo_val,
            }
            self._write_log(
                f"Remover Ponto: data={selection['data']} | periodo={selection['periodo']}\n",
                tag="info",
            )

        pending_label = self._pending_label or ""
        self._clear_pending_action()
        self._run_action(
            lambda: action(selection),
            pending_context={
                "label": pending_label,
                "action": action,
                "selection": selection,
            },
        )

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

                if msg == "__POPUP_WARN__":
                    messagebox.showwarning("Aviso", str(tag), parent=self)
                    continue

                if msg == "__REARM_PENDING__":
                    ctx = tag or {}
                    self._set_pending_action(
                        ctx.get("label") or "",
                        ctx.get("action"),
                        ctx.get("selection"),
                    )
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

    def _on_preview_mousewheel_horizontal(self, event):
        if not self._preview_pages:
            return
        delta = int(-1 * (event.delta / 120))
        self._preview_canvas.xview_scroll(delta, "units")

    def _on_preview_drag_start(self, event):
        if not self._preview_pages:
            return
        self._preview_canvas.scan_mark(event.x, event.y)

    def _on_preview_drag_move(self, event):
        if not self._preview_pages:
            return
        self._preview_canvas.scan_dragto(event.x, event.y, gain=1)

    def _render_preview_page(self):
        if not self._preview_pages:
            return

        width = max(self._preview_canvas.winfo_width(), 1)
        height = max(self._preview_canvas.winfo_height(), 1)
        page = self._preview_pages[self._preview_page_index]
        img_width, img_height = page.size

        fit_width_scale = (width / img_width)
        scale = fit_width_scale * self._preview_zoom

        new_size = (
            max(int(img_width * scale), 1),
            max(int(img_height * scale), 1),
        )
        resized = page.resize(new_size, Image.LANCZOS)

        self._preview_image = ImageTk.PhotoImage(resized)
        self._preview_canvas.delete("preview_img")
        pos_x = max((width - resized.size[0]) // 2, 0)
        pos_y = max((height - resized.size[1]) // 2, 0)
        self._preview_canvas.create_image(
            pos_x,
            pos_y,
            image=self._preview_image,
            anchor="nw",
            tags="preview_img",
        )

        self._preview_canvas.configure(
            scrollregion=(
                0,
                0,
                max(resized.size[0], width),
                max(resized.size[1], height),
            )
        )

        if self._preview_reset_scroll:
            self._preview_canvas.xview_moveto(0)
            self._preview_canvas.yview_moveto(0)
            self._preview_reset_scroll = False

        if self._preview_placeholder:
            self._preview_canvas.itemconfigure(self._preview_placeholder, state="hidden")

        self._update_preview_nav()

    def _set_preview_zoom(self, zoom):
        self._preview_zoom = max(0.3, min(self._preview_zoom_max, zoom))
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

            eh_fazer_ponto = "FAZER PONTO" in self._preview_mode_hint.upper()
            dpi_preview = 300 if eh_fazer_ponto else 200

            pages = None
            erros = []

            try:
                pages = convert_from_path(str(path), dpi=dpi_preview)
            except Exception as exc:
                erros.append(str(exc))

            if not pages:
                for poppler_dir in self._poppler_paths_candidatos():
                    try:
                        pages = convert_from_path(
                            str(path),
                            dpi=dpi_preview,
                            poppler_path=str(poppler_dir),
                        )
                        if pages:
                            break
                    except Exception as exc:
                        erros.append(str(exc))

            if not pages:
                self._write_log(
                    "Preview interno indisponivel neste computador (Poppler nao encontrado). Abrindo PDF externamente.\n",
                    tag="warn",
                )
                if erros:
                    self._write_log(f"Detalhe tecnico: {erros[-1]}\n", tag="warn")
                self._abrir_pdf_externo(path)
                return

            self._preview_pages = pages
            # O preview remove margens externas para visualizacao mais fiel ao conteudo.
            self._preview_pages = [
                self._trim_preview_image_whitespace(p) for p in self._preview_pages
            ]
            self._preview_page_index = 0
            self._preview_pdf_path = path
            self._preview_reset_scroll = True

            if eh_fazer_ponto:
                self._preview_zoom_max = 4.5

                # Adicionar Ponto: une paginas em uma imagem vertical unica.
                if len(self._preview_pages) > 1:
                    merged = self._merge_preview_pages_vertically(self._preview_pages)
                    self._preview_pages = [self._trim_preview_image_whitespace(merged)]
                    self._preview_page_index = 0

                # Conteudo mais proximo e navegacao horizontal por arraste.
                self._preview_zoom = 2.15
                self._preview_hscroll.pack_forget()
            else:
                self._preview_zoom_max = 2.5
                self._preview_zoom = 1.0
                self._preview_hscroll.pack_forget()

            self._render_preview_page()

            if eh_fazer_ponto:
                # Garante inicio no canto esquerdo apos o layout final do canvas.
                self.after(0, lambda: self._preview_canvas.xview_moveto(0))

            self._notebook.select(1)
        except Exception as exc:
            self._write_log(f"Falha ao renderizar PDF: {exc}\n", tag="warn")

    def _merge_preview_pages_vertically(self, pages):
        if not pages:
            raise ValueError("Lista de paginas vazia")
        if len(pages) == 1:
            return pages[0]

        largura = max(p.width for p in pages)
        altura_total = sum(p.height for p in pages)
        merged = Image.new("RGB", (largura, altura_total), "white")

        y = 0
        for page in pages:
            x = max((largura - page.width) // 2, 0)
            merged.paste(page, (x, y))
            y += page.height

        return merged

    def _trim_preview_image_whitespace(self, image):
        """
        Remove bordas brancas do preview para aproximar o conteudo util.
        """
        try:
            gray = image.convert("L")
            # Considera branco acima de ~245 como fundo.
            mask = gray.point(lambda p: 255 if p < 245 else 0)
            bbox = mask.getbbox()
            if not bbox:
                return image

            left, top, right, bottom = bbox
            pad = 18
            left = max(0, left - pad)
            top = max(0, top - pad)
            right = min(image.width, right + pad)
            bottom = min(image.height, bottom + pad)
            return image.crop((left, top, right, bottom))
        except Exception:
            return image

    def _poppler_paths_candidatos(self):
        return poppler_paths_candidatos()

    def _abrir_pdf_externo(self, path: Path):
        try:
            if hasattr(os, "startfile"):
                os.startfile(str(path))
            else:
                subprocess.Popen(["xdg-open", str(path)])
        except Exception as exc:
            self._write_log(f"Nao foi possivel abrir o PDF externamente: {exc}\n", tag="warn")

    def _clear_preview(self):
        self._preview_pil = None
        self._preview_image = None
        self._preview_pdf_path = None
        self._preview_pages = []
        self._preview_mode_hint = ""
        self._preview_reset_scroll = False
        self._preview_zoom = 1.0
        self._preview_zoom_max = 2.5
        self._preview_page_index = 0
        self._preview_canvas.delete("preview_img")
        self._preview_hscroll.pack_forget()
        self._preview_canvas.xview_moveto(0)
        if self._preview_placeholder:
            self._preview_canvas.itemconfigure(self._preview_placeholder, state="normal")

def run_desktop():
    app = DesktopApp()
    app.mainloop()


if __name__ == "__main__":
    run_desktop()


