"""
Desktop Yuta – Central de Processos.
Interface Tkinter para faturamentos, ponto e relatórios.
"""
import queue
import subprocess
import sys
import threading
import tkinter as tk
import os
from pathlib import Path
from tkinter import messagebox, ttk

from pdf2image import convert_from_path
from PIL import Image, ImageTk
from yuta_helpers import fechar_workbooks

from classes import (
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
    def __init__(self):
        super().__init__()
        self.title("Yuta - Central de Processos")
        self._center_window(self, 1280, 900)
        self.minsize(980, 640)
        self.protocol("WM_DELETE_WINDOW", self._on_close_app)

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
        self._startup_cancelled = False

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
        dialog = tk.Toplevel(self)
        dialog.title("Yuta · Iniciar sessão")
        dialog.configure(bg="#060b1a")
        dialog.resizable(False, False)
        dialog.transient(self)
        dialog.grab_set()

        largura = 700
        altura = 520
        self._center_window(dialog, largura, altura)
        dialog.minsize(largura, altura)

        # Fundo
        bg = tk.Frame(dialog, bg="#060b1a")
        bg.pack(fill="both", expand=True)

        # Card central
        card = tk.Frame(bg, bg="#0f172a", highlightthickness=1, highlightbackground="#1f2a44")
        card.place(relx=0.5, rely=0.5, anchor="center", width=520, height=420)

        header = tk.Frame(card, bg="#0f172a")
        header.pack(fill="x", padx=24, pady=(20, 8))

        tk.Label(
            header,
            text="✓ Yuta",
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
                text="○",
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
                    text="◉" if ativo else "○",
                    bg=fundo,
                    fg="#3b82f6" if ativo else "#93a4c2",
                )
                opt["nome_lbl"].configure(bg=fundo)
                opt["avatar"].configure(bg="#f3f4f6" if ativo else "#d1d5db")

        criar_opcao("Carol Carmo", "CC")
        criar_opcao("Diogo Barros", "DB")
        atualizar_estilo_opcoes()

        escolhido = {"nome": "Carol Carmo"}

        def confirmar():
            escolhido["nome"] = usuario_var.get().strip() or "Carol Carmo"
            dialog.destroy()

        def fechar_app():
            self._startup_cancelled = True
            try:
                dialog.grab_release()
            except Exception:
                pass
            dialog.destroy()

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

        dialog.bind("<Return>", lambda _e: confirmar())
        dialog.protocol("WM_DELETE_WINDOW", fechar_app)
        self.protocol("WM_DELETE_WINDOW", fechar_app)
        dialog.focus_force()
        dialog.wait_window()
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
        style.configure("Thin.TSeparator", background=t["separator"])

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
            bg=THEME["bg_root"],
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
            8, 8, anchor="nw", text="Sem pre-visualizacao",
            fill=THEME["fg_secondary"], font=THEME["font_sub"],
        )

        # Status bar bottom (progress)
        bottom = ttk.Frame(main, style="Card.TFrame")
        bottom.pack(fill="x", padx=12, pady=(0, 12))

        self._progress = ttk.Progressbar(bottom, mode="indeterminate")
        self._progress.pack(fill="x")

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
        return [
            {
                "label": "🧾 Faturamento (Normal)",
                "preview": lambda: FaturamentoCompleto(usuario_nome=self._usuario_nome).executar(preview=True),
                "action": lambda selection=None: FaturamentoCompleto(usuario_nome=self._usuario_nome).executar(
                    preview=False,
                    selection=selection,
                ),
            },
            {
                "label": "🧩 Faturamento (Atípico)",
                "preview": lambda: FaturamentoAtipico(usuario_nome=self._usuario_nome).executar(preview=True),
                "action": lambda selection=None: FaturamentoAtipico(usuario_nome=self._usuario_nome).executar(
                    preview=False,
                    selection=selection,
                ),
            },
            {
                "label": "⚓ Faturamento São Sebastião",
                "preview": lambda: FaturamentoSaoSebastiao(usuario_nome=self._usuario_nome).executar(preview=True),
                "action": lambda selection=None: FaturamentoSaoSebastiao(usuario_nome=self._usuario_nome).executar(
                    preview=False,
                    selection=selection,
                ),
            },
            {
                "label": "✅ De Acordo",
                "preview": lambda: FaturamentoDeAcordo(usuario_nome=self._usuario_nome).executar(preview=True),
                "action": lambda selection=None: FaturamentoDeAcordo(usuario_nome=self._usuario_nome).executar(
                    preview=False,
                    selection=selection,
                ),
            },
            {
                "label": "🕒 Fazer Ponto",
                "action": lambda selection=None: FazerPonto(debug=True).executar(selection=selection),
            },
            {
                "label": "↩️ Desfazer Ponto",
                "action": lambda selection=None: ProgramaRemoverPeriodo(debug=True).executar(selection=selection),
            },
            {"label": "📊 Relatório", "action": self._relatorio_safe},
            {"label": "📁 Criar Pasta", "action": self._criar_pasta_ui},
            {"label": "📄 Gerador NFS-e", "action": self._abrir_gerador_nf},
        ]

    def _relatorio_safe(self):
        if hasattr(GerarRelatorio, "executar"):
            GerarRelatorio().executar()
        else:
            self._write_log("Relatório não implementado.\n", tag="warn")

    def _abrir_gerador_nf(self):
        """Abre o Gerador de NFS-e (GINFES Santos) em uma janela separada (PyQt6)."""
        path_gerador = Path(__file__).resolve().parent / "classes" / "gerador_nf.py"
        if not path_gerador.exists():
            self._write_log(f"Gerador NFS-e não encontrado: {path_gerador}\n", tag="err")
            messagebox.showerror("Erro", "Arquivo gerador_nf.py não encontrado.", parent=self)
            return
        try:
            subprocess.Popen(
                [sys.executable, str(path_gerador)],
                cwd=str(path_gerador.parent.parent),
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
        """Coleta os dados e inicia a criação da pasta"""
        dados = self._pedir_dados_criar_pasta()
        if not dados:
            return
            
        cliente, navio, dn = dados

        self._write_log(
            f"Criar pasta: cliente={cliente} | navio={navio} | DN={dn} (automático)\n",
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
                self._write_log(f"✓ Confirmado: pasta existe no sistema\n", tag="ok")
            else:
                self._write_log(f"⚠️ AVISO: Pasta não encontrada após criação!\n", tag="err")

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

        # Cliente
        ttk.Label(frame, text="Cliente", style="Section.TLabel").grid(row=0, column=0, sticky="w", pady=(0, 4))
        
        # Combobox editável desde o início (permite digitação imediata)
        cliente_var = tk.StringVar()
        cliente_cb = ttk.Combobox(frame, textvariable=cliente_var, state="normal", width=38)
        cliente_cb.grid(row=1, column=0, sticky="ew", pady=(0, 10))
        
        # Label de status para indicar carregamento
        status_label = ttk.Label(frame, text="Carregando clientes...", style="Sub.TLabel")
        status_label.grid(row=2, column=0, sticky="w", pady=(0, 4))
        
        # Carrega clientes em background
        cp = CriarPasta()
        def carregar_clientes_bg():
            try:
                clientes = cp.listar_clientes()
                # Atualiza o combobox na thread principal
                dialog.after(0, lambda: _atualizar_combo(clientes))
            except Exception as e:
                dialog.after(0, lambda: status_label.configure(text=f"⚠️ Erro ao listar: {e}"))
        
        def _atualizar_combo(clientes):
            cliente_cb['values'] = clientes
            if clientes:
                status_label.configure(text=f"✓ {len(clientes)} clientes encontrados")
            else:
                status_label.configure(text="Nenhum cliente encontrado")
        
        # Inicia carregamento em thread separada
        threading.Thread(target=carregar_clientes_bg, daemon=True).start()

        # Nome do navio
        ttk.Label(frame, text="Nome do navio", style="Section.TLabel").grid(row=3, column=0, sticky="w", pady=(0, 4))
        navio_var = tk.StringVar()
        navio_entry = ttk.Entry(frame, textvariable=navio_var, width=40)
        navio_entry.grid(row=4, column=0, sticky="ew", pady=(0, 10))

        # DN (editável, com sugestão automática)
        ttk.Label(frame, text="DN (sugestão automática, editável)", style="Section.TLabel").grid(row=5, column=0, sticky="w", pady=(0, 4))
        proximo_dn = cp.obter_proximo_dn()
        dn_var = tk.StringVar(value=proximo_dn)
        dn_entry = ttk.Entry(frame, textvariable=dn_var, width=40)
        dn_entry.grid(row=6, column=0, sticky="ew", pady=(0, 12))

        buttons = ttk.Frame(frame, style="Card.TFrame")
        buttons.grid(row=7, column=0, sticky="e")

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

        navio_entry.focus_set()
        dialog.wait_window()
        return resultado["valor"]

    def _pedir_dados_periodo(self, programa_cls, titulo):
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

            caminho_navio = str(programa.caminho_navio or "")
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

    def _handle_menu_action(self, item):
        # Casos especiais que precisam coletar dados antes de executar
        if item["label"] == "📁 Criar Pasta":
            self._criar_pasta_ui()  # Chama diretamente (não via _run_action)
            return

        if item["label"] in ("🕒 Fazer Ponto", "↩️ Desfazer Ponto"):
            programa_cls = FazerPonto if item["label"] == "🕒 Fazer Ponto" else ProgramaRemoverPeriodo
            selecao = self._pedir_dados_periodo(programa_cls, item["label"])
            if not selecao:
                self._status.configure(text="Operação cancelada")
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
                import traceback
                self._write_log(f"\n❌ ERRO: {exc}\n", tag="err")
                self._write_log(f"Traceback:\n{traceback.format_exc()}\n", tag="err")
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

            pages = None
            erros = []

            try:
                pages = convert_from_path(str(path))
            except Exception as exc:
                erros.append(str(exc))

            if not pages:
                for poppler_dir in self._poppler_paths_candidatos():
                    try:
                        pages = convert_from_path(str(path), poppler_path=str(poppler_dir))
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
            self._preview_page_index = 0
            self._preview_pdf_path = path
            self._render_preview_page()
            self._notebook.select(1)
        except Exception as exc:
            self._write_log(f"Falha ao renderizar PDF: {exc}\n", tag="warn")

    def _poppler_paths_candidatos(self):
        candidatos = []

        env_poppler = os.environ.get("POPPLER_PATH")
        if env_poppler:
            candidatos.append(Path(env_poppler))

        path_env = os.environ.get("PATH", "")
        for parte in path_env.split(os.pathsep):
            if not parte:
                continue
            if "poppler" in parte.lower():
                candidatos.append(Path(parte))

        candidatos.extend(
            [
                Path(r"C:\poppler-25.12.0\Library\bin"),
                Path(r"C:\poppler\Library\bin"),
                Path(r"C:\Program Files\poppler\Library\bin"),
                Path(r"C:\Program Files (x86)\poppler\Library\bin"),
            ]
        )

        vistos = set()
        validos = []
        for pasta in candidatos:
            chave = str(pasta).lower().strip()
            if not chave or chave in vistos:
                continue
            vistos.add(chave)
            if pasta.exists() and (pasta / "pdfinfo.exe").exists():
                validos.append(pasta)
        return validos

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
