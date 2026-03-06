from yuta_helpers import *
import calendar
from .email_rascunho import criar_rascunho_email_cliente
from .criar_pasta import CriarPasta


class FaturamentoSaoSebastiao:
    """
    ✅ Objetivo (organizado e estável):
    - Selecionar 1 ou MAIS PDFs (Sea Side geralmente vem com 2)
    - Ler TODOS os PDFs selecionados e manter quebras de linha
    - Identificar cliente/porto pela pasta do CLIENTE
    - Se for layout SS (Wilson SS / Sea Side PSS):
        - extrair valores somando entre PDFs (se tiver 2)
        - colar no REPORT VIGIA com o MAPA_FIXO (você já deixou as células)
        - preencher FRONT VIGIA
        - preencher CREDIT NOTE se existir
    - Se for cliente padrão (Aquarius e outros):
        - usar report padrão (datas e horários)
        - (extração financeira pode ser diferente: por enquanto fica como TODO)

    ⚠️ IMPORTANTE:
    - Eu NÃO removo '\n' na normalização, porque sua extração depende de splitlines().
    - A extração do layout SS soma automaticamente tudo que casar (ótimo pra Sea Side com 2 PDFs).
    """

    # ==================================================
    # INIT
    # ==================================================
    def __init__(self, usuario_nome: str | None = None):
        self.caminhos_pdfs: list[Path] = []
        self.paginas_texto: list[dict] = []   # [{pdf, page, texto}]
        self.texto_pdf: str = ""
        self.dados: dict[str, float] = {}
        self.usuario_nome = (usuario_nome or "").strip()

    # ==================================================
    # UTIL: NORMALIZAÇÃO
    # ==================================================
    def _normalizar(self, s: str | None) -> str:
        if not s:
            return ""
        s = unicodedata.normalize("NFKD", str(s))
        s = s.encode("ASCII", "ignore").decode("ASCII")
        s = s.replace("-", " ")  # ajuda "sea-side" -> "sea side"
        return re.sub(r"\s+", " ", s).strip().lower()

    def _br_to_float(self, valor) -> float:
        """Converte '1.721,08' -> 1721.08 ; aceita float/int direto."""
        if valor in (None, "", "NÃO ENCONTRADO"):
            return 0.0
        if isinstance(valor, (int, float)):
            return float(valor)
        return float(str(valor).replace(".", "").replace(",", ".").strip())

    # Alias (você usava _to_float em alguns lugares)
    def _to_float(self, valor) -> float:
        return self._br_to_float(valor)

    # ==================================================
    # UTIL: EXCEL
    # ==================================================
    def _achar_aba(self, wb, nomes_possiveis: list[str]):
        for sheet in wb.sheets:
            nome = sheet.name.strip().lower()
            for n in nomes_possiveis:
                if nome == n.strip().lower():
                    return sheet
        raise RuntimeError(f"Aba não encontrada. Esperado uma de: {nomes_possiveis}")

    # ==================================================
    # IDENTIFICAÇÃO CLIENTE / PORTO
    # ==================================================
    def identificar_cliente_e_porto(self) -> tuple[str, str]:
        """
        Identifica cliente e porto pelo nome da pasta do CLIENTE (pai da pasta do navio).
        """
        if not self.caminhos_pdfs:
            raise RuntimeError("Nenhum PDF carregado para identificar cliente/porto.")

        pasta_navio = self.caminhos_pdfs[0].parent
        pasta_cliente = pasta_navio.parent

        nome_norm = self._normalizar(pasta_cliente.name)

        # WILSON SONS — SÃO SEBASTIÃO
        if "wilson" in nome_norm and "sebastiao" in nome_norm:
            return "WILSON SONS", "SAO SEBASTIAO"

        # SEA SIDE — PSS (mesmo layout de colagem do report)
        if "sea side" in nome_norm and "pss" in nome_norm:
            return "SEA SIDE", "PSS"

        # AQUARIUS — PSS
        if "aquarius" in nome_norm and "pss" in nome_norm:
            return "AQUARIUS", "PSS"

        # PADRÃO
        return pasta_cliente.name.strip().upper(), "PADRAO"

    def _usa_layout_ss(self, cliente: str, porto: str) -> bool:
        return (
            (cliente == "WILSON SONS" and porto == "SAO SEBASTIAO")
            or (cliente == "SEA SIDE" and porto == "PSS")
        )

    # ==================================================
    # PDF: SELEÇÃO E LEITURA (MULTI)
    # ==================================================
    def selecionar_pdfs_ogmo(self):
        root = Tk()
        root.withdraw()
        root.attributes("-topmost", True)

        caminhos = filedialog.askopenfilenames(
            title="Selecione 1 ou MAIS PDFs OGMO (Sea Side pode ter 2)",
            filetypes=[("PDF", "*.pdf")]
        )
        root.destroy()

        if not caminhos:
            raise RuntimeError("Nenhum PDF selecionado")

        self.caminhos_pdfs = [Path(c) for c in caminhos]
        print("📄 PDFs selecionados:")
        for p in self.caminhos_pdfs:
            print(f"   - {p.name}")

    def carregar_pdfs(self):
        self.paginas_texto.clear()

        for caminho in self.caminhos_pdfs:
            with pdfplumber.open(str(caminho)) as pdf:
                for i, page in enumerate(pdf.pages, start=1):
                    txt = page.extract_text() or ""
                    txt = txt.strip()

                    # se veio texto, guarda direto (SEM OCR)
                    if txt:
                        self.paginas_texto.append({"pdf": caminho.name, "page": i, "texto": txt, "src": "TXT"})
                    else:
                        ocr_txt = self._ocr_pagina(caminho, page_num=i)
                        if ocr_txt.strip():
                            self.paginas_texto.append({"pdf": caminho.name, "page": i, "texto": ocr_txt, "src": "OCR"})

        if not self.paginas_texto:
            raise RuntimeError("Nenhuma página com texto (nem pdfplumber nem OCR).")


        self.normalizar_texto_mantendo_linhas()




    def _money_to_float(self, s: str) -> float:
        if s is None:
            return 0.0
        s = str(s).strip()

        # remove espaços (OCR adora meter)
        s = s.replace(" ", "")

        # se tem vírgula e ponto, decide o decimal pelo ÚLTIMO separador
        if "," in s and "." in s:
            if s.rfind(",") > s.rfind("."):
                # 1.234,56  -> decimal = ,
                s = s.replace(".", "").replace(",", ".")
            else:
                # 1,234.56 -> decimal = .
                s = s.replace(",", "")
            return float(s)

        # só vírgula: 1234,56
        if "," in s:
            return float(s.replace(".", "").replace(",", "."))

        # só ponto: 1234.56
        return float(s)

    def _poppler_paths_candidatos(self) -> list[Path]:
        candidatos = []

        env_poppler = os.environ.get("POPPLER_PATH")
        if env_poppler:
            candidatos.append(Path(env_poppler))

        path_env = os.environ.get("PATH", "")
        for parte in path_env.split(os.pathsep):
            if parte and "poppler" in parte.lower():
                candidatos.append(Path(parte))

        if getattr(sys, "frozen", False):
            meipass = Path(getattr(sys, "_MEIPASS", Path(sys.executable).parent))
            exe_dir = Path(sys.executable).resolve().parent
            candidatos.extend(
                [
                    meipass / "poppler" / "Library" / "bin",
                    meipass / "poppler" / "bin",
                    exe_dir / "poppler" / "Library" / "bin",
                    exe_dir / "poppler" / "bin",
                ]
            )
        else:
            raiz_projeto = Path(__file__).resolve().parents[3]
            candidatos.extend(
                [
                    raiz_projeto / "poppler" / "Library" / "bin",
                    raiz_projeto / "poppler" / "bin",
                ]
            )

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

    def _configurar_tesseract(self):
        candidatos = []

        env_tesseract = os.environ.get("TESSERACT_EXE")
        if env_tesseract:
            candidatos.append(Path(env_tesseract))

        if getattr(sys, "frozen", False):
            meipass = Path(getattr(sys, "_MEIPASS", Path(sys.executable).parent))
            exe_dir = Path(sys.executable).resolve().parent
            candidatos.extend(
                [
                    meipass / "tesseract" / "tesseract.exe",
                    exe_dir / "tesseract" / "tesseract.exe",
                ]
            )

        candidatos.extend(
            [
                Path(r"C:\Program Files\Tesseract-OCR\tesseract.exe"),
                Path(r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe"),
            ]
        )

        for exe in candidatos:
            if exe.exists():
                pytesseract.pytesseract.tesseract_cmd = str(exe)
                tessdata_dir = exe.parent / "tessdata"
                if tessdata_dir.exists():
                    os.environ["TESSDATA_PREFIX"] = str(tessdata_dir)
                return



    def _ocr_pagina(self, caminho_pdf: Path, page_num: int, dpi: int = 350, lang: str = "por") -> str:

        self._configurar_tesseract()

        imgs = []
        erros = []

        try:
            imgs = convert_from_path(
                str(caminho_pdf),
                dpi=dpi,
                grayscale=True,
                first_page=page_num,
                last_page=page_num,
            )
        except Exception as exc:
            erros.append(str(exc))

        if not imgs:
            for poppler_dir in self._poppler_paths_candidatos():
                try:
                    imgs = convert_from_path(
                        str(caminho_pdf),
                        dpi=dpi,
                        grayscale=True,
                        poppler_path=str(poppler_dir),
                        first_page=page_num,
                        last_page=page_num,
                    )
                    if imgs:
                        break
                except Exception as exc:
                    erros.append(str(exc))

        if not imgs:
            if erros:
                print(f"⚠️ OCR indisponível (Poppler/Tesseract): {erros[-1]}")
            return ""

        return pytesseract.image_to_string(imgs[0], lang=lang, config="--oem 3 --psm 6")



    def normalizar_texto_mantendo_linhas(self):
        """
        Normaliza espaços mas NÃO remove '\n'.
        Isso mantém sua extração por linha estável.
        """
        blocos = []
        for item in self.paginas_texto:
            texto = item["texto"]
            texto = "\n".join(re.sub(r"[ \t]+", " ", ln).strip() for ln in texto.splitlines())
            item["texto"] = texto
            blocos.append(texto)

        self.texto_pdf = "\n\n".join(blocos)



    # ==================================================
    # PDF ORDER (OGMO 1..N)  -> agora retorna Path (não só nome)
    # ==================================================
    def _ordenar_pdfs_ogmo(self) -> list[Path]:
        """
        Retorna a lista de Paths ordenada pelo número do arquivo:
        FOLHAS OGMO 1.pdf, 2.pdf, 3.pdf ...
        Se não achar número, joga pro final mantendo ordem original.
        """
        def idx(p: Path) -> int:
            nome = p.name
            m = re.search(r"\bOGMO\s*(\d+)\b|\b(\d+)\b", nome, re.IGNORECASE)
            if not m:
                return 10_000
            g = m.group(1) or m.group(2)
            try:
                return int(g)
            except Exception:
                return 10_000

        return sorted(self.caminhos_pdfs, key=idx)


    def _pdfs_ordenados_nomes(self) -> list[str]:
        """Nomes ordenados (string) - útil se você quiser logar."""
        return [p.name for p in self._ordenar_pdfs_ogmo()]


    # ==================================================
    # EXTRAÇÃO - DATA (tolerante a OCR) por PDF (case-insensitive)
    # ==================================================
    def extrair_periodo_por_data(self, pdf_alvo: str | None = None) -> tuple[str, str]:
        if pdf_alvo:
            alvo_norm = pdf_alvo.strip().lower()
            textos = [it["texto"] for it in self.paginas_texto
                    if str(it.get("pdf","")).strip().lower() == alvo_norm]
            texto_busca = "\n".join(textos)
        else:
            texto_busca = self.texto_pdf or ""

        if not texto_busca.strip():
            raise RuntimeError("Período não encontrado no PDF (texto vazio).")

        # tolerância OCR
        rx_per = re.compile(r"per(?:[íi]|l|1|f|0)?odo", re.I)
        rx_ini = re.compile(r"inic(?:ial|iaI|ia1|lal)?", re.I)
        rx_fim = re.compile(r"fina(?:l|I|1)?", re.I)
        rx_data = re.compile(r"\b(\d{1,2}/\d{1,2}/\d{4})\b")

        linhas = texto_busca.splitlines()

        def achar_data(bloco_rx) -> str | None:
            for i, ln in enumerate(linhas):
                ln_norm = ln.replace("\u00ad", "")
                if rx_per.search(ln_norm) and bloco_rx.search(ln_norm):
                    # tenta na mesma linha
                    m = rx_data.search(ln_norm)
                    if m:
                        return m.group(1)
                    # tenta nas próximas 2 linhas (OCR às vezes joga a data abaixo)
                    for j in range(i+1, min(i+3, len(linhas))):
                        m2 = rx_data.search(linhas[j])
                        if m2:
                            return m2.group(1)
            return None

        data_ini = achar_data(rx_ini)
        data_fim = achar_data(rx_fim)

        if not data_ini or not data_fim:
            raise RuntimeError(f"Período (datas) não encontrado. ini={data_ini} fim={data_fim}")

        return data_ini, data_fim


    # ==================================================
    # EXTRAÇÃO - HORÁRIO (tolerante a OCR) por PDF (case-insensitive)
    # ==================================================
    def extrair_periodo_por_horario(self, pdf_alvo: str | None = None) -> tuple[str, str]:
        if pdf_alvo:
            alvo_norm = pdf_alvo.strip().lower()
            textos = [it["texto"] for it in self.paginas_texto
                    if str(it.get("pdf","")).strip().lower() == alvo_norm]
            texto_busca = "\n".join(textos)
        else:
            texto_busca = self.texto_pdf or ""

        if not texto_busca.strip():
            raise RuntimeError("Horários não encontrados (texto vazio).")

        rx_per = re.compile(r"per(?:[íi]|l|1|f|0)?odo", re.I)
        rx_ini = re.compile(r"inic(?:ial|iaI|ia1|lal)?", re.I)
        rx_fim = re.compile(r"fina(?:l|I|1)?", re.I)

        # aceita 07x13, 07×13, 07-13, 07h13
        rx_h = re.compile(r"\b(\d{1,2})\s*[x×h\-]\s*(\d{1,2})\b", re.I)

        linhas = texto_busca.splitlines()

        def achar_horario(bloco_rx) -> str | None:
            for i, ln in enumerate(linhas):
                if rx_per.search(ln) and bloco_rx.search(ln):
                    m = rx_h.search(ln)
                    if m:
                        a, b = int(m.group(1)) % 24, int(m.group(2)) % 24
                        return f"{a:02d}x{b:02d}"
                    for j in range(i+1, min(i+3, len(linhas))):
                        m2 = rx_h.search(linhas[j])
                        if m2:
                            a, b = int(m2.group(1)) % 24, int(m2.group(2)) % 24
                            return f"{a:02d}x{b:02d}"
            return None

        p_ini = achar_horario(rx_ini)
        p_fim = achar_horario(rx_fim)

        if not p_ini or not p_fim:
            raise RuntimeError(f"Período (horários) não encontrado. ini={p_ini} fim={p_fim}")

        p_ini_raw = self._normalizar_horario_texto(p_ini)
        p_fim_raw = self._normalizar_horario_texto(p_fim)

        p_ini_norm = self._bucket_horario_periodo(p_ini_raw)
        p_fim_norm = self._bucket_horario_periodo(p_fim_raw)

        if not p_ini_norm or not p_fim_norm:
            raise RuntimeError(f"Horários inválidos: ini={p_ini} fim={p_fim}")

        if p_ini_raw != p_ini_norm or p_fim_raw != p_fim_norm:
            print(
                f"⚠️ Horário atípico detectado "
                f"({p_ini_raw} -> bucket {p_ini_norm}, {p_fim_raw} -> bucket {p_fim_norm})."
            )

        # Retorna horário REAL do OGMO; o bucket é usado só para cálculo interno
        return p_ini_raw, p_fim_raw

    def _normalizar_horario_texto(self, periodo: str | None) -> str | None:
        if not periodo:
            return None

        s = str(periodo).strip().lower().replace(" ", "")
        s = s.replace("h", "x").replace("×", "x").replace("-", "x").replace(":", "x")

        m = re.match(r"^(\d{1,2})x(\d{1,2})$", s)
        if not m:
            return None

        inicio = int(m.group(1)) % 24
        fim = int(m.group(2)) % 24
        return f"{inicio:02d}x{fim:02d}"

    def _bucket_horario_periodo(self, periodo: str | None) -> str | None:
        """
        Converte horários para bucket padrão do report:
        07x13, 13x19, 19x01, 01x07.
        Ex.: 09x13 -> 07x13.
        """
        normalizado = self._normalizar_horario_texto(periodo)
        if not normalizado:
            return None

        inicio = int(normalizado.split("x", 1)[0])
        candidato = normalizado

        ordem = {"07x13", "13x19", "19x01", "01x07"}
        if candidato in ordem:
            return candidato

        # fallback por faixa da hora inicial
        if 1 <= inicio < 7:
            return "01x07"
        if 7 <= inicio < 13:
            return "07x13"
        if 13 <= inicio < 19:
            return "13x19"
        return "19x01"

    def _duracao_periodo_horas(self, periodo: str | None) -> float | None:
        normalizado = self._normalizar_horario_texto(periodo)
        if not normalizado:
            return None
        inicio, fim = [int(x) for x in normalizado.split("x", 1)]
        if fim >= inicio:
            return float(fim - inicio)
        return float((24 - inicio) + fim)

    # ==================================================
    # PERÍODO MESCLADO N PDFs (primeiro que tem INI, último que tem FIM)
    # ==================================================



    def extrair_datas_mescladas(self) -> tuple[str, str]:
        pdfs = self._ordenar_pdfs_ogmo()
        if not pdfs:
            raise RuntimeError("Nenhum PDF selecionado.")

        # ✅ início = menor OGMO (normalmente 1)
        p_ini = self._achar_pdf_menor_numero() or pdfs[0]

        # ✅ fim = maior OGMO (último: 2, 3, 4...)
        p_fim = self._achar_pdf_maior_numero() or pdfs[-1]

        try:
            di, _ = self.extrair_periodo_por_data(p_ini.name)
        except Exception as e:
            raise RuntimeError(
                f"Não consegui extrair a DATA INICIAL do OGMO {self._numero_ogmo(p_ini.name)} ({p_ini.name}). Erro: {e}"
            ) from e

        try:
            _, df = self.extrair_periodo_por_data(p_fim.name)
        except Exception as e:
            raise RuntimeError(
                f"Não consegui extrair a DATA FINAL do OGMO {self._numero_ogmo(p_fim.name)} ({p_fim.name}). Erro: {e}"
            ) from e

        print(f"✔ Data inicial de: {p_ini.name} -> {di}")
        print(f"✔ Data final de:   {p_fim.name} -> {df}")

        return di, df





    # ==================================================
    # EXTRAÇÃO: LAYOUT SS (WILSON SS / SEA SIDE PSS)
    # ==================================================
    def _somar_valor_item(self, regex_nome: str, paginas_validas: set[int] | None = None, pick: str = "last") -> float:
        total = 0.0

        # ✅ BR ou US "limpo", e evita pegar pedaços quando tem "1.229.35"
        padrao_valor = r"\d{1,3}(?:\.\d{3})*,\d{2}|\b\d+\.\d{2}\b(?!\.)"

        for item in self.paginas_texto:
            if paginas_validas is not None and item.get("page") not in paginas_validas:
                continue

            linhas = item["texto"].splitlines()

            for i, linha in enumerate(linhas):
                if re.search(regex_nome, linha, re.IGNORECASE):
                    vals = re.findall(padrao_valor, linha)
                    if vals:
                        escolhido = vals[0] if pick == "first" else vals[-1]
                        total += self._br_or_us_to_float(escolhido)
                        continue

                    if i + 1 < len(linhas):
                        prox = linhas[i + 1]
                        vals = re.findall(padrao_valor, prox)
                        if vals:
                            escolhido = vals[0] if pick == "first" else vals[-1]
                            total += self._br_or_us_to_float(escolhido)

        return total

    def _debug_match_valores(self, regex_nome: str, paginas_validas=None):
        padrao_valor = r"\d{1,3}(?:\.\d{3})*,\d{2}|\b\d+\.\d{2}\b(?!\.)"
        print(f"\n=== DEBUG MATCHES: {regex_nome} ===")
        for it in self.paginas_texto:
            if paginas_validas is not None and it.get("page") not in paginas_validas:
                continue
            for linha in it["texto"].splitlines():
                if re.search(regex_nome, linha, re.IGNORECASE):
                    vals = re.findall(padrao_valor, linha)
                    print(f"[{it['pdf']} pág {it['page']}] {linha}")
                    print(f"   -> valores: {vals}")
        print("=== FIM DEBUG ===\n")



    def _br_or_us_to_float(self, valor) -> float:
        if valor in (None, "", "NÃO ENCONTRADO"):
            return 0.0
        if isinstance(valor, (int, float)):
            return float(valor)

        s = str(valor).strip()

        # remove espaços dentro do número: "742 266.46" -> "742266.46"
        s = re.sub(r"(?<=\d)\s+(?=\d)", "", s)

        # pt-BR: 1.234,56
        if re.match(r"^\d{1,3}([.\s]\d{3})*,\d{2}$", s):
            s = s.replace(" ", "").replace(".", "").replace(",", ".")
            return float(s)

        # US com milhar: 1,234.56
        if re.match(r"^\d{1,3}([,\s]\d{3})*\.\d{2}$", s):
            s = s.replace(" ", "").replace(",", "")
            return float(s)

        # simples "1234,56"
        if re.match(r"^\d+,\d{2}$", s):
            return float(s.replace(",", "."))

        # simples "1234.56"
        if re.match(r"^\d+\.\d{2}$", s):
            return float(s)

        # fallback: tenta limpar tudo menos dígito , .
        s2 = re.sub(r"[^0-9.,]", "", s)
        if "," in s2 and "." in s2:
            # assume pt-BR (.) milhar e (, ) decimal
            s2 = s2.replace(".", "").replace(",", ".")
        elif "," in s2:
            s2 = s2.replace(",", ".")
        return float(s2)



    def _somar_rat_ajustado(self, paginas_validas: set[int] | None = None, lookahead: int = 6) -> float:
        """
        Pega o VALOR do INSS (RAT Ajustado) ignorando percentual (1,5000%)
        e sem cair no INSS (Terceiros/Previdência).
        Aceita BR (53,24) e US (53.24).
        """
        total = 0.0

        padrao_br = r"\d{1,3}(?:\.\d{3})*,\d{2}(?!\d)"
        padrao_us = r"\d+\.\d{2}(?!\d)"

        for item in self.paginas_texto:
            if paginas_validas is not None and item.get("page") not in paginas_validas:
                continue

            linhas = item["texto"].splitlines()

            for i, linha in enumerate(linhas):
                if re.search(r"INSS\s*\(\s*RAT", linha, re.IGNORECASE):

                    trecho = [linha]
                    for j in range(i + 1, min(len(linhas), i + 1 + lookahead)):
                        ln = linhas[j]

                        # para no próximo INSS que não seja RAT (pra não cair no Terceiros)
                        if re.search(r"INSS\s*\(", ln, re.IGNORECASE) and not re.search(r"INSS\s*\(\s*RAT", ln, re.IGNORECASE):
                            break

                        trecho.append(ln)

                    bloco = " ".join(trecho)

                    # remove percentuais tipo 1,5000%
                    bloco = re.sub(r"\d+(?:[.,]\d+)?\s*%", " ", bloco)

                    # 1) tenta BR
                    vals = re.findall(padrao_br, bloco)
                    if vals:
                        total += self._br_or_us_to_float(vals[-1])
                        continue

                    # 2) tenta US
                    vals = re.findall(padrao_us, bloco)
                    if vals:
                        total += self._br_or_us_to_float(vals[-1])
                        continue

        return total


    def _valor_apos_rs(self, linha: str) -> float | None:
        # pega números logo depois de "R$"
        m = re.search(r"R\$\s*([0-9][0-9\.\,\s]*[0-9][\.,][0-9]{2})", linha, re.IGNORECASE)
        if not m:
            return None
        return self._br_or_us_to_float(m.group(1))


    def _somar_seguranca_trabalhador_avulso(self, paginas_validas: set[int] | None = None) -> float:
        total = 0.0

        # dinheiro BR ou US
        padrao_valor = r"\d{1,3}(?:\.\d{3})*,\d{2}|\b\d+\.\d{2}\b(?!\.)"

        for item in self.paginas_texto:
            if paginas_validas is not None and item.get("page") not in paginas_validas:
                continue

            for linha in item["texto"].splitlines():
                if re.search(r"Seguran[cç]a\s+do\s+Trabalhador\s+Portu[aá]rio\s+Avulso", linha, re.IGNORECASE):
                    # ✅ pega só valores monetários e usa o ÚLTIMO (que é o valor)
                    vals = re.findall(padrao_valor, linha)
                    if vals:
                        total += self._br_or_us_to_float(vals[-1])

        return total



    def _pegar_valor_monetario_da_linha(self, linha: str) -> float | None:
        ln = str(linha)

        # BR: 3.483,17 ou 3483,17 ou 3 483,17
        br = re.findall(r"\d{1,3}(?:[.\s]\d{3})*,\d{2}", ln)

        # US: 1,229.35 ou 1229.35 ou 1 229.35
        us = re.findall(r"\d{1,3}(?:[,\s]\d{3})*\.\d{2}", ln)

        # junta e pega o último valor monetário real da linha
        vals = br + us
        if not vals:
            return None

        return self._money_to_float(vals[-1])



    def _texto_pdf_pagina(self, pdf_nome: str, page_num: int) -> str:
        blocos = [it["texto"] for it in self.paginas_texto
                if it.get("pdf") == pdf_nome and it.get("page") == page_num]
        return "\n".join(blocos)

    def _somar_rotulo_em_pagina(self, pdf_nome: str, page_num: int, rotulo_regex: str) -> float:
        texto = self._texto_pdf_pagina(pdf_nome, page_num)
        if not texto:
            return 0.0

        total = 0.0
        for ln in texto.splitlines():
            if re.search(rotulo_regex, self._normalizar(ln), re.IGNORECASE):
                v = self._valor_apos_rs(ln)  # ✅ sempre após R$
                if v is not None:
                    total += v
        return total



    def extrair_dados_layout_sea_side_wilson(self):
        print("🔍 Extraindo dados – layout SEA SIDE")

        PAG_FIN = {1}
        PAG_HE  = {2}

        self.dados = {
            "Salário Bruto (MMO)": self._somar_valor_item(r"Sal[aá]rio\s+Bruto\s*\(MMO\)", paginas_validas=PAG_FIN, pick="last"),
            "Vale Refeição": self._somar_valor_item(r"Vale\s+Refei", paginas_validas=PAG_FIN, pick="last"),

            # ✅ NOVO
            "Segurança do Trabalhador Portuário Avulso": self._somar_seguranca_trabalhador_avulso(paginas_validas=PAG_FIN),

            "Encargos Administrativos": self._somar_encargos_adm(paginas_validas=PAG_FIN),
            "INSS (RAT Ajustado)": self._somar_rat_ajustado(paginas_validas=PAG_FIN, lookahead=8),
            "Taxas Bancárias": self._somar_valor_item(r"Taxas\s+Banc", paginas_validas=PAG_FIN, pick="last"),
            "Horas Extras": self._somar_valor_item(r"Horas?\s+Extras?", paginas_validas=PAG_HE, pick="last"),
        }





    def _somar_ultimo_valor_por_linha_por_pdf(self, regex_nome: str, paginas_validas: set[int] | None = None) -> dict[str, float]:
        totais = {}
        for item in self.paginas_texto:
            if paginas_validas is not None and item.get("page") not in paginas_validas:
                continue

            pdf = item.get("pdf", "DESCONHECIDO")
            linhas = item["texto"].splitlines()

            for linha in linhas:
                if re.search(regex_nome, linha, re.IGNORECASE):
                    # pega BR e US e também casos com espaço no milhar
                    valores = re.findall(r"\d[\d\.\s]*,\d{2}|\d[\d\.\s]*\.\d{2}", linha)
                    if valores:
                        val = self._br_or_us_to_float(valores[-1].replace(" ", ""))
                        totais[pdf] = totais.get(pdf, 0.0) + val
        return totais


    def _somar_encargos_adm(self, paginas_validas: set[int] | None = None) -> float:
        total = 0.0

        # dinheiro BR ou US
        padrao_valor = r"\d{1,3}(?:\.\d{3})*,\d{2}|\b\d+\.\d{2}\b(?!\.)"

        for item in self.paginas_texto:
            if paginas_validas is not None and item.get("page") not in paginas_validas:
                continue

            for linha in item["texto"].splitlines():
                if re.search(r"Encargos\s+Administrativos?", linha, re.IGNORECASE):
                    # ✅ remove o bloco "TPAS 5,28155" ou "TPAS 5.91828"
                    linha_limpa = re.sub(r"\bTPAS\b\s*\d+(?:[.,]\d+)?", " ", linha, flags=re.IGNORECASE)

                    vals = re.findall(padrao_valor, linha_limpa)
                    if vals:
                        # ✅ aqui queremos o valor final da linha (ex: 68,66 / 23,67)
                        total += self._br_or_us_to_float(vals[-1])

        return total



    def _somar_ultimo_valor_por_linha(self, regex_nome: str, paginas_validas: set[int] | None = None) -> float:
        total = 0.0

        # valor BR ou US, aceitando espaços
        padrao_valor = r"\d{1,3}(?:\.\d{3})*,\d{2}|\b\d+\.\d{2}\b(?!\.)"


        # compila regex uma vez
        rx = re.compile(regex_nome, re.IGNORECASE)

        for item in self.paginas_texto:
            if paginas_validas is not None and item.get("page") not in paginas_validas:
                continue

            for linha in item["texto"].splitlines():
                ln = self._normalizar(linha)  # <<< AQUI É O PULO DO GATO
                if rx.search(ln):
                    vals = re.findall(padrao_valor, linha)  # pega do original pra manter número certo
                    if vals:
                        s = vals[-1].replace(" ", "")
                        total += self._br_or_us_to_float(s)

        return total



    def _somar_valor_apos_rotulo(self, regex_nome: str, paginas_validas: set[int] | None = None, lookahead: int = 12) -> float:
        """
        Acha o rótulo e busca o primeiro valor numérico nas próximas N linhas.
        Resolve:
        - valores em outra linha (Taxas Bancárias)
        - tabelas onde os rótulos vem e os números aparecem abaixo (Horas Extras)
        - número BR e US
        """
        total = 0.0
        padrao_valor = r"\d{1,3}(?:\.\d{3})*,\d{2}|\d+\.\d{2}"

        for item in self.paginas_texto:
            if paginas_validas is not None and item.get("page") not in paginas_validas:
                continue

            linhas = item["texto"].splitlines()

            for i, linha in enumerate(linhas):
                if re.search(regex_nome, linha, re.IGNORECASE):

                    # procura valor na mesma linha + próximas linhas
                    fim = min(len(linhas), i + 1 + lookahead)
                    bloco = " ".join(linhas[i:fim])

                    vals = re.findall(padrao_valor, bloco)
                    if vals:
                        total += self._br_or_us_to_float(vals[0])  # primeiro valor após o rótulo
        return total


    def _numero_ogmo(self, nome: str) -> int | None:
        """
        Extrai o número do OGMO do nome do arquivo.
        Aceita:
        - 'FOLHAS OGMO 1.pdf'
        - 'FOLHAS OGMO (2).pdf'
        - 'OGMO 3.pdf'
        """
        m = re.search(r"\bOGMO\s*\(?\s*(\d+)\s*\)?\b", nome, re.IGNORECASE)
        if m:
            return int(m.group(1))

        # fallback: tenta achar "(n)" no final
        m = re.search(r"\(\s*(\d+)\s*\)", nome)
        if m:
            return int(m.group(1))

        return None



    def _achar_pdf_menor_numero(self) -> Path | None:
        candidatos = []
        for p in self.caminhos_pdfs:
            n = self._numero_ogmo(p.name)
            if n is not None:
                candidatos.append((n, p))
        if not candidatos:
            return None
        return min(candidatos, key=lambda x: x[0])[1]


    def _achar_pdf_maior_numero(self) -> Path | None:
        candidatos = []
        for p in self.caminhos_pdfs:
            n = self._numero_ogmo(p.name)
            if n is not None:
                candidatos.append((n, p))
        if not candidatos:
            return None
        return max(candidatos, key=lambda x: x[0])[1]



    # ==================================================
    # REPORT VIGIA - LAYOUT SS (Wilson SS / Sea Side PSS)
    # ==================================================



    def colar_report_layout_ss(self, wb):
        aba = next(s for s in wb.sheets if s.name.strip().lower() == "report vigia")
        print("📌 Report (layout SS) – colando valores fixos")

        MAPA_FIXO = {
            "Salário Bruto (MMO)": "G22",
            "Vale Refeição": "G25",
            "Segurança do Trabalhador Portuário Avulso": "G26",
            "Encargos Administrativos": "G27",


            "INSS (RAT Ajustado)": "G30",

            "Taxas Bancárias": "G32",
            "Horas Extras": "G35",
        }

        for chave, celula in MAPA_FIXO.items():
            aba.range(celula).value = float(self.dados.get(chave, 0.0) or 0.0)


    def _garantir_linhas_report(self, aba, linha_base: int, total_linhas: int):
        """
        Garante que existam `total_linhas` linhas disponíveis a partir de `linha_base`,
        inserindo linhas abaixo e herdando formatação da linha de cima (sem Copy/PasteSpecial).

        Isso evita:
        - erro PasteSpecial
        - conflito com clipboard
        - bug com células mescladas
        """
        if total_linhas <= 1:
            return

        # Constantes do Excel
        xlShiftDown = -4121
        xlFormatFromLeftOrAbove = 0

        # Precisamos criar (total_linhas - 1) linhas abaixo da base
        qtd_inserir = total_linhas - 1

        # Insere em bloco (mais rápido e mais estável)
        # Ex: base=22, inserir 5 => insere linhas 23..27
        r = aba.api.Rows(linha_base + 1)
        for _ in range(qtd_inserir):
            r.Insert(Shift=xlShiftDown, CopyOrigin=xlFormatFromLeftOrAbove)



    # ==================================================
    # CONFIGURAÇÃO DE MODELO POR CLIENTE
    # ==================================================
    def obter_configuracao_cliente(self, cliente: str, porto: str) -> dict:
        """
        ✅ Aqui fica o coração do “qual modelo usar” e “qual colagem fazer”.
        Você falou:
        - Sea Side tem modelo DIFERENTE de Wilson
        - mas o REPORT (células) é o mesmo modo.
        """
        if self._usa_layout_ss(cliente, porto):
            if cliente == "WILSON SONS":
                modelo = "WILSON SONS - SAO SEBASTIAO.xlsx"
            elif cliente == "SEA SIDE":
                modelo = "SEA SIDE - PSS.xlsx"
            else:
                modelo = f"{cliente} - {porto}.xlsx"

            return {
                "modelo": modelo,
                "colar_report": self.colar_report_layout_ss
            }

        return {
            "modelo": f"{cliente}.xlsx",
            "colar_report": self.colar_report_padrao
        }


    def _escolher_pdf_inicio_fim(self) -> tuple[str, str]:
        """
        Decide qual PDF é o início e qual é o fim.
        - tenta identificar OGMO 1 e OGMO 2 pelo nome
        - fallback: primeiro selecionado = início, último = fim
        Retorna (nome_pdf_inicio, nome_pdf_fim)
        """
        nomes = [p.name for p in self.caminhos_pdfs]

        if len(nomes) == 1:
            return nomes[0], nomes[0]

        # tenta achar "1" e "2" pelo nome do arquivo
        n1 = next((n for n in nomes if re.search(r"(ogmo\s*1|folhas\s*ogmo\s*1|\b1\b)", n, re.I)), None)
        n2 = next((n for n in nomes if re.search(r"(ogmo\s*2|folhas\s*ogmo\s*2|\b2\b)", n, re.I)), None)

        if n1 and n2:
            return n1, n2

        return nomes[0], nomes[-1]



    # ==================================================
    # FRONT VIGIA
    # ==================================================
    def preencher_front_vigia(self, wb):
        try:
            aba = next(s for s in wb.sheets if s.name.strip().lower() == "front vigia")

            pasta = self.caminhos_pdfs[0].parent
            navio = obter_nome_navio(pasta, None)
            nd = obter_dn_da_pasta(pasta)

            # ✅ aqui é o pulo do gato
            if len(self.caminhos_pdfs) >= 2:
                data_ini, data_fim = self.extrair_datas_mescladas()
            else:
                data_ini, data_fim = self.extrair_periodo_por_data()


            def fmt(data_str: str) -> str:
                d = datetime.strptime(data_str, "%d/%m/%Y")
                return f"{calendar.month_name[d.month]} {d.day}, {d.year}"

            aba.range("D15").merge_area.value = navio
            aba.range("D16").merge_area.value = fmt(data_ini)
            aba.range("D17").merge_area.value = fmt(data_fim)

            ano = datetime.now().year % 100
            aba.range("C21").merge_area.value = f"DN {nd}/{ano:02d}"

            hoje = datetime.now()
            meses = ["", "janeiro", "fevereiro", "março", "abril", "maio", "junho",
                    "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"]
            aba.range("C39").merge_area.value = f"  Santos, {hoje.day} de {meses[hoje.month]} de {hoje.year}"

            print("✅ FRONT VIGIA preenchido")

        except StopIteration:
            print("⚠️ Aba FRONT VIGIA não encontrada")

    def atualizar_planilha_controle(self, wb):
        """
        Atualiza a planilha de controle com informações do faturamento VIGIA.
        Preenche colunas B (data), C (serviço), D (ETA), E (ETB), F (cliente), G (navio), J (DN), K (MMO/COSTS).
        """
        try:
            # Obter informações básicas
            pasta = self.caminhos_pdfs[0].parent
            navio = obter_nome_navio(pasta, None)
            nd = obter_dn_da_pasta(pasta)
            cliente_pasta = pasta.parent.name.strip()
            cliente_id, porto_id = self.identificar_cliente_e_porto()
            cliente = self._cliente_coluna_f_controle(cliente_id, porto_id, cliente_pasta)
            
            # Obter data atual
            from datetime import datetime
            data_hoje = datetime.now().strftime("%d/%m/%Y")
            
            # Obter datas do período extraído (já em formato dd/mm/yyyy)
            try:
                if len(self.caminhos_pdfs) >= 2:
                    data_ini, data_fim = self.extrair_datas_mescladas()
                else:
                    data_ini, data_fim = self.extrair_periodo_por_data()
                eta = data_ini
                etb = data_fim
            except Exception:
                eta = ""
                etb = ""
            
            # Buscar valores do REPORT VIGIA para preencher K/L no controle
            usar_mmo = self._cliente_usa_mmo(cliente)
            valores_report = self._buscar_costs_report(wb, return_all=True)
            valor_costs = valores_report.get("COSTS", "")
            valor_mmo = valores_report.get("MMO", "")

            valor_k = valor_costs or valor_mmo
            valor_l = valor_mmo if usar_mmo else None
            
            # ✅ Abrir workbook de controle uma única vez
            criar_pasta = CriarPasta()
            caminho_planilha = criar_pasta._encontrar_planilha()
            from yuta_helpers import openpyxl
            wb_controle = openpyxl.load_workbook(caminho_planilha)
            
            try:
                # Usar CriarPasta para gravar na planilha (reutilizando workbook)
                criar_pasta._gravar_planilha(
                    cliente=cliente,
                    navio=navio,
                    dn=nd,
                    servico="VIGIA",
                    data=data_hoje,
                    eta=eta if eta else "",
                    etb=etb if etb else "",
                    mmo=valor_k,
                    mmo_extra=valor_l,
                    wb_externo=wb_controle
                )
                
                # ✅ Salvar apenas uma vez
                criar_pasta.salvar_planilha_com_retry(wb_controle, caminho_planilha)
                print("✅ Planilha de controle atualizada")
            finally:
                # Fechar workbook
                wb_controle.close()
            
        except Exception as e:
            print(f"⚠️ Erro ao atualizar planilha de controle: {e}")

    def _cliente_coluna_f_controle(self, cliente_id: str, porto_id: str, cliente_padrao: str) -> str:
        """
        Define o nome padronizado para a coluna F da planilha de controle
        nos fluxos de São Sebastião.
        """
        if porto_id in {"PSS", "SAO SEBASTIAO"}:
            mapa_pss = {
                "WILSON SONS": "WILSON (PSS)",
                "AQUARIUS": "AQUARIUS (PSS)",
                "SEA SIDE": "SEA SIDE (PSS)",
            }
            return mapa_pss.get(cliente_id, cliente_padrao)

        return cliente_padrao

    def _normalizar_cliente(self, cliente: str) -> str:
        texto = unicodedata.normalize("NFKD", str(cliente or ""))
        texto = "".join(ch for ch in texto if not unicodedata.combining(ch))
        return texto.upper()

    def _cliente_usa_mmo(self, cliente: str) -> bool:
        cliente_norm = self._normalizar_cliente(cliente).replace(" ", "")
        return "NORTHSTAR" in cliente_norm or "CARGILL" in cliente_norm
    
    def _buscar_costs_report(self, wb, desired_label: str | None = None, return_all: bool = False):
        """
        Busca valores de COSTS/MMO no REPORT VIGIA dinamicamente.
        O valor fica na mesma linha do rótulo, podendo variar entre colunas F/G/H.
        Retorna formato brasileiro sem R$: 16.227,85
        """
        try:
            # Encontra aba REPORT VIGIA
            ws_report = None
            for sh in wb.sheets:
                if sh.name.strip().upper() == "REPORT VIGIA":
                    ws_report = sh
                    break
            
            if not ws_report:
                return ""

            valores_report = {"COSTS": None, "MMO": None}

            def _normalizar_valor(valor_celula):
                if valor_celula in (None, ""):
                    return None
                try:
                    if isinstance(valor_celula, (int, float)):
                        return f"{float(valor_celula):.2f}".replace(".", ",")
                    texto = str(valor_celula).replace("R$", "").replace(" ", "").strip()
                    if not texto:
                        return None
                    if "," in texto:
                        texto = texto.replace(".", "").replace(",", ".")
                    valor_num = float(texto)
                    return f"{valor_num:.2f}".replace(".", ",")
                except Exception:
                    return None

            def _valor_da_linha(linha: int, col_rotulo: str):
                candidatos = ["F", "G", "H"]
                if col_rotulo in ["C", "D", "E", "F", "G"]:
                    prox_col = chr(ord(col_rotulo) + 1)
                    if prox_col not in candidatos:
                        candidatos.insert(0, prox_col)

                for col in candidatos:
                    try:
                        cel = ws_report.range(f"{col}{linha}")
                        valor = _normalizar_valor(cel.value)
                        if valor:
                            return valor
                        try:
                            valor_texto = cel.api.Text
                        except Exception:
                            valor_texto = None
                        valor = _normalizar_valor(valor_texto)
                        if valor:
                            return valor
                    except Exception:
                        continue
                return None

            def _capturar_por_rotulo(rotulo):
                for linha in range(1, 301):
                    for col_letra in ["C", "D", "E", "F", "G", "H"]:
                        try:
                            txt = ws_report.range(f"{col_letra}{linha}").value
                            if not isinstance(txt, str):
                                continue
                            if rotulo not in txt.upper().strip():
                                continue

                            resultado = _valor_da_linha(linha, col_letra)
                            if resultado:
                                print(f"🔎 REPORT {rotulo}: rótulo em {col_letra}{linha}, valor={resultado}")
                                return resultado
                        except Exception:
                            continue
                return None

            valores_report["COSTS"] = _capturar_por_rotulo("COSTS")
            valores_report["MMO"] = _capturar_por_rotulo("MMO")
            
            # Fallback
            for linha in range(1, 301):
                for col_letra in ['C', 'D', 'E', 'F', 'G', 'H']:
                    try:
                        valor_celula = ws_report.range(f"{col_letra}{linha}").value
                        
                        # Verifica se contém COSTS ou MMO
                        if valor_celula and isinstance(valor_celula, str):
                            texto_upper = valor_celula.upper().strip()
                            
                            if "COSTS" in texto_upper or "MMO" in texto_upper:
                                # Busca valor na mesma linha do rótulo
                                try:
                                    resultado = _valor_da_linha(linha, col_letra)
                                    if resultado:
                                        if "MMO" in texto_upper and valores_report["MMO"] is None:
                                            valores_report["MMO"] = resultado
                                        if "COSTS" in texto_upper and valores_report["COSTS"] is None:
                                            valores_report["COSTS"] = resultado
                                except:
                                    pass
                    except:
                        continue

            if return_all:
                return {
                    "COSTS": valores_report["COSTS"] or "",
                    "MMO": valores_report["MMO"] or "",
                }

            if desired_label:
                chave = desired_label.strip().upper()
                if chave in valores_report:
                    return valores_report[chave] or ""
                return ""

            return valores_report["COSTS"] or valores_report["MMO"] or ""
            
            return ""
            
        except Exception as e:
            print(f"❌ Erro ao buscar COSTS: {e}")
            return ""

    # ==================================================
    # CREDIT NOTE
    # ==================================================
    def escrever_cn_credit_note(self, wb, nd: str):
        ws_credit = None
        for sheet in wb.sheets:
            if sheet.name.strip().lower() == "credit note":
                ws_credit = sheet
                break

        if ws_credit is None:
            print("ℹ️ Aba Credit Note não existe — seguindo fluxo.")
            return

        ano = datetime.now().year % 100
        ws_credit.range("C21").merge_area.value = f"CN {nd}/{ano:02d}"
        print("✅ Credit Note preenchida (C21)")

    # ==================================================
    # REPORT VIGIA - PADRÃO (Aquarius e outros)
    # ==================================================

    def _tarifa_por_status(self, ws_report, d: date, periodo: str, status: str) -> float:
        dom_fer = self._is_domingo_ou_feriado(d)
        noite = self._is_noite_por_periodo(periodo)

        # ✅ ATRACADO usa linha 9, FUNDEIO/AO_LARGO usam linha 16
        linha_ref = {
            "ATRACADO": 9,
            "FUNDEIO": 16,
            "AO_LARGO": 16,
        }.get(status)

        if linha_ref is None:
            return 0.0

        # escolhe coluna base
        if not dom_fer and not noite:
            col = "N"
        elif not dom_fer and noite:
            col = "O"
        elif dom_fer and not noite:
            col = "P"
        else:
            col = "Q"

        cell = f"{col}{linha_ref}"
        val = ws_report.range(cell).value
        return float(val or 0.0)



    def preencher_tarifa_por_linha(self, ws_report, linha_base: int, n: int, status: str, coluna_saida: str = "G"):
        """
        Lê data em C{linha} e período em E{linha}.
        Se status == ATRACADO/FUNDEIO/AO_LARGO: escreve tarifa na coluna_saida.
        """
        if status not in ("ATRACADO", "FUNDEIO", "AO_LARGO"):
            return

        for i in range(n):
            linha = linha_base + i
            d = ws_report.range(f"C{linha}").value
            p = ws_report.range(f"E{linha}").value

            if isinstance(d, datetime):
                d = d.date()
            if not isinstance(d, date):
                continue

            tarifa = self._tarifa_por_status(ws_report, d, str(p or ""), status=status)
            ws_report.range(f"{coluna_saida}{linha}").value = tarifa


    def gerar_horarios(self, periodo_inicial: str, periodo_final: str) -> list[str]:
        """
        Gera sequência entre início e fim, respeitando final diferente.
        """
        seq = ["01x07", "07x13", "13x19", "19x01"]
        if periodo_inicial not in seq or periodo_final not in seq:
            # fallback: devolve só inicial se algo vier fora do padrão
            return [periodo_inicial]

        horarios = []
        idx = seq.index(periodo_inicial)

        while True:
            atual = seq[idx]
            horarios.append(atual)
            if atual == periodo_final:
                break
            idx = (idx + 1) % len(seq)

        return horarios

    def preencher_coluna_horarios(self, ws_report, horarios_ogmo: list[str], linha_inicial: int = 22):
        for i, horario in enumerate(horarios_ogmo):
            ws_report.range(f"E{linha_inicial + i}").value = horario

    def _extrair_valores_reais_por_periodo_ogmo(self) -> dict[str, list[float]]:
        """
        Extrai valores monetários diretamente das linhas do OGMO que contenham horário.
        Retorna mapa no formato: {"09x13": [valor1, valor2], "13x19": [valor]}
        """
        mapa: dict[str, list[float]] = {}
        rx_h = re.compile(r"\b(\d{1,2})\s*[x×h\-:]\s*(\d{1,2})\b", re.I)
        rx_v = re.compile(r"\d{1,3}(?:\.\d{3})*,\d{2}|\b\d+\.\d{2}\b(?!\.)")

        for item in self.paginas_texto:
            for linha in item["texto"].splitlines():
                mh = rx_h.search(linha)
                if not mh:
                    continue

                periodo = f"{int(mh.group(1)) % 24:02d}x{int(mh.group(2)) % 24:02d}"
                valores = rx_v.findall(linha)
                if not valores:
                    continue

                try:
                    valor = self._br_or_us_to_float(valores[-1])
                except Exception:
                    continue

                mapa.setdefault(periodo, []).append(float(valor))

        return mapa


    # ==================================================
    # REPORT VIGIA - PADRÃO (Aquarius e outros)
    # ==================================================
    def colar_report_padrao(self, wb):
        aba = self._achar_aba(wb, ["report vigia"])
        print("📌 Report PADRÃO – Outros Clientes")

        if len(self.caminhos_pdfs) >= 2:
            data_ini, data_fim, periodo_inicial, periodo_final = self.extrair_periodo_mesclado_n()
        else:
            data_ini, data_fim = self.extrair_periodo_por_data()
            periodo_inicial, periodo_final = self.extrair_periodo_por_horario()


        print("DEBUG extração:",
                "data_ini=", data_ini,
                "data_fim=", data_fim,
                "p_ini=", periodo_inicial,
                "p_fim=", periodo_final)




        periodos_com_data = self.gerar_periodos_report_padrao_ssz_por_dia(
            data_ini=data_ini,
            data_fim=data_fim,
            periodo_inicial=periodo_inicial,
            periodo_final=periodo_final,
        )

        linha_base = 22
        n = len(periodos_com_data)

        self._garantir_linhas_report(aba, linha_base, n)

        for i, (d, p) in enumerate(periodos_com_data):
            linha = linha_base + i
            celula_data = aba.range(f"C{linha}")
            celula_data.value = self._fmt_data_excel(d)
            celula_data.number_format = "[$-en-US]mmmm d, aaaa"
            aba.range(f"E{linha}").value = p

        # ✅ status pelo nome do navio (o "nome" com (ATRACADO)/(AO LARGO))
        pasta = self.caminhos_pdfs[0].parent
        navio = obter_nome_navio(pasta, None)  # você já tem
        status = self._status_atracacao(navio)

        # ✅ preenche tarifa por linha usando C e E como base
        self.preencher_tarifa_por_linha(aba, linha_base, n, status=status, coluna_saida="G")

        # ✅ sobrescreve com valor real do OGMO quando disponível (horário + valor na mesma linha)
        mapa_valores_ogmo = self._extrair_valores_reais_por_periodo_ogmo()
        if mapa_valores_ogmo:
            sobrescritos = 0
            for i in range(n):
                linha = linha_base + i
                periodo_linha = self._normalizar_horario_texto(aba.range(f"E{linha}").value)
                if not periodo_linha:
                    continue

                lista = mapa_valores_ogmo.get(periodo_linha)
                if not lista:
                    continue

                valor_real = lista.pop(0)
                aba.range(f"G{linha}").value = float(valor_real)
                sobrescritos += 1

            if sobrescritos:
                print(f"✔ {sobrescritos} valor(es) de período sobrescrito(s) com extração direta do OGMO.")

        print(f"✔ Colado {n} períodos + tarifa (status={status}) a partir de C{linha_base}/E{linha_base}")


    def gerar_periodos_report_padrao_ssz_por_dia(self, data_ini, data_fim, periodo_inicial, periodo_final):
        ordem = ["07x13", "13x19", "19x01", "01x07"]

        def norm_periodo(p: str) -> str:
            p = (p or "").strip().lower().replace(" ", "")
            p = p.replace("h", "")
            p = p.replace("-", "x").replace("×", "x")
            p = p.replace(".", "")
            try:
                a, b = p.split("x")
                return f"{int(a):02d}x{int(b):02d}"
            except Exception:
                return (p or "").upper()

        def to_date(d):
            if isinstance(d, datetime):
                return d.date()
            if isinstance(d, date):
                return d
            s = str(d).strip()
            for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d"):
                try:
                    return datetime.strptime(s, fmt).date()
                except Exception:
                    pass
            raise ValueError(f"Data inválida: {d!r}")

        def seq_entre(inicio: str, fim: str) -> list[str]:
            i = ordem.index(inicio)
            out = []
            while True:
                out.append(ordem[i])
                if ordem[i] == fim:
                    break
                i = (i + 1) % 4
                if len(out) > 4:
                    break
            return out

        p_ini_real = norm_periodo(periodo_inicial)
        p_fim_real = norm_periodo(periodo_final)

        p_ini = self._bucket_horario_periodo(p_ini_real)
        p_fim = self._bucket_horario_periodo(p_fim_real)

        if p_ini not in ordem:
            raise ValueError(f"Período inicial inválido: {periodo_inicial!r} -> {p_ini_real!r}")
        if p_fim not in ordem:
            raise ValueError(f"Período final inválido: {periodo_final!r} -> {p_fim_real!r}")

        d_ini = to_date(data_ini)
        d_fim = to_date(data_fim)
        if d_fim < d_ini:
            raise ValueError(f"Data final menor que inicial: {d_ini} > {d_fim}")

        out = []
        dia = d_ini

        while dia <= d_fim:
            # Mantém sua regra: em dias “do meio”, começa sempre em 07x13
            inicio = p_ini if dia == d_ini else "07x13"

            # No último dia, termina no período final; caso contrário, vai até 01x07
            fim = p_fim if dia == d_fim else "01x07"

            for p in seq_entre(inicio, fim):
                out.append((dia, p))  # mantém 01x07 no mesmo dia (como você já faz)

            dia += timedelta(days=1)

            if len(out) > 400:
                raise RuntimeError("Proteção: períodos demais gerados. Verifique datas/períodos extraídos.")

        # preserva horários reais nas bordas quando forem atípicos
        if out:
            if p_ini_real:
                out[0] = (out[0][0], p_ini_real)
            if p_fim_real:
                out[-1] = (out[-1][0], p_fim_real)

        return out




    def _fmt_data_excel(self, d):
        if isinstance(d, datetime):
            return d.date()
        if isinstance(d, date):
            return d
        raise ValueError(f"Data inválida para Excel: {d!r}")


    def extrair_periodo_mesclado_n(self) -> tuple[str, str, str, str]:
        """
        Retorna (data_ini, data_fim, periodo_ini, periodo_fim)
        usando:
        - OGMO menor número = inicio
        - OGMO maior número = fim
        Funciona com 1 ou N PDFs.
        """
        pdfs = self._ordenar_pdfs_ogmo()
        if not pdfs:
            raise RuntimeError("Nenhum PDF selecionado.")

        p_ini = self._achar_pdf_menor_numero() or pdfs[0]
        p_fim = self._achar_pdf_maior_numero() or pdfs[-1]

        try:
            di, _ = self.extrair_periodo_por_data(p_ini.name)
        except Exception as e:
            raise RuntimeError(
                f"Não consegui extrair DATA INICIAL do OGMO {self._numero_ogmo(p_ini.name)} ({p_ini.name}). Erro: {e}"
            ) from e

        try:
            _, df = self.extrair_periodo_por_data(p_fim.name)
        except Exception as e:
            raise RuntimeError(
                f"Não consegui extrair DATA FINAL do OGMO {self._numero_ogmo(p_fim.name)} ({p_fim.name}). Erro: {e}"
            ) from e

        try:
            pi, _ = self.extrair_periodo_por_horario(p_ini.name)
        except Exception as e:
            raise RuntimeError(
                f"Não consegui extrair PERÍODO INICIAL do OGMO {self._numero_ogmo(p_ini.name)} ({p_ini.name}). Erro: {e}"
            ) from e

        try:
            _, pf = self.extrair_periodo_por_horario(p_fim.name)
        except Exception as e:
            raise RuntimeError(
                f"Não consegui extrair PERÍODO FINAL do OGMO {self._numero_ogmo(p_fim.name)} ({p_fim.name}). Erro: {e}"
            ) from e

        print(f"✔ Data inicial de: {p_ini.name} -> {di} ({pi})")
        print(f"✔ Data final de:   {p_fim.name} -> {df} ({pf})")

        return di, df, pi, pf



    # --------------------------------------------------
    # 1) status ATRACADO / AO LARGO pelo nome
    # --------------------------------------------------
    def _status_atracacao(self, nome: str) -> str | None:
        if not nome:
            return None

        s = str(nome).upper()

        # se tiver parênteses, pega dentro; se não, usa tudo
        m = re.search(r"\((.*?)\)", s)
        dentro = m.group(1).strip() if m else s

        dentro = dentro.replace("-", " ").replace("_", " ")
        dentro = re.sub(r"\s+", " ", dentro)

        if "ATRAC" in dentro:
            return "ATRACADO"
        if "FUNDE" in dentro:   # ✅ FUNDEIO
            return "FUNDEIO"
        if "AO LARGO" in dentro or "A LARGO" in dentro or "LARGO" in dentro:
            return "AO_LARGO"

        return None

    # --------------------------------------------------
    # 2) dia/noite pelo período OGMO (coluna E)
    # --------------------------------------------------
    def _is_noite_por_periodo(self, periodo: str) -> bool:
        p = (periodo or "").strip().upper().replace(" ", "")
        # noite: 19x01 e 01x07
        return p in ("19X01", "01X07", "19x01", "01x07")


    # --------------------------------------------------
    # 3) domingo/feriado (mínimo viável)
    #    (se você já tiver função de feriado no projeto, plugue aqui)
    # --------------------------------------------------
    def _is_domingo_ou_feriado(self, d: date) -> bool:
        if isinstance(d, datetime):
            d = d.date()
        # domingo
        if d.weekday() == 6:
            return True

        # ✅ feriados nacionais fixos (mínimo)
        fixos = {
            (1, 1),    # Confraternização Universal
            (4, 21),   # Tiradentes
            (5, 1),    # Dia do Trabalho
            (9, 7),    # Independência
            (10, 12),  # Nossa Sra Aparecida
            (11, 2),   # Finados
            (11, 15),  # Proclamação da República
            (12, 25),  # Natal
        }
        if (d.month, d.day) in fixos:
            return True

        # Se você quiser incluir feriados móveis (Carnaval/Paixão/Corpus Christi),
        # eu adiciono um cálculo de Páscoa e derivados aqui.
        return False


    # --------------------------------------------------
    # 4) pega a tarifa ATRACADO pela regra:
    #    - Seg-Sáb dia:   N9
    #    - Seg-Sáb noite: O9
    #    - Dom/Feriado dia:   P9
    #    - Dom/Feriado noite: Q9
    # --------------------------------------------------
    def _tarifa_atracado(self, ws_report, d: date, periodo: str) -> float:
        dom_fer = self._is_domingo_ou_feriado(d)
        noite = self._is_noite_por_periodo(periodo)

        if not dom_fer and not noite:
            cell = "N9"  # seg-sab dia
        elif not dom_fer and noite:
            cell = "O9"  # seg-sab noite
        elif dom_fer and not noite:
            cell = "P9"  # dom/fer dia
        else:
            cell = "Q9"  # dom/fer noite

        val = ws_report.range(cell).value
        return float(val or 0.0)


    # --------------------------------------------------
    # 5) aplica tarifa linha a linha (baseado em C=data e E=periodo)
    # --------------------------------------------------
    def preencher_tarifa_por_linha(self, ws_report, linha_base: int, n: int, status: str, coluna_saida: str = "G"):
        if status not in ("ATRACADO", "FUNDEIO", "AO_LARGO"):
            return

        # ATRACADO usa linha 9, FUNDEIO/AO_LARGO usam linha 16
        linha_ref = 9 if status == "ATRACADO" else 16

        for i in range(n):
            linha = linha_base + i
            d = ws_report.range(f"C{linha}").value
            p = ws_report.range(f"E{linha}").value

            if isinstance(d, datetime):
                d = d.date()
            if not isinstance(d, date):
                continue

            p_bucket = self._bucket_horario_periodo(str(p or ""))
            if not p_bucket:
                continue

            dom_fer = self._is_domingo_ou_feriado(d)
            noite = self._is_noite_por_periodo(p_bucket)

            if not dom_fer and not noite:
                cell = f"N{linha_ref}"
            elif not dom_fer and noite:
                cell = f"O{linha_ref}"
            elif dom_fer and not noite:
                cell = f"P{linha_ref}"
            else:
                cell = f"Q{linha_ref}"

            val = ws_report.range(cell).value
            tarifa_base = float(val or 0.0)
            ws_report.range(f"{coluna_saida}{linha}").value = tarifa_base


        print("DEBUG status:", status)

    # ==================================================
    # EXECUÇÃO PRINCIPAL
    # ==================================================

    def executar(self, preview=False, selection=None):
        if selection and isinstance(selection, dict) and selection.get("pdfs"):
            self.caminhos_pdfs = [Path(p) for p in selection["pdfs"]]
        else:
            self.selecionar_pdfs_ogmo()
        self.carregar_pdfs()   # já faz pdfplumber e OCR só se precisar
        self.normalizar_texto_mantendo_linhas()




        cliente, porto = self.identificar_cliente_e_porto()
        print(f"\n🚢 FATURAMENTO OGMO – {cliente} / {porto}")

        if cliente == "WILSON SONS":
            self.extrair_dados_layout_sea_side_wilson()
        elif cliente == "SEA SIDE":
            self.extrair_dados_layout_sea_side_wilson()

        else:
            self.dados = {}

        config = self.obter_configuracao_cliente(cliente, porto)

        modelo = obter_pasta_faturamentos() / config["modelo"]
        caminho_local = copiar_para_temp_xlwings(modelo)

        app = xw.App(visible=False, add_book=False)
        wb = app.books.open(str(caminho_local))


        try:
            pasta = self.caminhos_pdfs[0].parent
            navio = obter_nome_navio(pasta, None)
            nd = obter_dn_da_pasta(pasta)

            # FRONT
            self.preencher_front_vigia(wb)
            # ⚠️ NÃO atualiza planilha de controle aqui (será feito após preview)

            # CREDIT NOTE
            self.escrever_cn_credit_note(wb, nd)

            # REPORT
            colar_report = config.get("colar_report")
            if colar_report:
                colar_report(wb)

            # NF
            escrever_nf_faturamento_completo(wb, navio, nd)

            nome_base = f"FATURAMENTO - ND {nd} - MV {navio}"

            if preview:
                preview_pdf = self._export_preview_pdf(wb, nome_base)
                return {
                    "text": "",
                    "preview_pdf": str(preview_pdf) if preview_pdf else None,
                    "selection": {"pdfs": [str(p) for p in self.caminhos_pdfs]},
                }

            # ✅ ATUALIZAR PLANILHA DE CONTROLE (só na execução final)
            self.atualizar_planilha_controle(wb)

            # ✅ SALVAR EXCEL (com wb aberto)
            caminho_excel = salvar_excel_com_nome(wb, pasta, nome_base)
            print(f"💾 Excel salvo em: {caminho_excel}")

            # ✅ Verifica se é cliente WILLIAMS (apenas FRONT VIGIA no PDF)
            apenas_front = "WILLIAMS" in cliente.upper()

            # ✅ GERAR PDF SEM REABRIR O EXCEL (evita erro COM)
            caminho_pdf = gerar_pdf_do_wb_aberto(
                wb, pasta, nome_base, ignorar_abas=("NF",), apenas_front=apenas_front
            )

            nome_cliente = f"{cliente} - {porto}" if porto != "PADRAO" else cliente
            anexos = [caminho_pdf]  # ✅ Removido Excel dos anexos
            anexos.extend(self.caminhos_pdfs)
            try:
                criar_rascunho_email_cliente(
                    nome_cliente,
                    anexos=anexos,
                    dn=str(nd),
                    navio=navio,
                    usuario_nome=self.usuario_nome,
                )
                print("✅ Rascunho do Outlook criado com anexos.")
            except Exception as e:
                print(f"⚠️ Nao foi possivel criar rascunho do Outlook: {e}")

            print("✅ FATURAMENTO FINALIZADO")

        finally:
            wb.close()
            app.quit()

    def _format_brl(self, value):
        if value in (None, ""):
            return "R$ 0,00"
        try:
            num = float(value)
        except Exception:
            return str(value)

        texto = f"{num:,.2f}"
        texto = texto.replace(",", "X").replace(".", ",").replace("X", ".")
        return f"R$ {texto}"

    def _build_preview_text(self, cliente, porto, navio, nd, nome_base):
        linhas = [
            "PRE-VISUALIZACAO",
            "Processo: Faturamento Sao Sebastiao",
            f"Cliente: {cliente}",
            f"Porto: {porto}",
            f"Navio: {navio}",
            f"ND: {nd}",
            f"Nome base: {nome_base}",
        ]

        if self.dados:
            linhas.append("")
            linhas.append("Dados extraidos:")
            chaves = sorted(self.dados.keys())
            limite = min(len(chaves), 15)
            for i in range(limite):
                k = chaves[i]
                linhas.append(f"- {k}: {self._format_brl(self.dados[k])}")
            if len(chaves) > limite:
                linhas.append(f"... {len(chaves) - limite} itens omitidos")

        return "\n".join(linhas)

    def _export_preview_pdf(self, wb, nome_base):
        caminho_pdf = Path(gettempdir()) / f"preview_{nome_base}.pdf"
        if caminho_pdf.exists():
            caminho_pdf.unlink()

        vis_orig = {}
        for sh in wb.sheets:
            vis_orig[sh.name] = sh.api.Visible
            if sh.name.strip().lower() == "nf":
                sh.api.Visible = False

        try:
            ajustar_layout_abas_estrategicas_no_wb(wb)
        except Exception:
            pass

        try:
            wb.api.ExportAsFixedFormat(
                Type=0,
                Filename=str(caminho_pdf),
                Quality=0,
                IncludeDocProperties=True,
                IgnorePrintAreas=False,
                OpenAfterPublish=False,
            )
        finally:
            for sh in wb.sheets:
                if sh.name in vis_orig:
                    sh.api.Visible = vis_orig[sh.name]

        return caminho_pdf
