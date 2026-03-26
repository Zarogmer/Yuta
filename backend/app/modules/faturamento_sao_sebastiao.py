from backend.app.yuta_helpers import *
import calendar
from .email_rascunho import criar_rascunho_email_cliente
from .criar_pasta import CriarPasta
from backend.app.utils.path_utils import configurar_tesseract_runtime, poppler_paths_candidatos


class FaturamentoSaoSebastiao:
    """
    âœ… Objetivo (organizado e estÃ¡vel):
    - Selecionar 1 ou MAIS PDFs (Sea Side geralmente vem com 2)
    - Ler TODOS os PDFs selecionados e manter quebras de linha
    - Identificar cliente/porto pela pasta do CLIENTE
    - Se for layout SS (Wilson SS / Sea Side PSS):
        - extrair valores somando entre PDFs (se tiver 2)
        - colar no REPORT VIGIA com o MAPA_FIXO (vocÃª jÃ¡ deixou as cÃ©lulas)
        - preencher FRONT VIGIA
        - preencher CREDIT NOTE se existir
    - Se for cliente padrÃ£o (Aquarius e outros):
        - usar report padrÃ£o (datas e horÃ¡rios)
        - (extraÃ§Ã£o financeira pode ser diferente: por enquanto fica como TODO)

    âš ï¸ IMPORTANTE:
    - Eu NÃƒO removo '\n' na normalizaÃ§Ã£o, porque sua extraÃ§Ã£o depende de splitlines().
    - A extraÃ§Ã£o do layout SS soma automaticamente tudo que casar (Ã³timo pra Sea Side com 2 PDFs).
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
    # UTIL: NORMALIZAÃ‡ÃƒO
    # ==================================================
    def _normalizar(self, s: str | None) -> str:
        if not s:
            return ""
        s = unicodedata.normalize("NFKD", str(s))
        s = s.encode("ASCII", "ignore").decode("ASCII")
        s = s.replace("-", " ")  # ajuda "sea-side" -> "sea side"
        return re.sub(r"\s+", " ", s).strip().lower()

    def _corrigir_mojibake(self, s: str | None) -> str:
        if not s:
            return ""
        texto = str(s)
        trocas = {
            "NÃƒO": "NAO",
            "NÃ£O": "NAO",
            "NÃ£o": "Nao",
            "marÃ§o": "março",
            "MarÃ§o": "Março",
            "BerÃ§o": "Berco",
        }
        for antigo, novo in trocas.items():
            texto = texto.replace(antigo, novo)
        return texto

    def _br_to_float(self, valor) -> float:
        """Converte '1.721,08' -> 1721.08 ; aceita float/int direto."""
        if valor in (None, "", "NÃƒO ENCONTRADO", "NAO ENCONTRADO", "NÃO ENCONTRADO"):
            return 0.0
        if isinstance(valor, (int, float)):
            return float(valor)
        return float(str(valor).replace(".", "").replace(",", ".").strip())

    # Alias (vocÃª usava _to_float em alguns lugares)
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
        raise RuntimeError(f"Aba nÃ£o encontrada. Esperado uma de: {nomes_possiveis}")

    # ==================================================
    # IDENTIFICAÃ‡ÃƒO CLIENTE / PORTO
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

        # WILSON SONS â€” SÃƒO SEBASTIÃƒO
        if "wilson" in nome_norm and "sebastiao" in nome_norm:
            return "WILSON SONS", "SAO SEBASTIAO"

        # SEA SIDE â€” PSS (mesmo layout de colagem do report)
        if "sea side" in nome_norm and "pss" in nome_norm:
            return "SEA SIDE", "PSS"

        # AQUARIUS â€” PSS
        if "aquarius" in nome_norm and "pss" in nome_norm:
            return "AQUARIUS", "PSS"

        # PADRÃƒO
        return pasta_cliente.name.strip().upper(), "PADRAO"

    def _usa_layout_ss(self, cliente: str, porto: str) -> bool:
        return (
            (cliente == "WILSON SONS" and porto == "SAO SEBASTIAO")
            or (cliente == "SEA SIDE" and porto == "PSS")
        )

    # ==================================================
    # PDF: SELEÃ‡ÃƒO E LEITURA (MULTI)
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

        selecionados = [Path(c) for c in caminhos]

        # Se o usuario selecionar ao menos 1 OGMO, inclui automaticamente
        # os demais PDFs OGMO da mesma pasta.
        extras = []
        try:
            pasta_ref = selecionados[0].parent if selecionados else None
            if pasta_ref and pasta_ref.exists():
                for p in pasta_ref.glob("*.pdf"):
                    nome = self._normalizar(p.name)
                    if "ogmo" in nome:
                        extras.append(p)
        except Exception:
            extras = []

        unicos = {}
        for p in selecionados + extras:
            unicos[str(p.resolve()).lower()] = p

        self.caminhos_pdfs = list(unicos.values())
        print("ðŸ“„ PDFs selecionados:")
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
            raise RuntimeError(
                "Nenhuma pagina com texto (nem pdfplumber nem OCR). "
                "Verifique se o PDF contem texto selecionavel; se for imagem, instale Poppler e Tesseract no cliente."
            )


        self.normalizar_texto_mantendo_linhas()




    def _money_to_float(self, s: str) -> float:
        if s is None:
            return 0.0
        s = str(s).strip()

        # remove espaÃ§os (OCR adora meter)
        s = s.replace(" ", "")

        # se tem vÃ­rgula e ponto, decide o decimal pelo ÃšLTIMO separador
        if "," in s and "." in s:
            if s.rfind(",") > s.rfind("."):
                # 1.234,56  -> decimal = ,
                s = s.replace(".", "").replace(",", ".")
            else:
                # 1,234.56 -> decimal = .
                s = s.replace(",", "")
            return float(s)

        # sÃ³ vÃ­rgula: 1234,56
        if "," in s:
            return float(s.replace(".", "").replace(",", "."))

        # sÃ³ ponto: 1234.56
        return float(s)

    def _poppler_paths_candidatos(self) -> list[Path]:
        return poppler_paths_candidatos()

    def _configurar_tesseract(self):
        configurar_tesseract_runtime()



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
                print(
                    f"OCR indisponivel (Poppler/Tesseract): {erros[-1]} | "
                    "Instale Poppler e Tesseract ou configure POPPLER_PATH/TESSERACT_EXE."
                )
            return ""

        return pytesseract.image_to_string(imgs[0], lang=lang, config="--oem 3 --psm 6")



    def normalizar_texto_mantendo_linhas(self):
        """
        Normaliza espaÃ§os mas NÃƒO remove '\n'.
        Isso mantÃ©m sua extraÃ§Ã£o por linha estÃ¡vel.
        """
        blocos = []
        for item in self.paginas_texto:
            texto = item["texto"]
            texto = "\n".join(re.sub(r"[ \t]+", " ", ln).strip() for ln in texto.splitlines())
            item["texto"] = texto
            blocos.append(texto)

        self.texto_pdf = "\n\n".join(blocos)



    # ==================================================
    # PDF ORDER (OGMO 1..N)  -> agora retorna Path (nÃ£o sÃ³ nome)
    # ==================================================
    def _ordenar_pdfs_ogmo(self) -> list[Path]:
        """
        Retorna a lista de Paths ordenada pelo nÃºmero do arquivo:
        FOLHAS OGMO 1.pdf, 2.pdf, 3.pdf ...
        Se nÃ£o achar nÃºmero, joga pro final mantendo ordem original.
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
        """Nomes ordenados (string) - Ãºtil se vocÃª quiser logar."""
        return [p.name for p in self._ordenar_pdfs_ogmo()]

    def _coletar_periodos_operacionais(self, texto_busca: str) -> list[tuple[str, str]]:
        """
        Fallback para layouts em que o OCR perde "PerÃ­odo Inicial/Final",
        mas mantÃ©m a linha operacional com data e horÃ¡rio.
        """
        if not texto_busca:
            return []

        rx_data = re.compile(r"\b(\d{1,2}/\d{1,2}/\d{4})\b")
        rx_h = re.compile(r"\b(\d{1,2})\s*[xÃ—h\-:]\s*(\d{1,2})\b", re.I)
        encontrados: list[tuple[str, str]] = []

        for linha in texto_busca.splitlines():
            linha_limpa = linha.strip()
            if not linha_limpa:
                continue

            data_m = rx_data.search(linha_limpa)
            hora_m = rx_h.search(linha_limpa)
            if not data_m or not hora_m:
                continue

            linha_norm = self._normalizar(linha_limpa)
            if not any(chave in linha_norm for chave in ("vigil", "opera", "periodo", "simples")):
                continue

            periodo = f"{int(hora_m.group(1)) % 24:02d}x{int(hora_m.group(2)) % 24:02d}"
            encontrados.append((data_m.group(1), periodo))

        return encontrados


    # ==================================================
    # EXTRAÃ‡ÃƒO - DATA (tolerante a OCR) por PDF (case-insensitive)
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
            raise RuntimeError("PerÃ­odo nÃ£o encontrado no PDF (texto vazio).")

        # tolerÃ¢ncia OCR
        rx_per = re.compile(r"per(?:[Ã­i]|l|1|f|0)?odo", re.I)
        rx_ini = re.compile(r"inic(?:ial|iaI|ia1|lal)?", re.I)
        rx_fim = re.compile(r"fina(?:l|I|1)?", re.I)
        rx_data = re.compile(r"\b(\d{1,2}/\d{1,2}/\d{4})\b")

        linhas = texto_busca.splitlines()

        def _data_na_linha_ou_vizinhas(i: int) -> str | None:
            m = rx_data.search(linhas[i])
            if m:
                return m.group(1)

            # OCR/tabela podem quebrar o valor em linhas adjacentes.
            ini = max(0, i - 2)
            fim = min(len(linhas), i + 3)
            for j in range(ini, fim):
                if j == i:
                    continue
                m2 = rx_data.search(linhas[j])
                if m2:
                    return m2.group(1)
            return None

        def achar_data(bloco_rx) -> str | None:
            for i, ln in enumerate(linhas):
                ln_norm = ln.replace("\u00ad", "")
                if rx_per.search(ln_norm) and bloco_rx.search(ln_norm):
                    achada = _data_na_linha_ou_vizinhas(i)
                    if achada:
                        return achada

                # fallback: alguns OGMO trazem apenas "Inicial/Final" sem "Periodo"
                if bloco_rx.search(ln_norm):
                    achada = _data_na_linha_ou_vizinhas(i)
                    if achada:
                        return achada
            return None

        data_ini = achar_data(rx_ini)
        data_fim = achar_data(rx_fim)

        if not data_ini or not data_fim:
            periodos_operacionais = self._coletar_periodos_operacionais(texto_busca)
            if periodos_operacionais:
                if not data_ini:
                    data_ini = periodos_operacionais[0][0]
                if not data_fim:
                    data_fim = periodos_operacionais[-1][0]

        if not data_ini or not data_fim:
            raise RuntimeError(f"PerÃ­odo (datas) nÃ£o encontrado. ini={data_ini} fim={data_fim}")

        return data_ini, data_fim


    # ==================================================
    # EXTRAÃ‡ÃƒO - HORÃRIO (tolerante a OCR) por PDF (case-insensitive)
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
            raise RuntimeError("HorÃ¡rios nÃ£o encontrados (texto vazio).")

        rx_per = re.compile(r"per(?:[Ã­i]|l|1|f|0)?odo", re.I)
        rx_ini = re.compile(r"inic(?:ial|iaI|ia1|lal)?", re.I)
        rx_fim = re.compile(r"fina(?:l|I|1)?", re.I)

        # aceita 07x13, 07Ã—13, 07-13, 07h13
        rx_h = re.compile(r"\b(\d{1,2})\s*[xÃ—h\-]\s*(\d{1,2})\b", re.I)

        linhas = texto_busca.splitlines()

        def _horario_na_linha_ou_vizinhas(i: int) -> str | None:
            m = rx_h.search(linhas[i])
            if m:
                a, b = int(m.group(1)) % 24, int(m.group(2)) % 24
                return f"{a:02d}x{b:02d}"

            ini = max(0, i - 2)
            fim = min(len(linhas), i + 3)
            for j in range(ini, fim):
                if j == i:
                    continue
                m2 = rx_h.search(linhas[j])
                if m2:
                    a, b = int(m2.group(1)) % 24, int(m2.group(2)) % 24
                    return f"{a:02d}x{b:02d}"
            return None

        def achar_horario(bloco_rx) -> str | None:
            for i, ln in enumerate(linhas):
                ln_norm = ln.replace("\u00ad", "")
                if rx_per.search(ln_norm) and bloco_rx.search(ln_norm):
                    achado = _horario_na_linha_ou_vizinhas(i)
                    if achado:
                        return achado

                # fallback: alguns OGMO trazem so "Inicial/Final" na linha da tabela.
                if bloco_rx.search(ln_norm):
                    achado = _horario_na_linha_ou_vizinhas(i)
                    if achado:
                        return achado
            return None

        def achar_horario_global(regex: re.Pattern[str]) -> str | None:
            m = regex.search(texto_busca)
            if not m:
                return None
            a, b = int(m.group(1)) % 24, int(m.group(2)) % 24
            return f"{a:02d}x{b:02d}"

        p_ini = achar_horario(rx_ini)
        p_fim = achar_horario(rx_fim)

        if not p_ini:
            p_ini = achar_horario_global(
                re.compile(r"(?:periodo\s*)?inic(?:ial|iaI|ia1|lal)?[^\d\n]{0,25}(\d{1,2})\s*[xÃ—h\-:]\s*(\d{1,2})", re.I)
            )
        if not p_fim:
            p_fim = achar_horario_global(
                re.compile(r"(?:periodo\s*)?fina(?:l|I|1)?[^\d\n]{0,25}(\d{1,2})\s*[xÃ—h\-:]\s*(\d{1,2})", re.I)
            )

        if not p_ini or not p_fim:
            periodos_operacionais = self._coletar_periodos_operacionais(texto_busca)
            if periodos_operacionais:
                if not p_ini:
                    p_ini = periodos_operacionais[0][1]
                if not p_fim:
                    p_fim = periodos_operacionais[-1][1]

        if not p_ini or not p_fim:
            raise RuntimeError(f"PerÃ­odo (horÃ¡rios) nÃ£o encontrado. ini={p_ini} fim={p_fim}")

        p_ini_raw = self._normalizar_horario_texto(p_ini)
        p_fim_raw = self._normalizar_horario_texto(p_fim)

        p_ini_norm = self._bucket_horario_periodo(p_ini_raw)
        p_fim_norm = self._bucket_horario_periodo(p_fim_raw)

        if not p_ini_norm or not p_fim_norm:
            raise RuntimeError(f"HorÃ¡rios invÃ¡lidos: ini={p_ini} fim={p_fim}")

        if p_ini_raw != p_ini_norm or p_fim_raw != p_fim_norm:
            print(
                f"âš ï¸ HorÃ¡rio atÃ­pico detectado "
                f"({p_ini_raw} -> bucket {p_ini_norm}, {p_fim_raw} -> bucket {p_fim_norm})."
            )

        # Retorna horÃ¡rio REAL do OGMO; o bucket Ã© usado sÃ³ para cÃ¡lculo interno
        return p_ini_raw, p_fim_raw

    def _normalizar_horario_texto(self, periodo: str | None) -> str | None:
        if not periodo:
            return None

        s = str(periodo).strip().lower().replace(" ", "")
        s = s.replace("h", "x").replace("Ã—", "x").replace("-", "x").replace(":", "x")

        m = re.match(r"^(\d{1,2})x(\d{1,2})$", s)
        if not m:
            return None

        inicio = int(m.group(1)) % 24
        fim = int(m.group(2)) % 24
        return f"{inicio:02d}x{fim:02d}"

    def _bucket_horario_periodo(self, periodo: str | None) -> str | None:
        """
        Converte horÃ¡rios para bucket padrÃ£o do report:
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

    def _atrac_fund_do_nome(self, nome: str | None) -> str | None:
        """
        Extrai marcador ATRAC/FUND do nome do navio/pasta.
        Ex.: "FEDERAL DART (ATRACADO)" -> "ATRACADO", "(...FUND)" -> "FUNDEIO".
        """
        if not nome:
            return None

        s = self._normalizar(nome)
        if "fund" in s or "ao largo" in s:
            return "FUNDEIO"
        if "atrac" in s:
            return "ATRACADO"
        return None

    # ==================================================
    # PERÃODO MESCLADO N PDFs (primeiro que tem INI, Ãºltimo que tem FIM)
    # ==================================================



    def extrair_datas_mescladas(self) -> tuple[str, str]:
        pdfs = self._ordenar_pdfs_ogmo()
        if not pdfs:
            raise RuntimeError("Nenhum PDF selecionado.")

        # âœ… inÃ­cio = menor OGMO (normalmente 1)
        p_ini = self._achar_pdf_menor_numero() or pdfs[0]

        # âœ… fim = maior OGMO (Ãºltimo: 2, 3, 4...)
        p_fim = self._achar_pdf_maior_numero() or pdfs[-1]

        try:
            di, _ = self.extrair_periodo_por_data(p_ini.name)
        except Exception as e:
            raise RuntimeError(
                f"NÃ£o consegui extrair a DATA INICIAL do OGMO {self._numero_ogmo(p_ini.name)} ({p_ini.name}). Erro: {e}"
            ) from e

        try:
            _, df = self.extrair_periodo_por_data(p_fim.name)
        except Exception as e:
            raise RuntimeError(
                f"NÃ£o consegui extrair a DATA FINAL do OGMO {self._numero_ogmo(p_fim.name)} ({p_fim.name}). Erro: {e}"
            ) from e

        print(f"âœ” Data inicial de: {p_ini.name} -> {di}")
        print(f"âœ” Data final de:   {p_fim.name} -> {df}")

        return di, df





    # ==================================================
    # EXTRAÃ‡ÃƒO: LAYOUT SS (WILSON SS / SEA SIDE PSS)
    # ==================================================
    def _somar_valor_item(self, regex_nome: str, paginas_validas: set[int] | None = None, pick: str = "last") -> float:
        total = 0.0

        # âœ… BR ou US "limpo", e evita pegar pedaÃ§os quando tem "1.229.35"
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
                    print(f"[{it['pdf']} pÃ¡g {it['page']}] {linha}")
                    print(f"   -> valores: {vals}")
        print("=== FIM DEBUG ===\n")



    def _br_or_us_to_float(self, valor) -> float:
        if valor in (None, "", "NÃƒO ENCONTRADO", "NAO ENCONTRADO", "NÃO ENCONTRADO"):
            return 0.0
        if isinstance(valor, (int, float)):
            return float(valor)

        s = str(valor).strip()

        # remove espaÃ§os dentro do nÃºmero: "742 266.46" -> "742266.46"
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

        # fallback: tenta limpar tudo menos dÃ­gito , .
        s2 = re.sub(r"[^0-9.,]", "", s)
        if "," in s2 and "." in s2:
            # assume pt-BR (.) milhar e (, ) decimal
            s2 = s2.replace(".", "").replace(",", ".")
        elif "," in s2:
            s2 = s2.replace(",", ".")
        return float(s2)

    def extrair_berco_ou_warehouse(self) -> str | None:
        """
        Extrai BERCO/WAREHOUSE do texto dos PDFs OGMO ja carregados.
        """
        rotulos = [
            re.compile(r"\bber[cç]o\b\s*:?\s*(.*)", re.I),
            re.compile(r"\bwarehouse\b\s*:?\s*(.*)", re.I),
            re.compile(r"\bberthed\b\s*:?\s*(.*)", re.I),
        ]

        linhas = []
        for item in self.paginas_texto:
            texto = self._corrigir_mojibake(item.get("texto", ""))
            linhas.extend(texto.splitlines())

        for idx, linha in enumerate(linhas):
            linha_limpa = self._corrigir_mojibake(linha).strip()
            if not linha_limpa:
                continue

            for rx in rotulos:
                m = rx.search(linha_limpa)
                if not m:
                    continue

                valor = (m.group(1) or "").strip(" :-")
                if not valor:
                    for j in range(idx + 1, min(idx + 4, len(linhas))):
                        prox = self._corrigir_mojibake(linhas[j]).strip()
                        if not prox:
                            continue
                        if re.search(r"\b(ber[cç]o|warehouse|berthed|sailed|periodo|per[ií]odo)\b", prox, re.I):
                            continue
                        valor = prox
                        break

                if valor:
                    valor_norm = self._corrigir_mojibake(valor).strip()
                    return valor_norm.upper()

        return None



    def _somar_rat_ajustado(self, paginas_validas: set[int] | None = None, lookahead: int = 6) -> float:
        """
        Pega o VALOR do INSS (RAT Ajustado) ignorando percentual (1,5000%)
        e sem cair no INSS (Terceiros/PrevidÃªncia).
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
                    linha_sem_percentual = re.sub(r"\d+(?:[.,]\d+)?\s*%", " ", linha)
                    vals_mesma_linha = re.findall(padrao_br, linha_sem_percentual)
                    if vals_mesma_linha:
                        total += self._br_or_us_to_float(vals_mesma_linha[-1])
                        continue

                    trecho = [linha]
                    for j in range(i + 1, min(len(linhas), i + 1 + lookahead)):
                        ln = linhas[j]

                        # para no prÃ³ximo INSS que nÃ£o seja RAT (pra nÃ£o cair no Terceiros)
                        if (
                            (re.search(r"[TI]?NSS\s*\(", ln, re.IGNORECASE) and not re.search(r"INSS\s*\(\s*RAT", ln, re.IGNORECASE))
                            or re.search(r"Subtotal|Taxas\s+Ban|Total\s+a\s+ser\s+recolhido|Observa", ln, re.IGNORECASE)
                        ):
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

    def _somar_taxas_bancarias(self, paginas_validas: set[int] | None = None) -> float:
        total = 0.0
        padrao_valor = r"\d{1,3}(?:\.\d{3})*,\d{2}|\b\d+\.\d{2}\b(?!\.)"

        for item in self.paginas_texto:
            if paginas_validas is not None and item.get("page") not in paginas_validas:
                continue

            linhas = item["texto"].splitlines()
            for i, linha in enumerate(linhas):
                linha_norm = self._normalizar(linha)
                if "taxas" not in linha_norm or "ban" not in linha_norm:
                    continue

                vals = re.findall(padrao_valor, linha)
                if vals:
                    total += self._br_or_us_to_float(vals[-1])
                    continue

                if i + 1 < len(linhas):
                    prox = linhas[i + 1]
                    vals = re.findall(padrao_valor, prox)
                    if vals:
                        total += self._br_or_us_to_float(vals[0])

        return total


    def _valor_apos_rs(self, linha: str) -> float | None:
        # pega nÃºmeros logo depois de "R$"
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
                # tolera acento quebrado (mojibake) e OCR variando letras.
                if re.search(r"Segur.{0,6}\s+do\s+Trabalhador\s+Portu.{0,6}\s+Avulso", linha, re.IGNORECASE):
                    # âœ… pega sÃ³ valores monetÃ¡rios e usa o ÃšLTIMO (que Ã© o valor)
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

        # junta e pega o Ãºltimo valor monetÃ¡rio real da linha
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
                v = self._valor_apos_rs(ln)  # âœ… sempre apÃ³s R$
                if v is not None:
                    total += v
        return total

    def _detectar_paginas_layout_ss(self) -> tuple[set[int], set[int]]:
        """
        Detecta as paginas financeiras (operacionais) e administrativas por conteudo,
        sem depender de numero fixo de pagina.
        """
        pag_fin: set[int] = set()
        pag_adm: set[int] = set()

        for item in self.paginas_texto:
            p = int(item.get("page") or 0)
            texto_n = self._normalizar(item.get("texto", ""))
            if not p or not texto_n:
                continue

            if ("resumo do custeio" in texto_n and "valores operacionais" in texto_n) or "tpas" in texto_n:
                pag_fin.add(p)

            if ("resumo do custeio" in texto_n and "valores administrativos" in texto_n) or "ogmo" in texto_n and "administrativos" in texto_n:
                pag_adm.add(p)

        # fallback conservador para manter compatibilidade.
        if not pag_fin:
            pag_fin = {1}
        if not pag_adm:
            pag_adm = {2}

        return pag_fin, pag_adm



    def extrair_dados_layout_sea_side_wilson(self):
        print("ðŸ” Extraindo dados â€“ layout SEA SIDE")

        PAG_FIN, PAG_HE = self._detectar_paginas_layout_ss()
        print(f"ðŸ”Ž PAGINAS detectadas layout SS: operacionais={sorted(PAG_FIN)} administrativas={sorted(PAG_HE)}")

        self.dados = {
            "SalÃ¡rio Bruto (MMO)": self._somar_valor_item(r"Sal.{0,6}rio\s+Bruto\s*\(MM[O0]\)", paginas_validas=PAG_FIN, pick="last"),
            "Vale RefeiÃ§Ã£o": self._somar_valor_item(r"Vale\s+Refei", paginas_validas=PAG_FIN, pick="last"),

            # âœ… NOVO
            "SeguranÃ§a do Trabalhador PortuÃ¡rio Avulso": self._somar_seguranca_trabalhador_avulso(paginas_validas=PAG_FIN),

            "Encargos Administrativos": self._somar_encargos_adm(paginas_validas=PAG_FIN),
            "INSS (RAT Ajustado)": self._somar_rat_ajustado(paginas_validas=PAG_FIN, lookahead=8),
            "Taxas BancÃ¡rias": self._somar_taxas_bancarias(paginas_validas=PAG_FIN),
            "Horas Extras": self._somar_valor_item(r"Horas?\s+Extras?", paginas_validas=PAG_HE, pick="last"),
        }

        for k, v in self.dados.items():
            print(f"ðŸ”Ž EXTRAIDO {k}: {self._format_brl(v)}")





    def _somar_ultimo_valor_por_linha_por_pdf(self, regex_nome: str, paginas_validas: set[int] | None = None) -> dict[str, float]:
        totais = {}
        for item in self.paginas_texto:
            if paginas_validas is not None and item.get("page") not in paginas_validas:
                continue

            pdf = item.get("pdf", "DESCONHECIDO")
            linhas = item["texto"].splitlines()

            for linha in linhas:
                if re.search(regex_nome, linha, re.IGNORECASE):
                    # pega BR e US e tambÃ©m casos com espaÃ§o no milhar
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
                    # âœ… remove o bloco "TPAS 5,28155" ou "TPAS 5.91828"
                    linha_limpa = re.sub(r"\bTPAS\b\s*\d+(?:[.,]\d+)?", " ", linha, flags=re.IGNORECASE)

                    vals = re.findall(padrao_valor, linha_limpa)
                    if vals:
                        # âœ… aqui queremos o valor final da linha (ex: 68,66 / 23,67)
                        total += self._br_or_us_to_float(vals[-1])

        return total



    def _somar_ultimo_valor_por_linha(self, regex_nome: str, paginas_validas: set[int] | None = None) -> float:
        total = 0.0

        # valor BR ou US, aceitando espaÃ§os
        padrao_valor = r"\d{1,3}(?:\.\d{3})*,\d{2}|\b\d+\.\d{2}\b(?!\.)"


        # compila regex uma vez
        rx = re.compile(regex_nome, re.IGNORECASE)

        for item in self.paginas_texto:
            if paginas_validas is not None and item.get("page") not in paginas_validas:
                continue

            for linha in item["texto"].splitlines():
                ln = self._normalizar(linha)  # <<< AQUI Ã‰ O PULO DO GATO
                if rx.search(ln):
                    vals = re.findall(padrao_valor, linha)  # pega do original pra manter nÃºmero certo
                    if vals:
                        s = vals[-1].replace(" ", "")
                        total += self._br_or_us_to_float(s)

        return total



    def _somar_valor_apos_rotulo(self, regex_nome: str, paginas_validas: set[int] | None = None, lookahead: int = 12) -> float:
        """
        Acha o rÃ³tulo e busca o primeiro valor numÃ©rico nas prÃ³ximas N linhas.
        Resolve:
        - valores em outra linha (Taxas BancÃ¡rias)
        - tabelas onde os rÃ³tulos vem e os nÃºmeros aparecem abaixo (Horas Extras)
        - nÃºmero BR e US
        """
        total = 0.0
        padrao_valor = r"\d{1,3}(?:\.\d{3})*,\d{2}|\d+\.\d{2}"

        for item in self.paginas_texto:
            if paginas_validas is not None and item.get("page") not in paginas_validas:
                continue

            linhas = item["texto"].splitlines()

            for i, linha in enumerate(linhas):
                if re.search(regex_nome, linha, re.IGNORECASE):

                    # procura valor na mesma linha + prÃ³ximas linhas
                    fim = min(len(linhas), i + 1 + lookahead)
                    bloco = " ".join(linhas[i:fim])

                    vals = re.findall(padrao_valor, bloco)
                    if vals:
                        total += self._br_or_us_to_float(vals[0])  # primeiro valor apÃ³s o rÃ³tulo
        return total


    def _numero_ogmo(self, nome: str) -> int | None:
        """
        Extrai o nÃºmero do OGMO do nome do arquivo.
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
        print("ðŸ“Œ Report (layout SS) â€“ colando valores fixos")

        MAPA_FIXO = {
            "SalÃ¡rio Bruto (MMO)": "G22",
            "Vale RefeiÃ§Ã£o": "G25",
            "SeguranÃ§a do Trabalhador PortuÃ¡rio Avulso": "G26",
            "Encargos Administrativos": "G27",


            "INSS (RAT Ajustado)": "G30",

            "Taxas BancÃ¡rias": "G32",
            "Horas Extras": "G35",
        }

        for chave, celula in MAPA_FIXO.items():
            valor = float(self.dados.get(chave, 0.0) or 0.0)
            aba.range(celula).value = valor
            print(f"ðŸ”Ž REPORT {celula} <= {chave}: {self._format_brl(valor)}")


    def _garantir_linhas_report(self, aba, linha_base: int, total_linhas: int):
        """
        Garante que existam `total_linhas` linhas disponÃ­veis a partir de `linha_base`,
        inserindo linhas abaixo e herdando formataÃ§Ã£o da linha de cima (sem Copy/PasteSpecial).

        Isso evita:
        - erro PasteSpecial
        - conflito com clipboard
        - bug com cÃ©lulas mescladas
        """
        if total_linhas <= 1:
            return

        # Constantes do Excel
        xlShiftDown = -4121
        xlFormatFromLeftOrAbove = 0

        # Precisamos criar (total_linhas - 1) linhas abaixo da base
        qtd_inserir = total_linhas - 1

        # Insere em bloco (mais rÃ¡pido e mais estÃ¡vel)
        # Ex: base=22, inserir 5 => insere linhas 23..27
        r = aba.api.Rows(linha_base + 1)
        for _ in range(qtd_inserir):
            r.Insert(Shift=xlShiftDown, CopyOrigin=xlFormatFromLeftOrAbove)



    # ==================================================
    # CONFIGURAÃ‡ÃƒO DE MODELO POR CLIENTE
    # ==================================================
    def obter_configuracao_cliente(self, cliente: str, porto: str) -> dict:
        """
        âœ… Aqui fica o coraÃ§Ã£o do â€œqual modelo usarâ€ e â€œqual colagem fazerâ€.
        VocÃª falou:
        - Sea Side tem modelo DIFERENTE de Wilson
        - mas o REPORT (cÃ©lulas) Ã© o mesmo modo.
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
        Decide qual PDF Ã© o inÃ­cio e qual Ã© o fim.
        - tenta identificar OGMO 1 e OGMO 2 pelo nome
        - fallback: primeiro selecionado = inÃ­cio, Ãºltimo = fim
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
            cliente, porto = self.identificar_cliente_e_porto()

            # âœ… aqui Ã© o pulo do gato
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

            # Para Sea Side/Wilson SS, manter o valor padrao de WAREHOUSE do modelo.
            if not self._usa_layout_ss(cliente, porto):
                if cliente == "AQUARIUS" and porto == "PSS":
                    marcador = self._atrac_fund_do_nome(navio)
                    if marcador:
                        aba.range("D18").merge_area.value = marcador
                        print(f"ℹ️ FRONT VIGIA D18 preenchido por nome do navio (AQUARIUS): {marcador}")
                    else:
                        berco = self.extrair_berco_ou_warehouse()
                        aba.range("D18").merge_area.value = berco if berco else "NAO ENCONTRADO"
                else:
                    berco = self.extrair_berco_ou_warehouse()
                    aba.range("D18").merge_area.value = berco if berco else "NAO ENCONTRADO"
            else:
                print("ℹ️ FRONT VIGIA D18 mantido como padrao do modelo para layout SS.")

            ano = datetime.now().year % 100
            aba.range("C21").merge_area.value = f"DN {nd}/{ano:02d}"

            hoje = datetime.now()
            meses = ["", "janeiro", "fevereiro", "março", "abril", "maio", "junho",
                "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"]
            aba.range("C39").merge_area.value = f"  Santos, {hoje.day} de {meses[hoje.month]} de {hoje.year}"

            cliente_front = str(aba.range("C9").value or "").upper()
            if self.usuario_nome and "NORTH STAR" not in cliente_front:
                aba.range("C42").merge_area.value = f"  {self.usuario_nome}"

            print("âœ… FRONT VIGIA preenchido")

        except StopIteration:
            print("âš ï¸ Aba FRONT VIGIA nÃ£o encontrada")

    def atualizar_planilha_controle(self, wb):
        """
        Atualiza a planilha de controle com informaÃ§Ãµes do faturamento VIGIA.
        Preenche colunas B (data), C (serviÃ§o), D (ETA), E (ETB), F (cliente), G (navio), J (DN), K (MMO/COSTS).
        """
        try:
            # Obter informaÃ§Ãµes bÃ¡sicas
            pasta = self.caminhos_pdfs[0].parent
            navio = obter_nome_navio(pasta, None)
            nd = obter_dn_da_pasta(pasta)
            cliente_pasta = pasta.parent.name.strip()
            cliente_id, porto_id = self.identificar_cliente_e_porto()
            cliente = self._cliente_coluna_f_controle(cliente_id, porto_id, cliente_pasta)
            
            # Obter data atual
            from datetime import datetime
            data_hoje = datetime.now().strftime("%d/%m/%Y")
            
            # Obter datas do perÃ­odo extraÃ­do (jÃ¡ em formato dd/mm/yyyy)
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
            
            # âœ… Abrir workbook de controle uma Ãºnica vez
            criar_pasta = CriarPasta()
            caminho_planilha = criar_pasta._encontrar_planilha()
            from backend.app.yuta_helpers import openpyxl
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
                
                # âœ… Salvar apenas uma vez
                criar_pasta.salvar_planilha_com_retry(wb_controle, caminho_planilha)
                print("âœ… Planilha de controle atualizada")
            finally:
                # Fechar workbook
                wb_controle.close()
            
        except Exception as e:
            print(f"âš ï¸ Erro ao atualizar planilha de controle: {e}")

    def _cliente_coluna_f_controle(self, cliente_id: str, porto_id: str, cliente_padrao: str) -> str:
        """
        Define o nome padronizado para a coluna F da planilha de controle
        nos fluxos de SÃ£o SebastiÃ£o.
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
        O valor fica na mesma linha do rÃ³tulo, podendo variar entre colunas F/G/H.
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
                                print(f"ðŸ”Ž REPORT {rotulo}: rÃ³tulo em {col_letra}{linha}, valor={resultado}")
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
                        
                        # Verifica se contÃ©m COSTS ou MMO
                        if valor_celula and isinstance(valor_celula, str):
                            texto_upper = valor_celula.upper().strip()
                            
                            if "COSTS" in texto_upper or "MMO" in texto_upper:
                                # Busca valor na mesma linha do rÃ³tulo
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
            print(f"âŒ Erro ao buscar COSTS: {e}")
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
            print("â„¹ï¸ Aba Credit Note nÃ£o existe â€” seguindo fluxo.")
            return

        ano = datetime.now().year % 100
        ws_credit.range("C21").merge_area.value = f"CN {nd}/{ano:02d}"
        print("âœ… Credit Note preenchida (C21)")

    # ==================================================
    # REPORT VIGIA - PADRÃƒO (Aquarius e outros)
    # ==================================================

    def _tarifa_por_status(self, ws_report, d: date, periodo: str, status: str) -> float:
        dom_fer = self._is_domingo_ou_feriado(d)
        noite = self._is_noite_por_periodo(periodo)

        # âœ… ATRACADO usa linha 9, FUNDEIO/AO_LARGO usam linha 16
        linha_ref = {
            "ATRACADO": 9,
            "FUNDEIO": 16,
            "AO_LARGO": 16,
        }.get(status)

        if linha_ref is None:
            return 0.0

        col = self._coluna_tarifa_por_regra(dom_fer, noite)

        cell = f"{col}{linha_ref}"
        val = ws_report.range(cell).value
        if val in (None, ""):
            return 0.0

        if isinstance(val, (int, float)):
            return float(val)

        s = str(val).strip()
        # Extrai numero monetario quando a celula traz rotulo + valor.
        m = re.search(r"\d{1,3}(?:[\.\s]\d{3})*,\d{2}|\d+\.\d{2}", s)
        if m:
            try:
                return self._br_or_us_to_float(m.group(0))
            except Exception:
                return 0.0

        print(f"⚠️ Tarifa base invalida em {cell}: {val!r}. Assumindo 0,00")
        return 0.0



    def preencher_tarifa_por_linha(self, ws_report, linha_base: int, n: int, status: str, coluna_saida: str = "G"):
        """
        LÃª data em C{linha} e perÃ­odo em E{linha}.
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
        Gera sequÃªncia entre inÃ­cio e fim, respeitando final diferente.
        """
        seq = ["01x07", "07x13", "13x19", "19x01"]
        if periodo_inicial not in seq or periodo_final not in seq:
            # fallback: devolve sÃ³ inicial se algo vier fora do padrÃ£o
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
        Extrai valores monetÃ¡rios diretamente das linhas do OGMO que contenham horÃ¡rio.
        Retorna mapa no formato: {"09x13": [valor1, valor2], "13x19": [valor]}
        """
        mapa: dict[str, list[float]] = {}
        rx_h = re.compile(r"\b(\d{1,2})\s*[xÃ—h\-:]\s*(\d{1,2})\b", re.I)
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
    # REPORT VIGIA - PADRÃƒO (Aquarius e outros)
    # ==================================================
    def colar_report_padrao(self, wb):
        aba = self._achar_aba(wb, ["report vigia"])
        print("ðŸ“Œ Report PADRÃƒO â€“ Outros Clientes")

        if len(self.caminhos_pdfs) >= 2:
            data_ini, data_fim, periodo_inicial, periodo_final = self.extrair_periodo_mesclado_n()
        else:
            data_ini, data_fim = self.extrair_periodo_por_data()
            periodo_inicial, periodo_final = self.extrair_periodo_por_horario()


        print("DEBUG extraÃ§Ã£o:",
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

        # âœ… status pelo nome do navio (o "nome" com (ATRACADO)/(AO LARGO))
        pasta = self.caminhos_pdfs[0].parent
        navio = obter_nome_navio(pasta, None)  # vocÃª jÃ¡ tem
        status = self._status_atracacao(navio)

        # âœ… preenche tarifa por linha usando C e E como base
        self.preencher_tarifa_por_linha(aba, linha_base, n, status=status, coluna_saida="G")

        # âœ… sobrescreve com valor real do OGMO quando disponÃ­vel (horÃ¡rio + valor na mesma linha)
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
                print(f"âœ” {sobrescritos} valor(es) de perÃ­odo sobrescrito(s) com extraÃ§Ã£o direta do OGMO.")

        print(f"âœ” Colado {n} perÃ­odos + tarifa (status={status}) a partir de C{linha_base}/E{linha_base}")


    def gerar_periodos_report_padrao_ssz_por_dia(self, data_ini, data_fim, periodo_inicial, periodo_final):
        ordem = ["07x13", "13x19", "19x01", "01x07"]

        def norm_periodo(p: str) -> str:
            p = (p or "").strip().lower().replace(" ", "")
            p = p.replace("h", "")
            p = p.replace("-", "x").replace("Ã—", "x")
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
            raise ValueError(f"Data invÃ¡lida: {d!r}")

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
            raise ValueError(f"PerÃ­odo inicial invÃ¡lido: {periodo_inicial!r} -> {p_ini_real!r}")
        if p_fim not in ordem:
            raise ValueError(f"PerÃ­odo final invÃ¡lido: {periodo_final!r} -> {p_fim_real!r}")

        d_ini = to_date(data_ini)
        d_fim = to_date(data_fim)
        if d_fim < d_ini:
            raise ValueError(f"Data final menor que inicial: {d_ini} > {d_fim}")

        out = []
        dia = d_ini

        while dia <= d_fim:
            # MantÃ©m sua regra: em dias â€œdo meioâ€, comeÃ§a sempre em 07x13
            inicio = p_ini if dia == d_ini else "07x13"

            # No Ãºltimo dia, termina no perÃ­odo final; caso contrÃ¡rio, vai atÃ© 01x07
            fim = p_fim if dia == d_fim else "01x07"

            for p in seq_entre(inicio, fim):
                out.append((dia, p))  # mantÃ©m 01x07 no mesmo dia (como vocÃª jÃ¡ faz)

            dia += timedelta(days=1)

            if len(out) > 400:
                raise RuntimeError("ProteÃ§Ã£o: perÃ­odos demais gerados. Verifique datas/perÃ­odos extraÃ­dos.")

        # preserva horÃ¡rios reais nas bordas quando forem atÃ­picos
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
        raise ValueError(f"Data invÃ¡lida para Excel: {d!r}")


    def extrair_periodo_mesclado_n(self) -> tuple[str, str, str, str]:
        """
        Retorna (data_ini, data_fim, periodo_ini, periodo_fim)
        usando:
        - OGMO menor nÃºmero = inicio
        - OGMO maior nÃºmero = fim
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
                f"NÃ£o consegui extrair DATA INICIAL do OGMO {self._numero_ogmo(p_ini.name)} ({p_ini.name}). Erro: {e}"
            ) from e

        try:
            _, df = self.extrair_periodo_por_data(p_fim.name)
        except Exception as e:
            raise RuntimeError(
                f"NÃ£o consegui extrair DATA FINAL do OGMO {self._numero_ogmo(p_fim.name)} ({p_fim.name}). Erro: {e}"
            ) from e

        try:
            pi, _ = self.extrair_periodo_por_horario(p_ini.name)
        except Exception as e:
            raise RuntimeError(
                f"NÃ£o consegui extrair PERÃODO INICIAL do OGMO {self._numero_ogmo(p_ini.name)} ({p_ini.name}). Erro: {e}"
            ) from e

        try:
            _, pf = self.extrair_periodo_por_horario(p_fim.name)
        except Exception as e:
            raise RuntimeError(
                f"NÃ£o consegui extrair PERÃODO FINAL do OGMO {self._numero_ogmo(p_fim.name)} ({p_fim.name}). Erro: {e}"
            ) from e

        print(f"âœ” Data inicial de: {p_ini.name} -> {di} ({pi})")
        print(f"âœ” Data final de:   {p_fim.name} -> {df} ({pf})")

        return di, df, pi, pf



    # --------------------------------------------------
    # 1) status ATRACADO / AO LARGO pelo nome
    # --------------------------------------------------
    def _status_atracacao(self, nome: str) -> str | None:
        if not nome:
            return None

        s = str(nome).upper()

        # se tiver parÃªnteses, pega dentro; se nÃ£o, usa tudo
        m = re.search(r"\((.*?)\)", s)
        dentro = m.group(1).strip() if m else s

        dentro = dentro.replace("-", " ").replace("_", " ")
        dentro = re.sub(r"\s+", " ", dentro)

        if "ATRAC" in dentro:
            return "ATRACADO"
        if "FUNDE" in dentro or "FUND" in dentro:   # âœ… FUNDEIO/FUND = AO LARGO
            return "AO_LARGO"
        if "AO LARGO" in dentro or "A LARGO" in dentro or "LARGO" in dentro:
            return "AO_LARGO"

        return None

    # --------------------------------------------------
    # 2) dia/noite pelo perÃ­odo OGMO (coluna E)
    # --------------------------------------------------
    def _is_noite_por_periodo(self, periodo: str) -> bool:
        p = (periodo or "").strip().upper().replace(" ", "")
        # noite: 19x01 e 01x07
        return p in ("19X01", "01X07", "19x01", "01x07")


    # --------------------------------------------------
    # 3) domingo/feriado (mÃ­nimo viÃ¡vel)
    #    (se vocÃª jÃ¡ tiver funÃ§Ã£o de feriado no projeto, plugue aqui)
    # --------------------------------------------------
    def _is_domingo_ou_feriado(self, d: date) -> bool:
        if isinstance(d, datetime):
            d = d.date()
        # domingo
        if d.weekday() == 6:
            return True

        # âœ… feriados nacionais fixos (mÃ­nimo)
        fixos = {
            (1, 1),    # ConfraternizaÃ§Ã£o Universal
            (4, 21),   # Tiradentes
            (5, 1),    # Dia do Trabalho
            (9, 7),    # IndependÃªncia
            (10, 12),  # Nossa Sra Aparecida
            (11, 2),   # Finados
            (11, 15),  # ProclamaÃ§Ã£o da RepÃºblica
            (12, 25),  # Natal
        }
        if (d.month, d.day) in fixos:
            return True

        # Se vocÃª quiser incluir feriados mÃ³veis (Carnaval/PaixÃ£o/Corpus Christi),
        # eu adiciono um cÃ¡lculo de PÃ¡scoa e derivados aqui.
        return False


    # --------------------------------------------------
    # 4) pega a tarifa da tabela visivel do REPORT VIGIA:
    #    - ATRACADO: linha 9
    #    - AO LARGO/FUNDEIO: linha 16
    #    - Seg-Sáb dia: Q
    #    - Seg-Sáb noite: S
    #    - Dom/Feriado dia: U
    #    - Dom/Feriado noite: W
    # --------------------------------------------------
    def _coluna_tarifa_por_regra(self, dom_fer: bool, noite: bool) -> str:
        if not dom_fer and not noite:
            return "Q"
        if not dom_fer and noite:
            return "S"
        if dom_fer and not noite:
            return "U"
        return "W"

    def _tarifa_atracado(self, ws_report, d: date, periodo: str) -> float:
        dom_fer = self._is_domingo_ou_feriado(d)
        noite = self._is_noite_por_periodo(periodo)

        cell = f"{self._coluna_tarifa_por_regra(dom_fer, noite)}9"

        val = ws_report.range(cell).value
        if val in (None, ""):
            return 0.0
        if isinstance(val, (int, float)):
            return float(val)
        m = re.search(r"\d{1,3}(?:[\.\s]\d{3})*,\d{2}|\d+\.\d{2}", str(val))
        if m:
            try:
                return self._br_or_us_to_float(m.group(0))
            except Exception:
                return 0.0
        print(f"⚠️ Tarifa ATRACADO invalida em {cell}: {val!r}. Assumindo 0,00")
        return 0.0

    def _tarifas_da_linha(self, ws_report, linha_ref: int, col_ini: int = 17, col_fim: int = 23) -> list[float]:
        """
        Extrai valores monetarios da linha de tarifa (Q..W por padrao).
        Usado quando a celula base vem com texto (ex.: 'CUSTO').
        """
        valores: list[float] = []
        rx = re.compile(r"\d{1,3}(?:[\.\s]\d{3})*,\d{2}|\d+\.\d{2}")

        for col_idx in range(col_ini, col_fim + 1):
            try:
                v = ws_report.range((linha_ref, col_idx)).value
            except Exception:
                continue

            if v in (None, ""):
                continue

            if isinstance(v, (int, float)):
                valores.append(float(v))
                continue

            m = rx.search(str(v))
            if not m:
                continue
            try:
                valores.append(self._br_or_us_to_float(m.group(0)))
            except Exception:
                continue

        return valores

    def _escolher_tarifa_da_lista(self, valores: list[float], dom_fer: bool, noite: bool) -> float:
        if not valores:
            return 0.0

        idx = 0
        if not dom_fer and not noite:
            idx = 0
        elif not dom_fer and noite:
            idx = 1
        elif dom_fer and not noite:
            idx = 2
        else:
            idx = 3

        if len(valores) >= 4:
            return float(valores[idx])
        if len(valores) >= 2:
            return float(valores[1 if noite else 0])
        return float(valores[0])


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

            cell = f"{self._coluna_tarifa_por_regra(dom_fer, noite)}{linha_ref}"

            val = ws_report.range(cell).value
            if val in (None, ""):
                tarifa_base = 0.0
            elif isinstance(val, (int, float)):
                tarifa_base = float(val)
            else:
                m = re.search(r"\d{1,3}(?:[\.\s]\d{3})*,\d{2}|\d+\.\d{2}", str(val))
                if m:
                    try:
                        tarifa_base = self._br_or_us_to_float(m.group(0))
                    except Exception:
                        tarifa_base = 0.0
                else:
                    # AQUARIUS: tabela de tarifas traz 'CUSTO' na base; pega valores da linha inteira.
                    tarifas_linha = self._tarifas_da_linha(ws_report, linha_ref)
                    tarifa_base = self._escolher_tarifa_da_lista(tarifas_linha, dom_fer, noite)
                    if tarifa_base == 0.0:
                        print(f"⚠️ Tarifa base invalida em {cell}: {val!r}. Assumindo 0,00")
                    else:
                        print(f"ℹ️ Tarifa obtida da tabela da linha {linha_ref}: {tarifa_base:.2f}")
            ws_report.range(f"{coluna_saida}{linha}").value = tarifa_base


        print("DEBUG status:", status)

    # ==================================================
    # EXECUÃ‡ÃƒO PRINCIPAL
    # ==================================================

    def executar(self, preview=False, selection=None):
        if selection and isinstance(selection, dict) and selection.get("pdfs"):
            self.caminhos_pdfs = [Path(p) for p in selection["pdfs"]]
        else:
            self.selecionar_pdfs_ogmo()
        self.carregar_pdfs()   # jÃ¡ faz pdfplumber e OCR sÃ³ se precisar
        self.normalizar_texto_mantendo_linhas()




        cliente, porto = self.identificar_cliente_e_porto()
        print(f"\nðŸš¢ FATURAMENTO OGMO â€“ {cliente} / {porto}")

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
            # âš ï¸ NÃƒO atualiza planilha de controle aqui (serÃ¡ feito apÃ³s preview)

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

            # âœ… SALVAR EXCEL (com wb aberto)
            caminho_excel = salvar_excel_com_nome(wb, pasta, nome_base)
            print(f"ðŸ’¾ Excel salvo em: {caminho_excel}")

            # âœ… Verifica se Ã© cliente WILLIAMS (apenas FRONT VIGIA no PDF)
            apenas_front = "WILLIAMS" in cliente.upper()

            # âœ… GERAR PDF SEM REABRIR O EXCEL (evita erro COM)
            caminho_pdf = gerar_pdf_do_wb_aberto(
                wb, pasta, nome_base, ignorar_abas=("NF",), apenas_front=apenas_front
            )

            nome_cliente = f"{cliente} - {porto}" if porto != "PADRAO" else cliente
            nome_cliente_norm = cliente.strip().upper()
            if nome_cliente_norm in ("ROCHAMAR", "CARGONAVE"):
                anexos = [caminho_pdf]
            else:
                # Demais clientes: somente folhas OGMO (sem faturamento principal)
                anexos = []
            anexos.extend(self.caminhos_pdfs)
            try:
                criar_rascunho_email_cliente(
                    nome_cliente,
                    anexos=anexos,
                    dn=str(nd),
                    navio=navio,
                    usuario_nome=self.usuario_nome,
                )
                print("âœ… Rascunho do Outlook criado com anexos.")
            except Exception as e:
                print(f"âš ï¸ Nao foi possivel criar rascunho do Outlook: {e}")

            print("âœ… FATURAMENTO FINALIZADO")

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

