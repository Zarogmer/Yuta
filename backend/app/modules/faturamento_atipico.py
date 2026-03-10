from backend.app.yuta_helpers import *
from .faturamento_completo import FaturamentoCompleto


class FaturamentoAtipico(FaturamentoCompleto):
    """
    Faturamento ATÃPICO:
    - NÃƒO gera ciclos por regra.
    - LÃª linhas reais do RESUMO (NAVIO): B=data, C=periodo, Z=valor
    - Replica no REPORT VIGIA: C=data, E=periodo, G=valor
    - âœ… Corrige ordem: Data crescente + perÃ­odo (06x12,12x18,18x24,00x06)
    """

    # ordem oficial dos perÃ­odos no REPORT
    _RANK_PERIODO = {
        "06x12": 0,
        "12x18": 1,
        "18x24": 2,
        "00x06": 3,
    }

    # -------------------------
    # NormalizaÃ§Ã£o robusta do perÃ­odo vindo da coluna C
    # Aceita: "06h", "12h", "18h", "00h" e tambÃ©m "06x12"
    # -------------------------
    def normalizar_periodo_c(self, valor_c) -> str | None:
        if not valor_c:
            return None

        s = str(valor_c).strip().lower()
        s = s.replace(" ", "")

        # âœ… caso tÃ­pico do teu atÃ­pico: "06h", "12h", "18h", "00h"
        m_h = re.match(r"^(\d{1,2})h$", s)
        if m_h:
            hh = int(m_h.group(1)) % 24
            mapa = {0: "00x06", 6: "06x12", 12: "12x18", 18: "18x24"}
            return mapa.get(hh)

        # âœ… aceita "06", "12", "18", "00" (Ã s vezes vem sem 'h')
        if re.match(r"^\d{1,2}$", s):
            hh = int(s) % 24
            mapa = {0: "00x06", 6: "06x12", 12: "12x18", 18: "18x24"}
            return mapa.get(hh)

        # âœ… aceita formatos com x/h/-, etc: "06x12", "06-12", "06:12", "06h12"
        s = s.replace("h", "")
        s = s.replace(":", "x").replace("-", "x").replace("Ã—", "x")
        s = re.sub(r"[^0-9x]", "", s)

        m = re.match(r"^(\d{1,2})x(\d{1,2})$", s)
        if not m:
            return None

        a = int(m.group(1)) % 24
        b = int(m.group(2)) % 24
        periodo = f"{a:02d}x{b:02d}"

        return periodo if periodo in self._RANK_PERIODO else None

    # -------------------------
    # Extrai linhas reais do RESUMO (B,C,Z) e jÃ¡ devolve ORDENADO
    # -------------------------
    def extrair_linhas_atipico_resumo(self, ws_resumo, linha_inicio=2):
        last_row = ws_resumo.used_range.last_cell.row

        col_b = ws_resumo.range(f"B{linha_inicio}:B{last_row}").value
        col_c = ws_resumo.range(f"C{linha_inicio}:C{last_row}").value
        col_z = ws_resumo.range(f"Z{linha_inicio}:Z{last_row}").value

        if not isinstance(col_b, list): col_b = [col_b]
        if not isinstance(col_c, list): col_c = [col_c]
        if not isinstance(col_z, list): col_z = [col_z]

        linhas = []
        data_atual = None

        for i in range(len(col_b)):
            b = col_b[i]
            c = col_c[i]
            z = col_z[i]

            # ignora linhas "Total"
            if isinstance(c, str) and c.strip().lower().startswith("total"):
                continue

            # atualiza data quando B vem preenchido
            if isinstance(b, datetime):
                data_atual = b.date()
            elif isinstance(b, date):
                data_atual = b
            elif isinstance(b, str) and b.strip():
                try:
                    data_atual = datetime.strptime(b.strip(), "%d/%m/%Y").date()
                except:
                    pass

            if not data_atual:
                continue

            periodo = self.normalizar_periodo_c(c)
            if not periodo:
                continue

            try:
                valor = self.extrair_numero_excel(z)
            except:
                continue

            # guarda tambÃ©m Ã­ndice original pra desempate
            linhas.append((data_atual, periodo, float(valor), i))

        if not linhas:
            return []

        # âœ… AQUI estÃ¡ a correÃ§Ã£o da ordem:
        # 1) data crescente
        # 2) perÃ­odo na ordem fixa (06x12,12x18,18x24,00x06)
        # 3) desempate pelo Ã­ndice original (mantÃ©m estabilidade)
        linhas.sort(key=lambda x: (x[0], self._RANK_PERIODO.get(x[1], 99), x[3]))

        # remove o Ã­ndice antes de devolver
        return [(d, p, v) for (d, p, v, _) in linhas]

    # -------------------------
    # Monta o REPORT VIGIA baseado nas linhas extraÃ­das
    # -------------------------
    def montar_report_atipico(self):
        ws_report = self.wb2.sheets["REPORT VIGIA"]
        ws_resumo = self.ws1

        linhas = self.extrair_linhas_atipico_resumo(ws_resumo, linha_inicio=2)
        if not linhas:
            raise RuntimeError("ATÃPICO: nÃ£o encontrei linhas (B/C/Z) vÃ¡lidas no RESUMO do NAVIO.")

        linha_base = 22
        n = len(linhas)

        self.inserir_linhas_report(ws_report, linha_inicial=linha_base, periodos=n)

        for i, (d, p, v) in enumerate(linhas):
            linha = linha_base + i
            celula_data = ws_report.range(f"C{linha}")
            celula_data.value = d
            celula_data.number_format = "[$-en-US]mmmm d, aaaa"
            ws_report.range(f"E{linha}").value = p
            cell = ws_report.range(f"G{linha}")
            cell.value = v
            cell.api.NumberFormatLocal = "R$ #.##0,00"

        d_min = min(x[0] for x in linhas)
        d_max = max(x[0] for x in linhas)
        self.ws_front.range("D16").value = self.data_por_extenso(d_min)
        self.ws_front.range("D17").value = self.data_por_extenso(d_max)

        print(f"âœ… ATÃPICO: Report montado e ORDENADO com {n} linhas.")
        return linhas

    def processar(self):
        self.preencher_front_vigia()
        self.processar_MMO(self.wb1, self.wb2)
        self.montar_report_atipico()

        self.arredondar_para_baixo_50_se_cargonave()
        self.gerar_recibo_cargonave_word()
        self.gerar_planilha_calculo_cargonave()
        self.gerar_planilha_calculo_conesul()

        print("âœ… FATURAMENTO ATÃPICO finalizado com sucesso!")
        
        # âœ… Atualizar planilha de controle APÃ“S tudo estar pronto
        self.atualizar_planilha_controle()

    def _preview_title(self):
        return "Faturamento (Atipico)"

    def processar_preview(self):
        self.preencher_front_vigia()
        # âš ï¸ NÃƒO atualiza planilha de controle no preview (evita duplicaÃ§Ã£o)
        self.processar_MMO(self.wb1, self.wb2)
        linhas = self.montar_report_atipico()
        return len(linhas)

