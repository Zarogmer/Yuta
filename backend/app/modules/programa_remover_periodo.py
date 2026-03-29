from datetime import date, datetime
from pathlib import Path
from tempfile import gettempdir

import xlwings as xw

from backend.app.yuta_helpers import fechar_workbooks, feriados_br, selecionar_arquivo_navio


class ProgramaRemoverPeriodo:
    def __init__(self, debug=False):
        self.debug = debug
        self.caminho_navio = None
        self.app = None
        self.wb = None
        self.wb_navio = None
        self.wb_cliente = None
        self.ws = None
        self.datas = []

    PERIODOS_MENU = {"1": "06h", "2": "12h", "3": "18h", "4": "00h"}
    MAPA_PERIODOS = {
        "06h": "06h", "6h": "06h", "06": "06h",
        "12h": "12h", "12": "12h",
        "18h": "18h", "18": "18h",
        "00h": "00h", "0h": "00h", "00": "00h", "24h": "00h"
    }

    # ---------------------------
    # Abrir arquivo NAVIO
    # ---------------------------

    def abrir_arquivo_navio(self, caminho=None):
        caminho = caminho or self.caminho_navio or selecionar_arquivo_navio()
        if not caminho:
            return

        self.caminho_navio = caminho

        self.app = xw.App(visible=False, add_book=False)
        self.wb_navio = self.app.books.open(caminho)
        self.wb = self.wb_navio
        self.ws = self.wb.sheets[0]

    # ---------------------------
    # Utilidades
    # ---------------------------

    def normalizar_texto(self, texto):
        return str(texto).lower().replace(" ", "")

    def normalizar_periodo(self, texto):
        return self.MAPA_PERIODOS.get(self.normalizar_texto(texto))

    def parse_data(self, data_str):
        return datetime.strptime(data_str, "%d/%m/%Y")

    def is_dia_bloqueado(self, data_str):
        data = datetime.strptime(data_str, "%d/%m/%Y").date()
        if data.weekday() == 6:
            return True
        if data in feriados_br:
            return True
        return False

    def obter_nome_navio(self):
        return self.ws.range("A2").value




    # ---------------------------
    # Datas
    # ---------------------------

    def carregar_datas(self):
        ultima = self.ws.range("B" + str(self.ws.cells.last_cell.row)).end("up").row
        datas = []
        for i in range(1, ultima + 1):
            v = self.ws.range(f"B{i}").value
            if isinstance(v, (datetime, date)):
                datas.append(v.strftime("%d/%m/%Y"))
            elif isinstance(v, str) and "/" in v:
                datas.append(v.strip())

        self.datas = list(dict.fromkeys(datas))

    def escolher_data(self):
        print("\nDatas disponÃ­veis:")
        for i, d in enumerate(self.datas, 1):
            print(f"{i} - {d}")
        while True:
            try:
                return self.datas[int(input("Escolha a data: ")) - 1]
            except:
                print("OpÃ§Ã£o invÃ¡lida.")

    def escolher_periodo(self):
        print("\nHorÃ¡rio:")
        print("1 = 06h | 2 = 12h | 3 = 18h | 4 = 00h")
        while True:
            op = input("OpÃ§Ã£o: ").strip()
            if op in self.PERIODOS_MENU:
                return self.PERIODOS_MENU[op]

    # ---------------------------
    # LocalizaÃ§Ã£o
    # ---------------------------

    def encontrar_linha_data(self, data_str):
        ultima = self.ws.range("B" + str(self.ws.cells.last_cell.row)).end("up").row
        for i in range(1, ultima + 1):
            v = self.ws.range(f"B{i}").value
            if isinstance(v, (datetime, date)) and v.strftime("%d/%m/%Y") == data_str:
                return i
            elif v == data_str:
                return i
        return None

    def encontrar_total_data(self, linha_data):
        i = linha_data + 1
        while True:
            valor_c = self.ws.range(f"C{i}").value
            valor_a = self.ws.range(f"A{i}").value

            if isinstance(valor_a, str) and self.normalizar_texto(valor_a) == "totalgeral":
                return None

            if isinstance(valor_c, str) and self.normalizar_texto(valor_c) == "total":
                return i

            i += 1

    def encontrar_linha_total_geral(self):
        ultima = self.ws.cells.last_cell.row
        for i in range(1, ultima + 1):
            v = self.ws.range(f"A{i}").value
            if isinstance(v, str) and self.normalizar_texto(v) == "totalgeral":
                return i
        return None

    # ---------------------------
    # Encontrar perÃ­odo EXATO
    # ---------------------------




    def encontrar_linha_periodo(self, data, periodo):
        linha_data = self.encontrar_linha_data(data)
        if not linha_data:
            return None

        # ðŸ”´ REGRA ESPECIAL PARA 00h
        if periodo == "00h":
            linha_acima = linha_data - 1
            if linha_acima > 0:
                valor_c = self.ws.range(f"C{linha_acima}").value
                if isinstance(valor_c, str):
                    p = self.normalizar_periodo(valor_c)
                    if p == "00h":
                        return linha_acima

        # ðŸ”½ Procura normal abaixo da data
        i = linha_data + 1
        while True:
            valor_a = self.ws.range(f"A{i}").value
            valor_c = self.ws.range(f"C{i}").value

            if isinstance(valor_a, str) and self.normalizar_texto(valor_a) == "totalgeral":
                return None

            if isinstance(valor_c, str):
                p = self.normalizar_periodo(valor_c)
                if p == periodo:
                    return i

                if self.normalizar_texto(valor_c) == "total":
                    return None

            i += 1


    # ---------------------------
    # SubtraÃ§Ãµes
    # ---------------------------

    def subtrair_total_dia(self, linha_origem, linha_total_dia):
        ultima_col = self.ws.range("A1").expand("right").last_cell.column
        for col in range(3, ultima_col + 1):
            v = self.ws.range((linha_origem, col)).value
            if isinstance(v, (int, float)):
                celula = self.ws.range((linha_total_dia, col))
                celula.value = (celula.value or 0) - v

    def subtrair_total_geral(self, linha_origem):
        linha_total_geral = self.encontrar_linha_total_geral()
        if not linha_total_geral:
            return

        ultima_col = self.ws.range("A1").expand("right").last_cell.column
        for col in range(4, ultima_col + 1):
            v = self.ws.range((linha_origem, col)).value
            if isinstance(v, (int, float)):
                celula = self.ws.range((linha_total_geral, col))
                celula.value = (celula.value or 0) - v

    # ---------------------------
    # Remover perÃ­odo
    # ---------------------------

    def remover_periodo(self, data, periodo):
        if self.is_dia_bloqueado(data):
            print(f"â›” {data} Ã© domingo ou feriado â€” nenhuma aÃ§Ã£o executada")
            return

        linha = self.encontrar_linha_periodo(data, periodo)
        if not linha:
            print(f"â„¹ PerÃ­odo {periodo} nÃ£o existe em {data} â€” nada a remover")
            return

        linha_data = self.encontrar_linha_data(data)
        linha_total_dia = self.encontrar_total_data(linha_data)

        print(f"\nðŸ—‘ Removendo perÃ­odo {periodo} â€” Data {data}")

        if linha_total_dia:
            self.subtrair_total_dia(linha, linha_total_dia)

        self.subtrair_total_geral(linha)

        self.ws.api.Rows(linha).Delete()

        print("➡️ Linha removida e totais ajustados")


    # ---------------------------
    # Preview
    # ---------------------------

    def _detectar_area_util_preview(self, max_linhas_scan=1200, max_colunas_scan=180):
        try:
            used = self.ws.api.UsedRange
            used_last_row = int(used.Row + used.Rows.Count - 1)
            used_last_col = int(used.Column + used.Columns.Count - 1)
        except Exception:
            used_last_row = 200
            used_last_col = 40

        scan_rows = max(80, min(max_linhas_scan, used_last_row + 10))
        scan_cols = max(20, min(max_colunas_scan, used_last_col + 6))

        valores = self.ws.range((1, 1), (scan_rows, scan_cols)).value
        if not isinstance(valores, list):
            valores = [[valores]]
        elif valores and not isinstance(valores[0], list):
            valores = [valores]

        ultima_linha = 1
        ultima_coluna = 1
        for i in range(scan_rows):
            linha_vals = valores[i] if i < len(valores) else []
            for j in range(scan_cols):
                v = linha_vals[j] if j < len(linha_vals) else None
                tem_valor = v not in (None, "")
                if isinstance(v, str):
                    tem_valor = v.strip() != ""
                if tem_valor:
                    ultima_linha = max(ultima_linha, i + 1)
                    ultima_coluna = max(ultima_coluna, j + 1)

        ultima_linha = min(ultima_linha + 2, scan_rows)
        ultima_coluna = min(ultima_coluna + 1, scan_cols)
        return max(ultima_linha, 1), max(ultima_coluna, 1)

    def _export_preview_pdf(self):
        if not self.ws:
            return None

        nome_base = Path(str(self.caminho_navio or "navio")).stem
        caminho_pdf = Path(gettempdir()) / f"preview_desfazer_ponto_{nome_base}.pdf"
        if caminho_pdf.exists():
            caminho_pdf.unlink()

        try:
            ps = self.ws.api.PageSetup
            xl_landscape = 2
            xl_paper_a4 = 9

            ultima_linha, ultima_coluna = self._detectar_area_util_preview(
                max_linhas_scan=1200,
                max_colunas_scan=180,
            )
            area = self.ws.api.Range(
                self.ws.api.Cells(1, 1),
                self.ws.api.Cells(ultima_linha, ultima_coluna),
            )

            ps.PrintArea = area.Address
            ps.Orientation = xl_landscape
            ps.PaperSize = xl_paper_a4
            ps.Zoom = False
            ps.FitToPagesWide = 1
            ps.FitToPagesTall = False
            ps.CenterHorizontally = True
            ps.CenterVertically = False

            self.ws.activate()
            self.ws.api.ExportAsFixedFormat(
                Type=0,
                Filename=str(caminho_pdf),
                Quality=0,
                IncludeDocProperties=True,
                IgnorePrintAreas=False,
                OpenAfterPublish=False,
            )
            return str(caminho_pdf)
        except Exception:
            return None

    def executar_preview(self, selection=None):
        selection = selection or {}
        data_selecionada = selection.get("data")
        periodo_selecionado = selection.get("periodo")
        caminho_navio = selection.get("caminho_navio")

        if not caminho_navio:
            raise Exception("Arquivo NAVIO nao informado para preview")

        try:
            self.abrir_arquivo_navio(caminho=caminho_navio)
            self.carregar_datas()

            linhas = [
                "PRE-VISUALIZACAO",
                "Processo: Remover Ponto",
                f"Arquivo: {Path(caminho_navio).name}",
                f"Total de datas no arquivo: {len(self.datas)}",
            ]

            if data_selecionada and periodo_selecionado:
                periodo = periodo_selecionado
                if periodo not in self.MAPA_PERIODOS.values():
                    raise Exception(f"Periodo invalido: {periodo}")

                linhas.append(f"Data: {data_selecionada}")
                linhas.append(f"Periodo: {periodo}")

                if data_selecionada not in self.datas:
                    linhas.append("Status: data nao encontrada no arquivo")
                elif not self.encontrar_linha_periodo(data_selecionada, periodo):
                    linhas.append("Status: periodo nao existe nesta data (nada a remover)")
                else:
                    if self.is_dia_bloqueado(data_selecionada):
                        linhas.append("Regra: domingo/feriado - remocao bloqueada.")
                    else:
                        linhas.append("Status: periodo encontrado, pronto para remover")
            else:
                linhas.append("Status: selecione data e periodo ao clicar em 'Executar'.")

            return {
                "text": "\n".join(linhas),
                "preview_pdf": self._export_preview_pdf(),
                "selection": {
                    "caminho_navio": caminho_navio,
                    "datas": list(self.datas or []),
                },
            }
        finally:
            fechar_workbooks(
                app=self.app,
                wb_navio=self.wb_navio,
                wb_cliente=self.wb_cliente,
            )


    # ---------------------------
    # ExecuÃ§Ã£o
    # ---------------------------

    def executar(self, usar_arquivo_aberto=False, selection=None):
        try:
            selection = selection or {}
            data_selecionada = selection.get("data")
            periodo_selecionado = selection.get("periodo")
            caminho_navio = selection.get("caminho_navio")

            if not usar_arquivo_aberto or not self.ws:
                self.abrir_arquivo_navio(caminho=caminho_navio)

            if not self.ws:
                return

            self.carregar_datas()

            data = data_selecionada if data_selecionada else self.escolher_data()
            if data not in self.datas:
                raise Exception(f"Data invÃ¡lida para este arquivo: {data}")

            periodo = periodo_selecionado if periodo_selecionado else self.escolher_periodo()
            if periodo not in self.MAPA_PERIODOS.values():
                raise Exception(f"PerÃ­odo invÃ¡lido: {periodo}")

            self.remover_periodo(data, periodo)

            self.salvar()

        finally:
            if not usar_arquivo_aberto:
                fechar_workbooks(
                    app=self.app,
                    wb_navio=self.wb_navio,
                    wb_cliente=self.wb_cliente
                )

    def salvar(self):
        if self.wb:
            self.wb.save()
            print("ðŸ’¾ Arquivo salvo com sucesso")

