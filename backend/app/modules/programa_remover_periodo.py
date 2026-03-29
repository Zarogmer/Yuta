import shutil
import time
from datetime import date, datetime
from pathlib import Path
from tempfile import gettempdir

import xlwings as xw

from backend.app.yuta_helpers import fechar_workbooks, feriados_br, selecionar_arquivo_navio


class ProgramaRemoverPeriodo:
    def __init__(self, debug=False):
        self.debug = debug
        self.caminho_navio = None
        self._caminho_navio_destino = None
        self._caminho_navio_temp = None
        self._save_copy_path = None
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
            raise FileNotFoundError("Arquivo do NAVIO nao selecionado")

        caminho_destino = Path(caminho).resolve()
        self._caminho_navio_destino = str(caminho_destino)
        self._caminho_navio_temp = None
        self.caminho_navio = str(caminho_destino)

        self.app = xw.App(visible=False, add_book=False)
        self.wb_navio = self.app.books.open(str(caminho_destino))

        # Se abrir em somente leitura (rede/OneDrive), usa copia local de trabalho.
        if bool(getattr(self.wb_navio.api, "ReadOnly", False)):
            print("Arquivo abriu em SOMENTE LEITURA. Usando copia local temporaria para edicao.")
            try:
                self.wb_navio.close()
            except Exception:
                pass

            temp_name = f"remover_ponto_work_{caminho_destino.stem}_{int(time.time() * 1000)}.xlsx"
            caminho_temp = Path(gettempdir()) / temp_name
            if caminho_temp.exists():
                caminho_temp.unlink()
            shutil.copy2(caminho_destino, caminho_temp)

            self.wb_navio = self.app.books.open(str(caminho_temp))
            self._caminho_navio_temp = str(caminho_temp)
            self.caminho_navio = str(caminho_temp)

        self.wb = self.wb_navio
        self.ws = self._selecionar_worksheet_dados(self.wb)

    def _selecionar_worksheet_dados(self, wb):
        # xlWorksheet = -4167
        for sh in wb.sheets:
            try:
                if int(sh.api.Type) == -4167:
                    _ = sh.range("A1").value
                    return sh
            except Exception:
                continue

        for sh in wb.sheets:
            try:
                _ = sh.range("A1").value
                return sh
            except Exception:
                continue

        raise RuntimeError("Nenhuma aba de planilha valida encontrada no arquivo NAVIO.")

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

    def _to_float(self, valor):
        if isinstance(valor, (int, float)):
            return float(valor)
        if isinstance(valor, str):
            texto = valor.strip().replace(" ", "")
            if not texto:
                return None
            if "," in texto and "." in texto:
                texto = texto.replace(".", "").replace(",", ".")
            elif "," in texto:
                texto = texto.replace(",", ".")
            try:
                return float(texto)
            except Exception:
                return None
        return None

    def _limite_scan_planilha(self):
        try:
            ultima = self.ws.range("B" + str(self.ws.cells.last_cell.row)).end("up").row
            return min(max(ultima + 10, 20), 200)
        except Exception:
            return 200

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
        print("\nDatas disponiveis:")
        for i, d in enumerate(self.datas, 1):
            print(f"{i} - {d}")
        while True:
            try:
                return self.datas[int(input("Escolha a data: ")) - 1]
            except Exception:
                print("Opcao invalida.")

    def escolher_periodo(self):
        print("\nHorario:")
        print("1 = 06h | 2 = 12h | 3 = 18h | 4 = 00h")
        while True:
            op = input("Opcao: ").strip()
            if op in self.PERIODOS_MENU:
                return self.PERIODOS_MENU[op]

    # ---------------------------
    # Localizacao
    # ---------------------------

    def encontrar_linha_data(self, data_str):
        ultima = self._limite_scan_planilha()
        for i in range(1, ultima + 1):
            v = self.ws.range(f"B{i}").value
            if isinstance(v, (datetime, date)) and v.strftime("%d/%m/%Y") == data_str:
                return i
            elif v == data_str:
                return i
        return None

    def encontrar_total_data(self, linha_data):
        i = linha_data + 1
        limite = linha_data + 30
        while i <= limite:
            valor_c = self.ws.range(f"C{i}").value
            valor_a = self.ws.range(f"A{i}").value

            if isinstance(valor_a, str) and self.normalizar_texto(valor_a) == "totalgeral":
                return None

            if isinstance(valor_c, str) and self.normalizar_texto(valor_c) == "total":
                return i

            i += 1
        return None

    def encontrar_linha_total_geral(self):
        ultima = self._limite_scan_planilha()
        for i in range(1, ultima + 1):
            v = self.ws.range(f"A{i}").value
            if isinstance(v, str) and self.normalizar_texto(v) == "totalgeral":
                return i
        return None

    def encontrar_linha_periodo(self, data, periodo):
        linha_data = self.encontrar_linha_data(data)
        if not linha_data:
            return None

        # REGRA ESPECIAL PARA 00h: pode estar na linha acima da data
        if periodo == "00h":
            linha_acima = linha_data - 1
            if linha_acima > 0:
                valor_c = self.ws.range(f"C{linha_acima}").value
                if isinstance(valor_c, str):
                    p = self.normalizar_periodo(valor_c)
                    if p == "00h":
                        return linha_acima

        # Procura a partir da propria linha da data (periodo pode estar na mesma linha)
        linha_total_dia = self.encontrar_total_data(linha_data)
        limite = linha_total_dia if linha_total_dia else linha_data + 30

        for i in range(linha_data, limite):
            valor_c = self.ws.range(f"C{i}").value
            if not isinstance(valor_c, str):
                continue
            p = self.normalizar_periodo(valor_c)
            if p == periodo:
                return i

        return None

    # ---------------------------
    # Contar periodos de um dia
    # ---------------------------

    def _contar_periodos_dia(self, linha_data, linha_total_dia):
        """Conta quantos periodos (06h, 12h, 18h, 00h) existem entre linha_data e linha_total_dia."""
        count = 0
        for i in range(linha_data, linha_total_dia):
            valor_c = self.ws.range(f"C{i}").value
            if isinstance(valor_c, str) and self.normalizar_periodo(valor_c):
                count += 1
        return count

    # ---------------------------
    # Recalculo (do zero, nao por subtracao)
    # ---------------------------

    def recalcular_total_do_dia(self, linha_data):
        linha_total_dia = self.encontrar_total_data(linha_data)
        if not linha_total_dia:
            return

        ultima_col = self.ws.range("A1").expand("right").last_cell.column
        for col in range(4, ultima_col + 1):
            soma_col = 0.0
            for i in range(linha_data, linha_total_dia):
                valor_c = self.ws.range(f"C{i}").value
                if not isinstance(valor_c, str):
                    continue
                periodo = self.normalizar_periodo(valor_c)
                if not periodo:
                    continue
                v = self._to_float(self.ws.range((i, col)).value)
                if v is not None:
                    soma_col += v
            self.ws.range((linha_total_dia, col)).value = soma_col

        print(f"Total do dia recalculado na linha {linha_total_dia}")

    def recalcular_totais_rodape(self):
        ultima_linha = self._limite_scan_planilha()
        ultima_col = self.ws.range("A1").expand("right").last_cell.column

        # Soma todas as linhas de periodos reais (06h, 12h, 18h, 00h)
        soma = {col: 0.0 for col in range(4, ultima_col + 1)}
        for i in range(1, ultima_linha + 1):
            valor_c = self.ws.range(f"C{i}").value
            if not isinstance(valor_c, str):
                continue
            periodo = self.normalizar_periodo(valor_c)
            if not periodo:
                continue
            for col in range(4, ultima_col + 1):
                v = self._to_float(self.ws.range((i, col)).value)
                if v is not None:
                    soma[col] += v

        # Atualiza linhas de rodape finais (Total Geral e Total rodape)
        for i in range(1, ultima_linha + 1):
            a = self.ws.range(f"A{i}").value
            b = self.ws.range(f"B{i}").value
            c = self.ws.range(f"C{i}").value

            eh_total_geral = isinstance(a, str) and self.normalizar_texto(a) == "totalgeral"
            eh_total_rodape = (
                isinstance(b, str)
                and self.normalizar_texto(b) == "total"
                and not (isinstance(c, str) and self.normalizar_texto(c) == "total")
            )

            if not (eh_total_geral or eh_total_rodape):
                continue

            for col in range(4, ultima_col + 1):
                self.ws.range((i, col)).value = soma[col]

        print("Totais de rodape (Total / Total Geral) recalculados")

    # ---------------------------
    # Remover periodo
    # ---------------------------

    def remover_periodo(self, data, periodo):
        linha = self.encontrar_linha_periodo(data, periodo)
        if not linha:
            return {
                "changed": False,
                "message": f"Periodo {periodo} nao existe em {data} - nada a remover.",
            }

        linha_data = self.encontrar_linha_data(data)
        linha_total_dia = self.encontrar_total_data(linha_data)

        qtd_periodos = self._contar_periodos_dia(linha_data, linha_total_dia) if linha_total_dia else 1

        print(f"\nRemovendo periodo {periodo} - Data {data}")

        if qtd_periodos <= 1:
            # Ultimo periodo do dia: remover bloco inteiro (data + periodo + total)
            # Determina range de linhas a deletar
            primeira_linha = linha_data
            # 00h pode estar acima da data
            if linha > 0 and linha < linha_data:
                primeira_linha = linha
            ultima_linha_bloco = linha_total_dia if linha_total_dia else linha

            qtd_linhas = ultima_linha_bloco - primeira_linha + 1
            self.ws.api.Rows(f"{primeira_linha}:{ultima_linha_bloco}").Delete()
            print(f"Dia inteiro removido ({qtd_linhas} linhas: {primeira_linha} a {ultima_linha_bloco})")
        else:
            # Mais de 1 periodo: remover apenas a linha do periodo
            self.ws.api.Rows(linha).Delete()
            print(f"Linha {linha} removida")

            # Recalcular total do dia
            # Re-encontrar linha_data pois pode ter mudado se linha < linha_data
            linha_data_atual = self.encontrar_linha_data(data)
            if linha_data_atual:
                self.recalcular_total_do_dia(linha_data_atual)

        # Recalcular Total e Total Geral do rodape
        self.recalcular_totais_rodape()

        # Recarregar datas
        self.carregar_datas()

        return {
            "changed": True,
            "message": f"Periodo {periodo} de {data} removido com sucesso.",
        }

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
        caminho_pdf = Path(gettempdir()) / f"preview_remover_ponto_{nome_base}.pdf"
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

            caminho_rede = str(self._caminho_navio_destino or caminho_navio)

            linhas = [
                "PRE-VISUALIZACAO",
                "Processo: Remover Ponto",
                f"Arquivo: {Path(caminho_rede).name}",
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
                    linhas.append("Status: periodo encontrado, pronto para remover")
            else:
                linhas.append("Status: selecione data e periodo ao clicar em 'Executar'.")

            return {
                "text": "\n".join(linhas),
                "preview_pdf": self._export_preview_pdf(),
                "selection": {
                    "caminho_navio": caminho_rede,
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
    # Execucao
    # ---------------------------

    def executar(self, usar_arquivo_aberto=False, selection=None):
        caminho_final = None
        resultado = None
        try:
            selection = selection or {}
            data_selecionada = selection.get("data")
            periodo_selecionado = selection.get("periodo")
            caminho_navio = selection.get("caminho_navio")

            if not usar_arquivo_aberto or not self.ws:
                self.abrir_arquivo_navio(caminho=caminho_navio)

            if not self.ws:
                return {"changed": False, "message": "Nenhuma planilha encontrada."}

            print(f"Arquivo em uso para Remover Ponto: {self.caminho_navio}")
            if self._caminho_navio_destino:
                print(f"Destino final (rede): {self._caminho_navio_destino}")

            caminho_final = Path(str(self._caminho_navio_destino or self.caminho_navio)).resolve()

            self.carregar_datas()

            data = data_selecionada if data_selecionada else self.escolher_data()
            if data not in self.datas:
                raise Exception(f"Data invalida para este arquivo: {data}")

            periodo = periodo_selecionado if periodo_selecionado else self.escolher_periodo()
            if periodo not in self.MAPA_PERIODOS.values():
                raise Exception(f"Periodo invalido: {periodo}")

            resultado = self.remover_periodo(data, periodo)

            if resultado and resultado.get("changed"):
                self.salvar()

            return resultado

        finally:
            if not usar_arquivo_aberto:
                fechar_workbooks(
                    app=self.app,
                    wb_navio=self.wb_navio,
                    wb_cliente=self.wb_cliente
                )

                # Write-through: copia o SaveCopyAs para o arquivo final na rede
                if self._save_copy_path and caminho_final:
                    try:
                        shutil.copy2(self._save_copy_path, caminho_final)
                        time.sleep(0.15)
                        print(f"Write-through aplicado: {caminho_final}")
                    finally:
                        try:
                            if Path(self._save_copy_path).exists():
                                Path(self._save_copy_path).unlink()
                        except Exception:
                            pass
                        self._save_copy_path = None

                if self._caminho_navio_temp:
                    try:
                        p_temp = Path(self._caminho_navio_temp)
                        if p_temp.exists():
                            p_temp.unlink()
                    except Exception:
                        pass
                    self._caminho_navio_temp = None

    def salvar(self):
        if not self.wb:
            raise Exception("Nenhum workbook aberto para salvar")

        try:
            if bool(self.wb.api.ReadOnly):
                raise PermissionError("Arquivo de trabalho abriu em SOMENTE LEITURA.")
        except AttributeError:
            pass

        self.wb.api.Save()

        # Gera copia fisica para write-through apos fechar o workbook
        caminho_final = Path(str(self._caminho_navio_destino or self.caminho_navio or "")).resolve()
        copy_name = f"remover_ponto_write_{caminho_final.stem}_{int(time.time() * 1000)}.xlsx"
        caminho_copy = Path(gettempdir()) / copy_name
        if caminho_copy.exists():
            caminho_copy.unlink()
        self.wb.api.SaveCopyAs(str(caminho_copy))
        self._save_copy_path = str(caminho_copy)

        print(f"Arquivo NAVIO salvo com sucesso em: {caminho_final}")
