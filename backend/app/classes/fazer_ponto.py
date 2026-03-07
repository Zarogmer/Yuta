from datetime import date, datetime
from pathlib import Path
from tempfile import gettempdir
import shutil
import time

import xlwings as xw

from yuta_helpers import (
    fechar_workbooks,
    feriados_br,
    selecionar_arquivo_navio,
)


class FazerPonto:
    def __init__(self, debug=False):
        self.debug = debug
        self.caminho_navio = None
        self.app = None
        self.wb = None
        self.wb_navio = None
        self.wb_cliente = None
        self.ws = None
        self.ws_front = None
        self.pasta_navio = None
        self.datas = []
        self._save_copy_path = None
        self._caminho_navio_destino = None
        self._caminho_navio_temp = None
        self._modelo_info = None


    PERIODOS_MENU = {"1": "06h", "2": "12h", "3": "18h", "4": "00h"}
    MAPA_PERIODOS = {
        "06h": "06h", "6h": "06h", "06": "06h",
        "12h": "12h", "12": "12h",
        "18h": "18h", "18": "18h",
        "00h": "00h", "0h": "00h", "00": "00h", "24h": "00h"
    }
    EQUIVALENTES = {
        "06h": ["06h", "12h"],
        "12h": ["12h", "06h"],
        "18h": ["18h", "00h"],
        "00h": ["00h", "18h"]
    }
    BLOCOS = {"06h": 1, "12h": 1, "18h": 2, "00h": 2}
    def garantir_total_geral_ultima_linha(self):
        """
        Garante que a linha 'Total Geral' permaneça como a última informação da planilha.
        """
        linha_total_geral = self.encontrar_linha_total_geral_opcional()
        if not linha_total_geral:
            return

        # Ultima linha com informacao relevante em A/B/C (estrutura do relatorio)
        ultima_linha_dados = 1
        ultima_scan = self.ws.range("A" + str(self.ws.cells.last_cell.row)).end("up").row
        for i in range(1, ultima_scan + 1):
            a = self.ws.range(f"A{i}").value
            b = self.ws.range(f"B{i}").value
            c = self.ws.range(f"C{i}").value
            if a not in (None, "") or b not in (None, "") or c not in (None, ""):
                ultima_linha_dados = i

        # 'Total Geral' deve ser a ultima informacao. Se houver linhas abaixo, move para o fim.
        if linha_total_geral < ultima_linha_dados:
            self.ws.api.Rows(linha_total_geral).Cut()
            self.ws.api.Rows(ultima_linha_dados + 1).Insert(Shift=-4121)
            if self.debug:
                print(f"📌 'Total Geral' movido para a linha {ultima_linha_dados + 1}")
    ORDEM_PERIODOS = {"00h": 0, "06h": 1, "12h": 2, "18h": 3}


    # ---------------------------
    # Abrir arquivo NAVIO
    # ---------------------------



    def abrir_arquivo_navio(self, caminho=None):
        caminho = caminho or self.caminho_navio or selecionar_arquivo_navio()
        if not caminho:
            raise FileNotFoundError("Arquivo do NAVIO não selecionado")

        caminho_destino = Path(caminho).resolve()
        self._caminho_navio_destino = str(caminho_destino)
        self._caminho_navio_temp = None
        self.caminho_navio = str(caminho_destino)

        self.app = xw.App(visible=False, add_book=False)
        self.wb_navio = self.app.books.open(str(caminho_destino))

        # Se abrir em somente leitura (rede/OneDrive), usa copia local de trabalho.
        if bool(getattr(self.wb_navio.api, "ReadOnly", False)):
            print("⚠️ Arquivo abriu em SOMENTE LEITURA. Usando copia local temporaria para edicao.")
            try:
                self.wb_navio.close()
            except Exception:
                pass

            temp_name = f"fazer_ponto_work_{caminho_destino.stem}_{int(time.time() * 1000)}.xlsx"
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
        """
        Retorna a primeira aba de planilha valida (evita chart sheet/objeto sem Range).
        """
        # xlWorksheet = -4167
        for sh in wb.sheets:
            try:
                if int(sh.api.Type) == -4167:
                    # valida que a aba aceita leitura de celula
                    _ = sh.range("A1").value
                    return sh
            except Exception:
                continue

        # fallback: tenta qualquer aba que aceite range
        for sh in wb.sheets:
            try:
                _ = sh.range("A1").value
                return sh
            except Exception:
                continue

        raise RuntimeError("Nenhuma aba de planilha valida encontrada no arquivo NAVIO.")



    # ---------------------------
    # Datas
    # ---------------------------
    def is_domingo(self, data_str):
        d = datetime.strptime(data_str, "%d/%m/%Y")
        return d.weekday() == 6

    def is_feriado(self, data_str):
        d = datetime.strptime(data_str, "%d/%m/%Y")
        return d in feriados_br

    def is_dia_bloqueado(self, data_str):
        """
        Retorna True se for domingo ou feriado nacional
        data_str no formato DD/MM/YYYY
        """
        data = datetime.strptime(data_str, "%d/%m/%Y").date()

        # Domingo
        if data.weekday() == 6:
            return True

        # Feriado nacional
        if data in feriados_br:
            return True

        return False

    def parse_data(self, data_str):
        return datetime.strptime(data_str, "%d/%m/%Y")

    def normalizar_data_str(self, data_str):
        if data_str is None:
            raise ValueError("Data nao informada")
        texto = str(data_str).strip()
        return datetime.strptime(texto, "%d/%m/%Y").strftime("%d/%m/%Y")

    def normalizar_texto(self, texto):
        return str(texto).lower().replace(" ", "")

    def normalizar_periodo(self, texto):
        t = self.normalizar_texto(texto)
        return self.MAPA_PERIODOS.get(t, None)




    # ---------------------------
    # Datas
    # ---------------------------
    def carregar_datas(self):
        try:
            ultima = self.ws.range("B" + str(self.ws.cells.last_cell.row)).end("up").row
        except Exception:
            # fallback COM direto (xlUp=-4162) quando Range via wrapper falhar
            ultima = int(self.ws.api.Cells(self.ws.api.Rows.Count, 2).End(-4162).Row)
        datas = []
        for i in range(1, ultima + 1):
            v = self.ws.range(f"B{i}").value
            if isinstance(v, (datetime, date)):
                datas.append(v.strftime("%d/%m/%Y"))
            elif isinstance(v, str) and "/" in v:
                datas.append(v.strip())
        self.datas = list(dict.fromkeys(datas))
        if not self.datas:
            raise Exception("Nenhuma data encontrada na coluna B.")

    def escolher_data(self):
        print("\nDatas disponíveis:")
        for i, d in enumerate(self.datas, 1):
            print(f"{i} - {d}")
        while True:
            try:
                return self.datas[int(input("Escolha a data: ")) - 1]
            except:
                print("Opção inválida.")

    def escolher_periodo(self):
        print("\nHorário:")
        print("1 = 06h | 2 = 12h | 3 = 18h | 4 = 00h")
        while True:
            op = input("Opção: ").strip()
            if op in self.PERIODOS_MENU:
                return self.PERIODOS_MENU[op]


    # ---------------------------
    # Localização
    # ---------------------------



    def encontrar_linha_data(self, data_str):
        ultima = self.ws.range("B" + str(self.ws.cells.last_cell.row)).end("up").row
        for i in range(1, ultima + 1):
            valor = self.ws.range(f"B{i}").value
            if isinstance(valor, (datetime, date)) and valor.strftime("%d/%m/%Y") == data_str:
                return i
            elif valor == data_str:
                return i
        raise Exception(f"Data {data_str} não encontrada.")

    def encontrar_total_data(self, linha_data):
        i = linha_data + 1
        while True:
            valor_c = self.ws.range(f"C{i}").value
            valor_a = self.ws.range(f"A{i}").value
            if isinstance(valor_a, str) and self.normalizar_texto(valor_a) == "totalgeral":
                raise Exception("❌ Total do dia não encontrado antes do Total Geral")
            if isinstance(valor_c, str) and self.normalizar_texto(valor_c) == "total":
                return i
            if i > self.ws.cells.last_cell.row:
                raise Exception("❌ Fim da planilha sem encontrar 'Total' do dia")
            i += 1

    def encontrar_linha_periodo(self, data, periodo):
        linha_data = self.encontrar_linha_data(data)
        linha_total_dia = self.encontrar_total_data(linha_data)

        if periodo == "00h":
            linha_acima = linha_data - 1
            if linha_acima > 0:
                valor_c = self.ws.range(f"C{linha_acima}").value
                if isinstance(valor_c, str):
                    p = self.normalizar_periodo(valor_c)
                    if p == "00h":
                        return linha_acima

        for i in range(linha_data, linha_total_dia):
            valor_c = self.ws.range(f"C{i}").value
            if not isinstance(valor_c, str):
                continue
            p = self.normalizar_periodo(valor_c)
            if p == periodo:
                return i

        return None

    def encontrar_linha_insercao_periodo(self, linha_data, linha_total_dia, periodo):
        # Regra operacional: sempre inserir imediatamente acima do TOTAL do dia.
        return linha_total_dia

    def _encontrar_linha_insercao_data(self, data_str):
        data_destino = self.parse_data(data_str)
        linha_total_geral = self.encontrar_linha_total_geral_opcional()

        # Mapeia blocos validos de dia (inicio->total) para evitar inserir no meio do bloco.
        blocos = []
        for data_existente in sorted(self.datas, key=lambda d: self.parse_data(d)):
            try:
                linha_inicio = self.encontrar_linha_data(data_existente)
                linha_total = self.encontrar_total_data(linha_inicio)
                blocos.append(
                    {
                        "data": self.parse_data(data_existente),
                        "linha_inicio": linha_inicio,
                        "linha_total": linha_total,
                    }
                )
            except Exception:
                continue

        if not blocos:
            return 2

        # TOPO: data menor que a primeira do arquivo -> inserir apos cabecalho (linha 1).
        if data_destino < blocos[0]["data"]:
            return 2

        # MEIO: inserir antes do proximo bloco maior.
        for bloco in blocos:
            if bloco["data"] > data_destino:
                return bloco["linha_inicio"]

        # FINAL: data maior que todas -> ancora na coluna C para evitar conflitos.
        try:
            ultima_c = self.ws.range("C" + str(self.ws.cells.last_cell.row)).end("up").row
        except Exception:
            ultima_c = blocos[-1]["linha_total"]

        linha_final = max(ultima_c + 1, blocos[-1]["linha_total"] + 1)
        if linha_total_geral:
            linha_final = min(linha_final, linha_total_geral)
        return max(linha_final, 2)

    def _to_float(self, valor):
        if isinstance(valor, (int, float)):
            return float(valor)
        if isinstance(valor, str):
            texto = valor.strip().replace(" ", "")
            if not texto:
                return None
            # Trata formatos pt-BR e en-US.
            if "," in texto and "." in texto:
                texto = texto.replace(".", "").replace(",", ".")
            elif "," in texto:
                texto = texto.replace(",", ".")
            try:
                return float(texto)
            except Exception:
                return None
        return None

    # ---------------------------
    # Buscar modelo inteligente
    # ---------------------------

    def encontrar_modelo_periodo(self, data_destino, periodo):
        """
        Retorna: (linha_modelo, data_modelo)
        """

        datas_ordenadas = sorted(self.datas, key=lambda d: self.parse_data(d))
        if not datas_ordenadas:
            raise Exception("Nenhuma data valida encontrada para buscar modelo")

        datas_dt = [self.parse_data(d) for d in datas_ordenadas]
        data_destino_dt = self.parse_data(data_destino)

        if data_destino in datas_ordenadas:
            idx = datas_ordenadas.index(data_destino)
        else:
            idx = 0
            while idx < len(datas_dt) and datas_dt[idx] < data_destino_dt:
                idx += 1
            if idx >= len(datas_dt):
                idx = len(datas_dt) - 1

        def procurar_em_data(data, aceitar_equivalente):
            linha_data = self.encontrar_linha_data(data)
            i = linha_data + 1

            while True:
                valor_a = self.ws.range(f"A{i}").value
                valor_c = self.ws.range(f"C{i}").value

                if isinstance(valor_a, str) and self.normalizar_texto(valor_a) == "totalgeral":
                    return None

                if not isinstance(valor_c, str):
                    i += 1
                    continue

                texto = self.normalizar_texto(valor_c)

                if texto == "total":
                    return None

                p = self.normalizar_periodo(texto)
                if not p:
                    i += 1
                    continue

                if p == periodo:
                    self._modelo_info = f"Usando periodo {p} da data {data}"
                    if self.debug:
                        print(f"✔ Usando {p} da data {data}")
                    return i, data

                if aceitar_equivalente and p in self.EQUIVALENTES.get(periodo, []):
                    self._modelo_info = f"Usando equivalente {p} da data {data}"
                    if self.debug:
                        print(f"⚠ Usando equivalente {p} da data {data}")
                    return i, data

                i += 1

        # 1️⃣ Mesmo dia
        if data_destino in datas_ordenadas:
            resultado = procurar_em_data(data_destino, aceitar_equivalente=True)
            if resultado:
                return resultado
        else:
            # Quando a data ainda nao existe, tenta primeiro o dia mais proximo.
            resultado = procurar_em_data(datas_ordenadas[idx], aceitar_equivalente=True)
            if resultado:
                return resultado

        # 2️⃣ Outros dias (duas passagens)
        # Passagem 1: somente período exato
        # Passagem 2: permitir equivalente (ex.: 00h <-> 18h)
        for aceitar_equivalente in (False, True):
            for offset in range(1, len(datas_ordenadas)):
                for novo_idx in (idx - offset, idx + offset):
                    if 0 <= novo_idx < len(datas_ordenadas):
                        data = datas_ordenadas[novo_idx]

                        if self.is_dia_bloqueado(data):
                            if self.debug:
                                print(f"⛔ Pulando data bloqueada: {data}")
                            continue

                        resultado = procurar_em_data(data, aceitar_equivalente=aceitar_equivalente)
                        if resultado:
                            return resultado

        raise Exception(
            f"Nenhum modelo encontrado para o período '{periodo}' "
            f"a partir da data {data_destino}"
        )

    def _criar_novo_dia_com_periodo(self, data, periodo):
        linha_modelo, data_modelo = self.encontrar_modelo_periodo(data, periodo)
        linha_data_modelo = self.encontrar_linha_data(data_modelo)
        linha_total_modelo = self.encontrar_total_data(linha_data_modelo)
        linha_insercao = self._encontrar_linha_insercao_data(data)

        estrategia = "meio"
        datas_ordenadas = sorted(self.datas, key=lambda d: self.parse_data(d))
        if datas_ordenadas:
            if self.parse_data(data) < self.parse_data(datas_ordenadas[0]):
                estrategia = "topo"
            elif self.parse_data(data) > self.parse_data(datas_ordenadas[-1]):
                estrategia = "final"

        print(
            f"\n✅ Criando novo dia {data} no NAVIO - "
            f"Periodo: {periodo} (modelo: {data_modelo})"
        )
        print(f"🧭 Estrategia de insercao: {estrategia} | linha alvo: {linha_insercao}")

        # Estrutura minima do dia: 1 linha de periodo + 1 linha de total.
        self.ws.api.Rows(linha_insercao).Insert()
        self.ws.api.Rows(linha_insercao).Insert()

        if linha_modelo >= linha_insercao:
            linha_modelo += 2
        if linha_total_modelo >= linha_insercao:
            linha_total_modelo += 2

        self.ws.api.Rows(linha_modelo).Copy()
        self.ws.api.Rows(linha_insercao).PasteSpecial(-4163)

        self.ws.api.Rows(linha_total_modelo).Copy()
        self.ws.api.Rows(linha_insercao + 1).PasteSpecial(-4163)

        self.ws.api.Rows(linha_insercao).Font.Bold = True
        cel_data = self.ws.range((linha_insercao, 2))
        # Mantem exatamente o que o usuario digitou (DD/MM/AAAA),
        # sem conversao de locale do Excel.
        try:
            cel_data.number_format = "@"
        except Exception:
            try:
                cel_data.api.NumberFormatLocal = "@"
            except Exception:
                pass
        cel_data.value = str(data)
        self.ws.range((linha_insercao, 3)).value = periodo

        self.ws.range((linha_insercao + 1, 2)).value = None
        self.ws.range((linha_insercao + 1, 3)).value = "Total"

        ultima_col = self.ws.range("A1").expand("right").last_cell.column
        for col in range(4, ultima_col + 1):
            v = self.ws.range((linha_insercao, col)).value
            v_num = self._to_float(v)
            self.ws.range((linha_insercao + 1, col)).value = v_num if v_num is not None else v

        self.somar_linha_no_total_geral(linha_insercao)
        self.garantir_total_geral_ultima_linha()
        self.carregar_datas()

        retorno = {
            "changed": True,
            "message": f"Novo dia {data} criado e periodo {periodo} inserido com sucesso.",
        }
        if self._modelo_info:
            retorno["info"] = self._modelo_info
        return retorno

    # ---------------------------
    # Copiar e colar
    # ---------------------------


    def copiar_colar(self, data, periodo):
        self._modelo_info = None
        data = self.normalizar_data_str(data)

        if self.is_dia_bloqueado(data):
            print(f"⛔ {data} é domingo ou feriado — período não será criado")
            return {
                "changed": False,
                "message": f"{data} e domingo/feriado. Nenhuma alteracao aplicada.",
            }

        if data not in self.datas:
            return self._criar_novo_dia_com_periodo(data, periodo)

        if self.encontrar_linha_periodo(data, periodo):
            print(f"ℹ Período {periodo} já existe em {data} — nada a criar")
            return {
                "changed": False,
                "message": f"Periodo {periodo} ja existe em {data}. Nenhuma alteracao aplicada.",
            }

        # ⚠️ CHAMAR APENAS UMA VEZ
        linha_modelo, data_modelo = self.encontrar_modelo_periodo(data, periodo)

        linha_data = self.encontrar_linha_data(data)
        linha_total_dia = self.encontrar_total_data(linha_data)
        linha_insercao = self.encontrar_linha_insercao_periodo(linha_data, linha_total_dia, periodo)

        valor_data_linha_data = self.ws.range((linha_data, 2)).value

        print(
            f"\n✅ Executando FAZER PONTO no NAVIO - "
            f"Data: {data}, Período: {periodo} "
            f"(modelo: {data_modelo})"
        )

        self.ws.api.Rows(linha_insercao).Insert()

        if linha_modelo >= linha_insercao:
            linha_modelo += 1

        self.ws.api.Rows(linha_modelo).Copy()
        self.ws.api.Rows(linha_insercao).PasteSpecial(-4163)

        self.ws.api.Rows(linha_insercao).Font.Bold = True
        self.ws.range((linha_insercao, 3)).value = periodo

        if linha_insercao == linha_data and valor_data_linha_data:
            self.ws.range((linha_insercao, 2)).value = valor_data_linha_data
            self.ws.range((linha_insercao + 1, 2)).value = None

        linha_nova = linha_insercao
        if linha_insercao <= linha_total_dia:
            linha_total_dia += 1

        self.somar_linha_no_total_do_dia(linha_nova, linha_total_dia)
        self.somar_linha_no_total_geral(linha_nova)

        print("➕ Linha adicionada e somada ao TOTAL DO DIA e TOTAL GERAL")
        retorno = {
            "changed": True,
            "message": f"Periodo {periodo} inserido em {data} com sucesso.",
        }
        if self._modelo_info:
            retorno["info"] = self._modelo_info
        return retorno




    # ---------------------------
    # Soma totais
    # ---------------------------
    def somar_linha_no_total_do_dia(self, linha_origem, linha_total_dia):
        ultima_col = self.ws.range("A1").expand("right").last_cell.column
        for col in range(3, ultima_col + 1):
            v_origem = self.ws.range((linha_origem, col)).value
            v_total = self.ws.range((linha_total_dia, col)).value
            v_origem = self._to_float(v_origem)
            if v_origem is None:
                continue
            v_total = self._to_float(v_total)
            if v_total is None:
                v_total = 0
            self.ws.range((linha_total_dia, col)).value = v_total + v_origem
        if self.debug:
            print(f"➕ Linha {linha_origem} somada ao TOTAL DO DIA")

    def encontrar_linha_total_geral(self):
        ultima_linha = self.ws.cells.last_cell.row
        for i in range(1, ultima_linha + 1):
            valor_a = self.ws.range(f"A{i}").value
            if isinstance(valor_a, str) and self.normalizar_texto(valor_a) == "totalgeral":
                return i
        raise Exception("Total Geral não encontrado.")

    def encontrar_linha_total_geral_opcional(self):
        ultima_linha = self.ws.cells.last_cell.row
        for i in range(1, ultima_linha + 1):
            valor_a = self.ws.range(f"A{i}").value
            if isinstance(valor_a, str) and self.normalizar_texto(valor_a) == "totalgeral":
                return i
        return None

    def somar_linha_no_total_geral(self, linha_origem):
        linha_total_geral = self.encontrar_linha_total_geral_opcional()
        if not linha_total_geral:
            if self.debug:
                print("ℹ Linha 'Total Geral' nao encontrada; soma geral ignorada para este arquivo.")
            return
        ultima_col = self.ws.range("A1").expand("right").last_cell.column
        for col in range(4, ultima_col + 1):
            valor_origem = self.ws.range((linha_origem, col)).value
            valor_origem = self._to_float(valor_origem)
            if valor_origem is None:
                continue
            celula_total = self.ws.range((linha_total_geral, col))
            total_atual = self._to_float(celula_total.value)
            if total_atual is None:
                total_atual = 0
            celula_total.value = total_atual + valor_origem
        if self.debug:
            print(f"➕ Linha {linha_origem} somada ao TOTAL GERAL")

    # ---------------------------
    # Executar
    # ---------------------------

    def executar(self, usar_arquivo_aberto=False, selection=None):
        caminho_final = None
        mtime_antes = None
        data = None
        periodo = None
        resultado = None
        try:
            selection = selection or {}
            data_selecionada = selection.get("data")
            periodo_selecionado = selection.get("periodo")
            caminho_navio = selection.get("caminho_navio")

            if not usar_arquivo_aberto or not self.ws:
                self.abrir_arquivo_navio(caminho=caminho_navio)

            print(f"📄 Arquivo em uso para Fazer Ponto: {self.caminho_navio}")
            if self._caminho_navio_destino:
                print(f"🎯 Destino final (rede): {self._caminho_navio_destino}")

            caminho_final = Path(str(self._caminho_navio_destino or self.caminho_navio)).resolve()
            if caminho_final.exists():
                mtime_antes = caminho_final.stat().st_mtime_ns

            self.carregar_datas()

            data = data_selecionada if data_selecionada else self.escolher_data()
            data = self.normalizar_data_str(data)

            periodo = periodo_selecionado if periodo_selecionado else self.escolher_periodo()
            if periodo not in self.MAPA_PERIODOS.values():
                raise Exception(f"Período inválido: {periodo}")

            resultado = self.copiar_colar(data, periodo) or {
                "changed": False,
                "message": "Nenhuma alteracao aplicada.",
            }

            self.salvar()
            return resultado

        finally:
            if not usar_arquivo_aberto:
                fechar_workbooks(
                    app=self.app,
                    wb_navio=self.wb_navio,
                    wb_cliente=self.wb_cliente
                )

                # Escrita robusta em rede/OneDrive: copia o SaveCopyAs para o arquivo final
                # apenas depois de fechar o workbook (evita lock de arquivo).
                if self._save_copy_path and caminho_final:
                    try:
                        shutil.copy2(self._save_copy_path, caminho_final)
                        time.sleep(0.15)
                        mtime_depois = caminho_final.stat().st_mtime_ns if caminho_final.exists() else None
                        print(f"💾 Write-through aplicado: {caminho_final}")
                        print(f"🕒 mtime antes={mtime_antes} | depois={mtime_depois}")

                        if resultado and bool(resultado.get("changed")):
                            ok = self._periodo_existe_em_data_no_arquivo(caminho_final, data, periodo)
                            if not ok:
                                raise RuntimeError(
                                    "Alteracao nao confirmada no arquivo final apos gravacao. "
                                    "Verifique cache/sincronizacao de rede/OneDrive."
                                )
                            print(f"✅ Verificacao pos-save: periodo {periodo} confirmado em {data}")
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

        # Salva de forma explicita via API do Excel (mais confiavel em caminhos de rede).
        self.wb.api.Save()

        # Tambem gera uma copia fisica para write-through apos fechar o workbook.
        caminho_final = Path(str(self._caminho_navio_destino or self.caminho_navio or "")).resolve()
        copy_name = f"fazer_ponto_write_{caminho_final.stem}_{int(time.time() * 1000)}.xlsx"
        caminho_copy = Path(gettempdir()) / copy_name
        if caminho_copy.exists():
            caminho_copy.unlink()
        self.wb.api.SaveCopyAs(str(caminho_copy))
        self._save_copy_path = str(caminho_copy)

        if not caminho_final.exists():
            raise FileNotFoundError(
                f"Salvou, mas o caminho final nao foi encontrado: {caminho_final}"
            )

        print(f"💾 Arquivo NAVIO salvo com sucesso em: {caminho_final}")

    def _periodo_existe_em_data_no_arquivo(self, caminho_arquivo: Path, data: str, periodo: str) -> bool:
        app = None
        wb = None
        try:
            app = xw.App(visible=False, add_book=False)
            wb = app.books.open(str(caminho_arquivo), read_only=True)
            ws = wb.sheets[0]

            ultima = ws.range("B" + str(ws.cells.last_cell.row)).end("up").row
            linha_data = None
            for i in range(1, ultima + 1):
                v = ws.range(f"B{i}").value
                if isinstance(v, (datetime, date)) and v.strftime("%d/%m/%Y") == data:
                    linha_data = i
                    break
                if isinstance(v, str) and v.strip() == data:
                    linha_data = i
                    break

            if linha_data is None:
                return False

            i = linha_data + 1
            linha_total = None
            while i <= ws.cells.last_cell.row:
                c = ws.range(f"C{i}").value
                if isinstance(c, str) and self.normalizar_texto(c) == "total":
                    linha_total = i
                    break
                i += 1

            if linha_total is None:
                return False

            for j in range(linha_data, linha_total):
                c = ws.range(f"C{j}").value
                if isinstance(c, str):
                    p = self.normalizar_periodo(c)
                    if p == periodo:
                        return True
            return False
        finally:
            try:
                if wb:
                    wb.close()
            except Exception:
                pass
            try:
                if app:
                    app.quit()
            except Exception:
                pass

    def _export_preview_pdf(self):
        if not self.ws:
            return None

        nome_base = Path(str(self.caminho_navio or "navio")).stem
        caminho_pdf = Path(gettempdir()) / f"preview_fazer_ponto_{nome_base}.pdf"
        if caminho_pdf.exists():
            caminho_pdf.unlink()

        try:
            # Preview do Fazer Ponto: paisagem + area util real.
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

    def _detectar_area_util_preview(self, max_linhas_scan=1200, max_colunas_scan=180):
        """
        Detecta a area com conteudo real (valor/formula), evitando colunas fantasmas.
        """
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
        formulas = self.ws.range((1, 1), (scan_rows, scan_cols)).formula

        if not isinstance(valores, list):
            valores = [[valores]]
        elif valores and not isinstance(valores[0], list):
            valores = [valores]

        if not isinstance(formulas, list):
            formulas = [[formulas]]
        elif formulas and not isinstance(formulas[0], list):
            formulas = [formulas]

        ultima_linha = 1
        ultima_coluna = 1
        for i in range(scan_rows):
            linha_vals = valores[i] if i < len(valores) else []
            linha_for = formulas[i] if i < len(formulas) else []
            for j in range(scan_cols):
                v = linha_vals[j] if j < len(linha_vals) else None
                f = linha_for[j] if j < len(linha_for) else None

                tem_valor = v not in (None, "")
                if isinstance(v, str):
                    tem_valor = v.strip() != ""

                # Para preview, formulas que resultam em vazio nao devem inflar PrintArea.
                if tem_valor:
                    ultima_linha = max(ultima_linha, i + 1)
                    ultima_coluna = max(ultima_coluna, j + 1)

        # padding leve para nao cortar totais
        ultima_linha = min(ultima_linha + 2, scan_rows)
        ultima_coluna = min(ultima_coluna + 1, scan_cols)
        return max(ultima_linha, 1), max(ultima_coluna, 1)

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
                "Processo: Fazer Ponto",
                f"Arquivo: {Path(caminho_rede).name}",
                f"Total de datas no arquivo: {len(self.datas)}",
            ]

            if data_selecionada and periodo_selecionado:
                data_selecionada = self.normalizar_data_str(data_selecionada)
                periodo = periodo_selecionado
                if periodo not in self.MAPA_PERIODOS.values():
                    raise Exception(f"Periodo invalido: {periodo}")

                linhas.append(f"Data: {data_selecionada}")
                linhas.append(f"Periodo: {periodo}")

                if self.is_dia_bloqueado(data_selecionada):
                    linhas.append("Status: bloqueado (domingo/feriado)")
                elif data_selecionada in self.datas and self.encontrar_linha_periodo(data_selecionada, periodo):
                    linhas.append("Status: periodo ja existe (nao sera criado)")
                else:
                    _, data_modelo = self.encontrar_modelo_periodo(data_selecionada, periodo)
                    if data_selecionada in self.datas:
                        linhas.append(f"Status: pronto para inserir (modelo base: {data_modelo})")
                    else:
                        linhas.append(f"Status: novo dia sera criado (modelo base: {data_modelo})")
            else:
                linhas.append("Status: selecione data e periodo ao clicar em 'Gerar Excel'.")

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


ProgramaCopiarPeriodo = FazerPonto
