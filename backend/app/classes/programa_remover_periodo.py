from datetime import date, datetime

import xlwings as xw

from yuta_helpers import fechar_workbooks, feriados_br, selecionar_arquivo_navio


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
    # Encontrar período EXATO
    # ---------------------------




    def encontrar_linha_periodo(self, data, periodo):
        linha_data = self.encontrar_linha_data(data)
        if not linha_data:
            return None

        # 🔴 REGRA ESPECIAL PARA 00h
        if periodo == "00h":
            linha_acima = linha_data - 1
            if linha_acima > 0:
                valor_c = self.ws.range(f"C{linha_acima}").value
                if isinstance(valor_c, str):
                    p = self.normalizar_periodo(valor_c)
                    if p == "00h":
                        return linha_acima

        # 🔽 Procura normal abaixo da data
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
    # Subtrações
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
    # Remover período
    # ---------------------------

    def remover_periodo(self, data, periodo):
        if self.is_dia_bloqueado(data):
            print(f"⛔ {data} é domingo ou feriado — nenhuma ação executada")
            return

        linha = self.encontrar_linha_periodo(data, periodo)
        if not linha:
            print(f"ℹ Período {periodo} não existe em {data} — nada a remover")
            return

        linha_data = self.encontrar_linha_data(data)
        linha_total_dia = self.encontrar_total_data(linha_data)

        print(f"\n🗑 Removendo período {periodo} — Data {data}")

        if linha_total_dia:
            self.subtrair_total_dia(linha, linha_total_dia)

        self.subtrair_total_geral(linha)

        self.ws.api.Rows(linha).Delete()

        print("➖ Linha removida e totais ajustados")

    # ---------------------------
    # Execução
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
                raise Exception(f"Data inválida para este arquivo: {data}")

            periodo = periodo_selecionado if periodo_selecionado else self.escolher_periodo()
            if periodo not in self.MAPA_PERIODOS.values():
                raise Exception(f"Período inválido: {periodo}")

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
            print("💾 Arquivo salvo com sucesso")
