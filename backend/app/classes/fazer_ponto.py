from datetime import date, datetime

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
    ORDEM_PERIODOS = {"00h": 0, "06h": 1, "12h": 2, "18h": 3}


    # ---------------------------
    # Abrir arquivo NAVIO
    # ---------------------------



    def abrir_arquivo_navio(self, caminho=None):
        caminho = caminho or self.caminho_navio or selecionar_arquivo_navio()
        if not caminho:
            raise FileNotFoundError("Arquivo do NAVIO não selecionado")

        self.caminho_navio = caminho

        self.app = xw.App(visible=False, add_book=False)
        self.wb_navio = self.app.books.open(caminho)
        self.wb = self.wb_navio
        self.ws = self.wb.sheets[0]



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

    def normalizar_texto(self, texto):
        return str(texto).lower().replace(" ", "")

    def normalizar_periodo(self, texto):
        t = self.normalizar_texto(texto)
        return self.MAPA_PERIODOS.get(t, None)




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
        ordem_alvo = self.ORDEM_PERIODOS.get(periodo)
        if ordem_alvo is None:
            return linha_total_dia

        for i in range(linha_data, linha_total_dia):
            valor_c = self.ws.range(f"C{i}").value
            if not isinstance(valor_c, str):
                continue
            p = self.normalizar_periodo(valor_c)
            if not p:
                continue
            if self.ORDEM_PERIODOS.get(p, 999) > ordem_alvo:
                return i

        return linha_total_dia

    # ---------------------------
    # Buscar modelo inteligente
    # ---------------------------

    def encontrar_modelo_periodo(self, data_destino, periodo):
        """
        Retorna: (linha_modelo, data_modelo)
        """

        datas_ordenadas = sorted(self.datas, key=lambda d: self.parse_data(d))
        if data_destino not in datas_ordenadas:
            raise Exception(f"Data base {data_destino} não está na lista de datas válidas")

        idx = datas_ordenadas.index(data_destino)

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
                    if self.debug:
                        print(f"✔ Usando {p} da data {data}")
                    return i, data

                if aceitar_equivalente and p in self.EQUIVALENTES.get(periodo, []):
                    if self.debug:
                        print(f"⚠ Usando equivalente {p} da data {data}")
                    return i, data

                i += 1

        # 1️⃣ Mesmo dia
        resultado = procurar_em_data(data_destino, aceitar_equivalente=True)
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

    # ---------------------------
    # Copiar e colar
    # ---------------------------


    def copiar_colar(self, data, periodo):
        if self.is_dia_bloqueado(data):
            print(f"⛔ {data} é domingo ou feriado — período não será criado")
            return

        if self.encontrar_linha_periodo(data, periodo):
            print(f"ℹ Período {periodo} já existe em {data} — nada a criar")
            return

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




    # ---------------------------
    # Soma totais
    # ---------------------------
    def somar_linha_no_total_do_dia(self, linha_origem, linha_total_dia):
        ultima_col = self.ws.range("A1").expand("right").last_cell.column
        for col in range(3, ultima_col + 1):
            v_origem = self.ws.range((linha_origem, col)).value
            v_total = self.ws.range((linha_total_dia, col)).value
            try:
                v_origem = float(v_origem)
            except:
                continue
            try:
                v_total = float(v_total or 0)
            except:
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

    def somar_linha_no_total_geral(self, linha_origem):
        linha_total_geral = self.encontrar_linha_total_geral()
        ultima_col = self.ws.range("A1").expand("right").last_cell.column
        for col in range(4, ultima_col + 1):
            valor_origem = self.ws.range((linha_origem, col)).value
            if isinstance(valor_origem, (int, float)):
                celula_total = self.ws.range((linha_total_geral, col))
                celula_total.value = (celula_total.value or 0) + valor_origem
        if self.debug:
            print(f"➕ Linha {linha_origem} somada ao TOTAL GERAL")

    # ---------------------------
    # Executar
    # ---------------------------

    def executar(self, usar_arquivo_aberto=False, selection=None):
        try:
            selection = selection or {}
            data_selecionada = selection.get("data")
            periodo_selecionado = selection.get("periodo")
            caminho_navio = selection.get("caminho_navio")

            if not usar_arquivo_aberto or not self.ws:
                self.abrir_arquivo_navio(caminho=caminho_navio)

            self.carregar_datas()

            data = data_selecionada if data_selecionada else self.escolher_data()
            if data not in self.datas:
                raise Exception(f"Data inválida para este arquivo: {data}")

            periodo = periodo_selecionado if periodo_selecionado else self.escolher_periodo()
            if periodo not in self.MAPA_PERIODOS.values():
                raise Exception(f"Período inválido: {periodo}")

            self.copiar_colar(data, periodo)

            self.salvar()

        finally:
            if not usar_arquivo_aberto:
                fechar_workbooks(
                    app=self.app,
                    wb_navio=self.wb_navio,
                    wb_cliente=self.wb_cliente
                )



    def salvar(self):
        if not self.wb:
            raise Exception("Nenhum workbook aberto para salvar")

        self.wb.save()
        print("💾 Arquivo NAVIO salvo com sucesso")


ProgramaCopiarPeriodo = FazerPonto
