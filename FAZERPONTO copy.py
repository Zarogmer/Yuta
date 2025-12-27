import xlwings as xw
from pathlib import Path
from datetime import datetime, date


class ProgramaCopiarPeriodo:
    PERIODOS_MENU = {
        "1": "06h",
        "2": "12h",
        "3": "18h",
        "4": "00h"
    }

    MAPA_PERIODOS = {
        "06h": "06h", "6h": "06h", "06:00": "06h",
        "12h": "12h", "12:00": "12h",
        "18h": "18h", "18:00": "18h",
        "00h": "00h", "0h": "00h", "00:00": "00h",
    }

    EQUIVALENTES = {
        "06h": ["06h", "12h"],
        "12h": ["12h", "06h"],
        "18h": ["18h", "00h"],
        "00h": ["00h", "18h"],
    }

    def __init__(self, debug=False):
        self.debug = debug
        self.app = None
        self.wb = None
        self.ws = None
        self.datas = []

    # ======================================================
    # Excel
    # ======================================================
    def abrir_excel(self):
        caminho = Path.home() / "Desktop" / "1.xlsx"
        if not caminho.exists():
            raise FileNotFoundError(caminho)

        self.app = xw.App(visible=False)
        self.wb = self.app.books.open(str(caminho))
        self.ws = self.wb.sheets[0]


    def salvar(self):
        novo = Path.home() / "Desktop" / "1_atualizado.xlsx"
        self.wb.save(novo)
        print(f"\n✔ Arquivo salvo em {novo}")



    # ======================================================
    # Utilidades
    # ======================================================
    def normalizar_texto(self, v):
        return str(v).strip().lower().replace(" ", "")

    def normalizar_periodo(self, v):
        return self.MAPA_PERIODOS.get(self.normalizar_texto(v))

    def parse_data(self, d):
        return datetime.strptime(d, "%d/%m/%Y")

    def is_domingo(self, d):
        return self.parse_data(d).weekday() == 6

    # ======================================================
    # Datas
    # ======================================================
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
            raise Exception("Nenhuma data encontrada")

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

    # ======================================================
    # Localização
    # ======================================================

    def linha_total_dia(self, linha_data):
        """
        Procura o 'Total' do dia na coluna C,
        a partir da linha da data, sem varrer a planilha inteira
        """
        ultima = self.ultima_linha_real()

        for i in range(linha_data + 1, ultima + 1):
            valor_c = self.ws.range(f"C{i}").value
            valor_a = self.ws.range(f"A{i}").value

            # se chegou no Total Geral antes → erro estrutural
            if isinstance(valor_a, str) and self.normalizar_texto(valor_a) == "totalgeral":
                break

            if isinstance(valor_c, str) and self.normalizar_texto(valor_c) == "total":
                return i

        raise Exception("Total do dia não encontrado.")





    def linha_da_data(self, data_str):
        ultima = self.ultima_linha_real()

        for i in range(1, ultima + 1):
            valor = self.ws.range(f"B{i}").value

            if isinstance(valor, (datetime, date)):
                if valor.strftime("%d/%m/%Y") == data_str:
                    return i

            elif isinstance(valor, str) and valor.strip() == data_str:
                return i

        raise Exception(f"Data {data_str} não encontrada.")


    # ======================================================
    # Modelo inteligente
    # ======================================================
    def encontrar_modelo(self, data_destino, periodo):
        eh_domingo = self.is_domingo(data_destino)

        datas_busca = [data_destino] + [
            d for d in self.datas if d != data_destino
        ]

        for d in datas_busca:
            if self.is_domingo(d) != eh_domingo:
                continue

            linha = self.linha_da_data(d)
            i = linha + 1

            while True:
                texto = self.ws.range(f"C{i}").value
                if self.normalizar_texto(texto or "") == "total":
                    break

                p = self.normalizar_periodo(texto)
                if p in self.EQUIVALENTES.get(periodo, []):
                    if self.debug:
                        print(f"✔ Modelo {p} usado de {d} (linha {i})")
                    return i
                i += 1

        raise Exception(f"Nenhum modelo válido para {periodo}")

    # ======================================================
    # Somas
    # ======================================================
    def somar_total_dia(self, linha_origem, linha_total):
        ultima_col = self.ws.range("A1").expand("right").last_cell.column

        for col in range(4, ultima_col + 1):
            v1 = self.ws.range((linha_origem, col)).value
            v2 = self.ws.range((linha_total, col)).value

            try:
                v1 = float(v1 or 0)
                v2 = float(v2 or 0)
            except:
                continue

            self.ws.range((linha_total, col)).value = v1 + v2


    def somar_total_geral(self, linha_origem):
        ultima = self.ws.cells.last_cell.row
        for i in range(1, ultima + 1):
            if self.normalizar_texto(self.ws.range(f"A{i}").value or "") == "totalgeral":
                total_geral = i
                break
        else:
            raise Exception("Total Geral não encontrado")

        ultima_col = self.ws.range("A1").expand("right").last_cell.column
        for c in range(4, ultima_col + 1):
            try:
                self.ws.range((total_geral, c)).value = (
                    float(self.ws.range((total_geral, c)).value or 0)
                    + float(self.ws.range((linha_origem, c)).value or 0)
                )
            except:
                pass

    def somar_linha_com_linha_abaixo(self, linha_origem):
        """
        Soma a linha_origem com a linha imediatamente abaixo dela
        (ignora colunas A, B, C)
        """
        linha_destino = linha_origem + 1
        ultima_col = self.ws.range("A1").expand("right").last_cell.column

        for col in range(4, ultima_col + 1):  # D em diante
            v_origem = self.ws.range((linha_origem, col)).value
            v_destino = self.ws.range((linha_destino, col)).value

            try:
                v_origem = float(v_origem or 0)
                v_destino = float(v_destino or 0)
            except:
                continue

            self.ws.range((linha_destino, col)).value = v_destino + v_origem

        if self.debug:
            print(f"➕ Linha {linha_origem} somada com a linha {linha_destino}")


    # ======================================================
    # Copiar / Colar
    # ======================================================

    def encontrar_ultimas_linhas_soma(self):
        """
        Retorna:
        - penultima_linha
        - ultima_linha
        Considera linhas após o Total Geral, sem data (col B) e sem período (col C)
        """

        ultima = self.ws.cells.last_cell.row
        linhas_validas = []

        for i in range(ultima, 1, -1):
            valor_a = self.ws.range(f"A{i}").value
            valor_b = self.ws.range(f"B{i}").value
            valor_c = self.ws.range(f"C{i}").value

            # ignora linhas totalmente vazias
            if valor_a is None and valor_b is None and valor_c is None:
                continue

            # linha especial: sem data e sem período
            if valor_b is None and valor_c is None:
                linhas_validas.append(i)

            if len(linhas_validas) == 2:
                return linhas_validas[1], linhas_validas[0]

        raise Exception("Não foi possível identificar as duas últimas linhas de soma.")
    
    def somar_em_linha(self, linha_origem, linha_destino):
        ultima_col = self.ws.range("A1").expand("right").last_cell.column

        for col in range(4, ultima_col + 1):
            v_origem = self.ws.range((linha_origem, col)).value
            v_destino = self.ws.range((linha_destino, col)).value

            try:
                v_origem = float(v_origem)
            except:
                continue

            try:
                v_destino = float(v_destino or 0)
            except:
                v_destino = 0

            self.ws.range((linha_destino, col)).value = v_destino + v_origem


    def ultima_linha_real(self):
        """
        Retorna a última linha com dado real na planilha
        (baseada na coluna A)
        """
        return self.ws.range(
            "A" + str(self.ws.cells.last_cell.row)
        ).end("up").row
    

    def somar_com_linha_abaixo(self, linha_origem):
        """
        Soma a linha_origem com a linha imediatamente abaixo
        (ignora colunas A, B, C)
        """
        ultima_col = self.ws.range("A1").expand("right").last_cell.column
        linha_destino = linha_origem + 1

        for col in range(4, ultima_col + 1):
            v1 = self.ws.range((linha_origem, col)).value
            v2 = self.ws.range((linha_destino, col)).value

            try:
                v1 = float(v1)
                v2 = float(v2 or 0)
            except:
                continue

            self.ws.range((linha_destino, col)).value = v1 + v2

    def encontrar_ultimas_linhas(self):
        """
        Retorna:
        (linha_penultima, linha_ultima)

        Onde:
        - linha_ultima = Total Geral
        - linha_penultima = linha válida imediatamente acima
        """
        ultima = self.ultima_linha_real()

        linha_total_geral = None

        for i in range(1, ultima + 1):
            valor_a = self.ws.range(f"A{i}").value
            if isinstance(valor_a, str) and self.normalizar_texto(valor_a) == "totalgeral":
                linha_total_geral = i
                break

        if not linha_total_geral:
            raise Exception("Total Geral não encontrado.")

        # procura penúltima linha válida acima
        i = linha_total_geral - 1
        while i > 0:
            if self.ws.range(f"A{i}").value not in (None, ""):
                return i, linha_total_geral
            i -= 1

        raise Exception("Penúltima linha não encontrada.")


    def copiar_colar(self, data, periodo):
        linha_data = self.linha_da_data(data)
        linha_total_original = self.linha_total_dia(linha_data)
        linha_modelo = self.encontrar_modelo(data, periodo)

        self.ws.api.Rows(linha_total_original).Insert()

        linha_nova = linha_total_original
        linha_total_dia = linha_total_original + 1

        ultima_col = self.ws.range("A1").expand("right").last_cell.column

        origem = self.ws.range((linha_modelo, 1), (linha_modelo, ultima_col))
        destino = self.ws.range((linha_nova, 1), (linha_nova, ultima_col))

        origem.copy(destino)
        destino.api.Font.Bold = True
        self.ws.range((linha_nova, 3)).value = periodo

        self.somar_total_dia(linha_nova, linha_total_dia)




    # ======================================================
    # Execução
    # ======================================================


    def executar(self):
        try:
            self.abrir_excel()
            self.carregar_datas()
            data = self.escolher_data()
            periodo = self.escolher_periodo()
            self.copiar_colar(data, periodo)
            self.salvar()
        finally:
            if self.app:
                self.app.quit()


if __name__ == "__main__":
    ProgramaCopiarPeriodo(debug=False).executar()
