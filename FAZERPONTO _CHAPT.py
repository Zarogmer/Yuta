import xlwings as xw
from pathlib import Path
from datetime import datetime, date
import holidays
from datetime import datetime

feriados_br = holidays.Brazil()


class ProgramaCopiarPeriodo:
    PERIODOS_MENU = {
        "1": "06h",
        "2": "12h",
        "3": "18h",
        "4": "00h"
    }

    MAPA_PERIODOS = {
        "06h": "06h",
        "6h": "06h",
        "06": "06h",

        "12h": "12h",
        "12": "12h",

        "18h": "18h",
        "18": "18h",

        "00h": "00h",
        "0h": "00h",
        "00": "00h",
        "24h": "00h"
    }

    EQUIVALENTES = {
        "06h": ["06h", "12h"],
        "12h": ["12h", "06h"],
        "18h": ["18h", "00h"],
        "00h": ["00h", "18h"]
    }

    BLOCOS = {
        "06h": 1,
        "12h": 1,
        "18h": 2,
        "00h": 2
    }


    def __init__(self, ws=None, debug=False):
        self.ws = ws
        self.debug = debug
        self.datas = []

    # ---------------------------
    # Utilitários
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

    # ---------------------------
    # Buscar modelo inteligente
    # ---------------------------
    

    def encontrar_modelo_periodo_inteligente(self, data_destino, periodo):
        datas_ordenadas = sorted(self.datas, key=lambda d: self.parse_data(d))
        idx = datas_ordenadas.index(data_destino)
        bloco_alvo = self.BLOCOS[periodo]

        def procurar_na_data(d, mesmo_dia=False):
            linha_data = self.encontrar_linha_data(d)
            i = linha_data + 1

            while True:
                valor_a = self.ws.range(f"A{i}").value
                valor_c = self.ws.range(f"C{i}").value

                if isinstance(valor_a, str) and self.normalizar_texto(valor_a) == "totalgeral":
                    break

                if not isinstance(valor_c, str):
                    i += 1
                    continue

                texto = self.normalizar_texto(valor_c)
                if texto == "total":
                    break

                p = self.normalizar_periodo(texto)
                if not p:
                    i += 1
                    continue

                # ✅ MESMO DIA → aceita qualquer equivalente
                if mesmo_dia and p in self.EQUIVALENTES[periodo]:
                    if self.debug:
                        print(f"✔ Mesmo dia: usando {p} de {d}")
                    return i

                # ✅ OUTRO DIA → BLOCO OBRIGATÓRIO
                if not mesmo_dia and self.BLOCOS[p] == bloco_alvo:
                    if self.debug:
                        print(f"✔ Outro dia: usando {p} de {d} (bloco {bloco_alvo})")
                    return i

                i += 1

            return None

        # 1️⃣ TENTA NA DATA ESCOLHIDA (SEM RESTRIÇÃO)
        linha = procurar_na_data(data_destino, mesmo_dia=True)
        if linha:
            return linha

        # 2️⃣ OUTRAS DATAS (SEM DOMINGO / FERIADO)
        for offset in range(1, len(datas_ordenadas)):
            for novo_idx in (idx + offset, idx - offset):
                if 0 <= novo_idx < len(datas_ordenadas):
                    d = datas_ordenadas[novo_idx]

                    if self.is_dia_bloqueado(d):
                        if self.debug:
                            print(f"⛔ Pulando data bloqueada: {d}")
                        continue

                    linha = procurar_na_data(d, mesmo_dia=False)
                    if linha:
                        return linha

        raise Exception(
            f"Nenhum modelo válido encontrado para {periodo} "
            f"(busca completa realizada)"
        )


    # ---------------------------
    # Copiar e colar
    # ---------------------------

    
    def copiar_colar(self, data, periodo):
        # 1️⃣ BLOQUEIO DE CALENDÁRIO
        if self.is_dia_bloqueado(data):
            print(f"⛔ {data} é domingo ou feriado — período não será criado")
            return

        linha_data = self.encontrar_linha_data(data)
        linha_total_dia = self.encontrar_total_data(linha_data)
        linha_modelo = self.encontrar_modelo_periodo_inteligente(data, periodo)
        print(data, "bloqueado?", self.is_dia_bloqueado(data))

        self.ws.api.Rows(linha_total_dia).Insert()
        ultima_col = self.ws.range("A1").expand("right").last_cell.column

        origem = self.ws.range((linha_modelo, 1), (linha_modelo, ultima_col))
        destino = self.ws.range((linha_total_dia, 1), (linha_total_dia, ultima_col))

        origem.copy(destino)
        destino.api.Font.Bold = True
        self.ws.range((linha_total_dia, 3)).value = periodo

        self.somar_linha_no_total_do_dia(linha_total_dia, linha_total_dia + 1)
        self.somar_linha_no_total_geral(linha_total_dia)





    # ---------------------------
    # Soma totais
    # ---------------------------
    def somar_linha_no_total_do_dia(self, linha_origem, linha_total_dia):
        ultima_col = self.ws.range("A1").expand("right").last_cell.column
        for col in range(4, ultima_col + 1):
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
    def executar(self):
        app = xw.App(visible=False)
        try:
            caminho = Path.home() / "Desktop" / "1.xlsx"
            wb = app.books.open(str(caminho))
            self.ws = wb.sheets[0]

            self.carregar_datas()
            data = self.escolher_data()
            periodo = self.escolher_periodo()
            
            self.copiar_colar(data, periodo)

            wb.save(Path.home() / "Desktop" / "1_atualizado.xlsx")
            print("✔ Arquivo salvo em Desktop/1_atualizado.xlsx")
        finally:
            app.quit()


if __name__ == "__main__":
    ProgramaCopiarPeriodo(debug=True).executar()
