import xlwings as xw
from pathlib import Path
from datetime import datetime, date


class ProgramaCopiarPeriodo:
    PERIODOS = {
        "1": "06h",
        "2": "12h",
        "3": "18h",
        "4": "00h"
    }

    def __init__(self, debug=False):
        self.app = None
        self.wb = None
        self.ws = None
        self.datas = None
        self.debug = debug

    # =========================
    # Abrir Excel
    # =========================
    def abrir(self):
        caminho = Path.home() / "Desktop" / "1.xlsx"
        if not caminho.exists():
            raise FileNotFoundError(f"Arquivo não encontrado: {caminho}")

        self.app = xw.App(visible=False)
        self.wb = self.app.books.open(str(caminho))
        self.ws = self.wb.sheets[0]

    # =========================
    # Utilidades
    # =========================
    def parse_data_str(self, data_str):
        return datetime.strptime(data_str, "%d/%m/%Y")

    def normalizar_texto(self, valor):
        return (
            str(valor)
            .strip()
            .lower()
            .replace(" ", "")
            .replace("\n", "")
            .replace("\r", "")
        )

    def normalizar_periodo(self, texto):
        mapa = {
            "06h": "06h", "6h": "06h", "06:00": "06h",
            "12h": "12h", "12:00": "12h",
            "18h": "18h", "18:00": "18h",
            "00h": "00h", "0h": "00h", "00:00": "00h",
            "00h00": "00h", "midnight": "00h"
        }
        return mapa.get(texto)

    # =========================
    # Datas
    # =========================
    def obter_datas(self):
        datas = []
        ultima_linha = self.ws.range("B" + str(self.ws.cells.last_cell.row)).end("up").row

        for i in range(1, ultima_linha + 1):
            valor = self.ws.range(f"B{i}").value
            if isinstance(valor, (datetime, date)):
                datas.append(valor.strftime("%d/%m/%Y"))
            elif isinstance(valor, str) and "/" in valor:
                datas.append(valor.strip())

        datas_unicas = list(dict.fromkeys(datas))
        if not datas_unicas:
            raise Exception("Nenhuma data encontrada na coluna B.")

        return datas_unicas

    # =========================
    # Períodos da data
    # =========================
    def obter_periodos_da_data(self, linha_data):
        periodos = []
        i = linha_data + 1

        while True:
            valor = self.ws.range(f"C{i}").value
            if valor is None:
                i += 1
                continue

            texto = self.normalizar_texto(valor)

            if texto == "total":
                break

            periodo = self.normalizar_periodo(texto)
            if periodo:
                periodos.append(periodo)

            i += 1

        return periodos

    # =========================
    # Escolhas
    # =========================
    def escolher_data(self, datas):
        print("\nDatas disponíveis:")
        for i, d in enumerate(datas, 1):
            print(f"{i} - {d}")

        while True:
            try:
                op = int(input("Escolha a data: "))
                return datas[op - 1]
            except:
                print("Opção inválida.")

    def escolher_periodo(self):
        print("\nHorário:")
        print("1 = 06h | 2 = 12h | 3 = 18h | 4 = 00h")
        while True:
            op = input("Opção: ").strip()
            if op in self.PERIODOS:
                return self.PERIODOS[op]
            

    def is_domingo(self, data_str):
        dt = self.parse_data_str(data_str)
        return dt.weekday() == 6  # 6 = domingo


    # =========================
    # Localização
    # =========================
    def encontrar_linha_data(self, data_str):
        ultima = self.ws.range("B" + str(self.ws.cells.last_cell.row)).end("up").row
        for i in range(1, ultima + 1):
            valor = self.ws.range(f"B{i}").value
            if isinstance(valor, (datetime, date)):
                if valor.strftime("%d/%m/%Y") == data_str:
                    return i
            elif valor == data_str:
                return i
        raise Exception("Data não encontrada.")



    def encontrar_total_data(self, linha_data):
        i = linha_data + 1

        while True:
            valor_c = self.ws.range(f"C{i}").value
            valor_a = self.ws.range(f"A{i}").value

            # ⛔ não pode passar do Total Geral
            if isinstance(valor_a, str) and self.normalizar_texto(valor_a) == "totalgeral":
                raise Exception("❌ 'Total' do dia não encontrado antes do Total Geral")

            if isinstance(valor_c, str):
                if self.normalizar_texto(valor_c) == "total":
                    return i

            # ⛔ segurança absoluta (fim da planilha)
            if i > self.ws.cells.last_cell.row:
                raise Exception("❌ Fim da planilha sem encontrar 'Total' do dia")

            i += 1


    def encontrar_linha_modelo_na_data(self, linha_data, periodo):
        i = linha_data + 1

        while True:
            valor_c = self.ws.range(f"C{i}").value
            valor_a = self.ws.range(f"A{i}").value

            # ⛔ segurança: não passar do Total Geral
            if isinstance(valor_a, str) and valor_a.strip().lower() == "total geral":
                break

            if valor_c is None:
                i += 1
                continue

            texto = self.normalizar_texto(valor_c)

            # ⛔ fim do bloco do dia
            if texto == "total":
                break

            # 🎯 período encontrado
            if self.normalizar_periodo(texto) == periodo:
                return i

            i += 1

        raise Exception(f"Período '{periodo}' não encontrado nessa data.")

        

    def resolver_linha_periodo(self, data, periodo):
        linha_data = self.encontrar_linha_data(data)

        # =========================
        # 1️⃣ MESMO DIA
        # =========================
        i = linha_data + 1
        while True:
            valor = self.ws.range(f"C{i}").value
            if not isinstance(valor, str):
                i += 1
                continue

            texto = self.normalizar_texto(valor)
            if texto == "total":
                break

            if self.normalizar_periodo(texto) == periodo:
                return i

            i += 1

        # =========================
        # 2️⃣ DIA SEGUINTE (PRIORIDADE PARA 06h)
        # =========================
        datas_ordenadas = sorted(self.datas, key=self.parse_data_str)
        idx = datas_ordenadas.index(data)

        if idx < len(datas_ordenadas) - 1:
            proxima_data = datas_ordenadas[idx + 1]
            linha_d = self.encontrar_linha_data(proxima_data)

            i = linha_d + 1
            while True:
                valor = self.ws.range(f"C{i}").value
                if not isinstance(valor, str):
                    i += 1
                    continue

                texto = self.normalizar_texto(valor)
                if texto == "total":
                    break

                if self.normalizar_periodo(texto) == periodo:
                    return i

                i += 1

        # =========================
        # 3️⃣ FALLBACK CONTROLADO (06h ← 12h)
        # =========================
        if periodo == "06h":
            i = linha_data + 1
            while True:
                valor = self.ws.range(f"C{i}").value
                if not isinstance(valor, str):
                    i += 1
                    continue

                texto = self.normalizar_texto(valor)
                if texto == "total":
                    break

                if self.normalizar_periodo(texto) == "12h":
                    return i

                i += 1

        raise Exception(f"Nenhum modelo válido encontrado para {periodo}")

    

    def somar_linha_no_total_do_dia(self, linha_origem, linha_total_dia):
        """
        Soma a linha colada no TOTAL do próprio dia
        (ignora colunas A, B, C)
        """
        ultima_col = self.ws.range("A1").expand("right").last_cell.column

        for col in range(4, ultima_col + 1):  # D em diante
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
            print(f"➕ Linha {linha_origem} somada ao Total Geral")

    def encontrar_linha_total_do_dia(self, linha_inicio):
        """
        Desce a partir da linha_inicio até encontrar:
        - 'Total' na coluna C  → retorna
        - 'Total Geral' na coluna A → para (não atravessa)
        """

        ultima_linha = self.ws.cells.last_cell.row

        for i in range(linha_inicio + 1, ultima_linha + 1):
            valor_c = self.ws.range(f"C{i}").value
            valor_a = self.ws.range(f"A{i}").value

            # ⛔ não atravessar Total Geral
            if isinstance(valor_a, str) and self.normalizar_texto(valor_a) == "totalgeral":
                break

            if not isinstance(valor_c, str):
                continue

            if self.normalizar_texto(valor_c) == "total":
                return i

        raise Exception("Total do dia não encontrado.")
    


    def encontrar_modelo_periodo_inteligente(self, data_destino, periodo):
        """
        Regras:
        - Primeiro tenta o mesmo dia
        - Depois tenta equivalentes no mesmo dia
        - Depois busca em outras datas (não domingo)
        """

        # equivalências permitidas
        equivalentes = {
            "06h": ["06h", "12h"],
            "12h": ["12h", "06h"],
            "18h": ["18h", "00h"],
            "00h": ["00h", "18h"],
        }

        eh_domingo_destino = self.is_domingo(data_destino)

        datas_busca = self.datas.copy()

        # prioridade: mesma data primeiro
        datas_busca.remove(data_destino)
        datas_busca.insert(0, data_destino)

        for d in datas_busca:
            if self.is_domingo(d) != eh_domingo_destino:
                continue

            linha_data = self.encontrar_linha_data(d)
            i = linha_data + 1

            while True:
                valor_c = self.ws.range(f"C{i}").value
                valor_a = self.ws.range(f"A{i}").value

                if isinstance(valor_a, str) and self.normalizar_texto(valor_a) == "totalgeral":
                    break

                if not isinstance(valor_c, str):
                    i += 1
                    continue

                texto = self.normalizar_texto(valor_c)

                if texto == "total":
                    break

                p = self.normalizar_periodo(texto)

                if p in equivalentes.get(periodo, []):
                    if self.debug:
                        print(f"✔ Modelo {p} usado de {d} (linha {i})")
                    return i

                i += 1

        raise Exception(f"Nenhum modelo válido encontrado para {periodo}")



    # =========================
    # Copiar / Colar
    # =========================



    def copiar_colar(self, data, periodo):
        linha_data = self.encontrar_linha_data(data)
        linha_total_dia = self.encontrar_total_data(linha_data)

        linha_modelo = self.encontrar_modelo_periodo_inteligente(data, periodo)

        # ➕ Insere linha antes do TOTAL do dia
        self.ws.api.Rows(linha_total_dia).Insert()

        ultima_col = self.ws.range("A1").expand("right").last_cell.column

        origem = self.ws.range((linha_modelo, 1), (linha_modelo, ultima_col))
        destino = self.ws.range((linha_total_dia, 1), (linha_total_dia, ultima_col))

        origem.copy(destino)

        # ✅ NEGRITO NA LINHA ADICIONADA
        destino.api.Font.Bold = True

        # ✏️ Ajusta período (coluna C)
        self.ws.range((linha_total_dia, 3)).value = periodo

        # ➕ somas
        self.somar_linha_no_total_do_dia(
            linha_origem=linha_total_dia,
            linha_total_dia=linha_total_dia + 1
        )

        self.somar_linha_no_total_geral(linha_total_dia)



    # =========================
    # Execução
    # =========================
    def executar(self):
        try:
            self.abrir()
            self.datas = self.obter_datas()
            data = self.escolher_data(self.datas)
            periodo = self.escolher_periodo()
            self.copiar_colar(data, periodo)

            novo = Path.home() / "Desktop" / "1_atualizado.xlsx"
            self.wb.save(novo)
            print(f"\n✔ Arquivo salvo em {novo}")

        finally:
            if self.app:
                self.app.quit()


if __name__ == "__main__":
    ProgramaCopiarPeriodo().executar()
