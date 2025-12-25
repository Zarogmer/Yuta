# ==============================
# IMPORTS LIMPOS (sem duplicados)
# ==============================
import sys
import time
import re
import shutil
import tempfile
import ssl
import certifi
import urllib.request
from datetime import datetime, date, timedelta, timezone
from pathlib import Path
from tkinter import Tk, filedialog
import tkinter as tk
import xlwings as xw
import pandas as pd
import os
from itertools import cycle
import msvcrt
import os


# ==============================
# FUNÇÕES AUXILIARES GLOBAIS (compartilhadas pelas duas classes)
# ==============================

def obter_pasta_faturamentos() -> Path:
    print("\n=== BUSCANDO PASTA FATURAMENTOS AUTOMATICAMENTE ===")

    possiveis_bases = [
        Path(r"C:\Users\Carol\SANPORT LOGÍSTICA PORTUÁRIA LTDA"),
        Path(r"C:\Users\Carol\OneDrive - SANPORT LOGÍSTICA PORTUÁRIA LTDA"),
        Path.home() / "SANPORT LOGÍSTICA PORTUÁRIA LTDA",
        Path.home() / "OneDrive" / "SANPORT LOGÍSTICA PORTUÁRIA LTDA",
    ]

    caminho_alvo = None
    for base in possiveis_bases:
        if base.exists():
            print(f"✅ Encontrada pasta base: {base}")
            candidatos = list(base.rglob("FATURAMENTOS"))
            for candidato in candidatos:
                if "01. FATURAMENTOS" in candidato.parent.as_posix():
                    caminho_alvo = candidato
                    print(f"✅ Pasta FATURAMENTOS encontrada:\n   {caminho_alvo}")
                    break
            if caminho_alvo:
                break
        else:
            print(f"❌ Não encontrada: {base}")

    if not caminho_alvo:
        print("❌ Pasta FATURAMENTOS não localizada automaticamente.")
        raise FileNotFoundError("Pasta FATURAMENTOS não encontrada")

    # Lista arquivos para debug
    arquivos = sorted(caminho_alvo.glob("*.xlsx"))[:10]
    print(f"\nArquivos .xlsx encontrados ({len(list(caminho_alvo.glob('*.xlsx')))}):")
    for arq in arquivos:
        print(f"   • {arq.name}")
    if len(list(caminho_alvo.glob("*.xlsx"))) > 10:
        print("   ... (mais arquivos)")
    print("========================================\n")
    return caminho_alvo

def _obter_proxima_nf(self, pasta_nfs):

        if not os.path.exists(pasta_nfs):
            return 1
        numeros = [int(re.match(r"(\d+)", f).group(1)) for f in os.listdir(pasta_nfs) if re.match(r"(\d+)", f) and f.lower().endswith(".pdf")]
        return max(numeros) + 1 if numeros else 1



def obter_dn_da_pasta(pasta: Path) -> str:
    numeros = re.findall(r"\d+", pasta.name)
    if not numeros:
        raise ValueError("Não foi possível identificar o DN no nome da pasta.")
    return numeros[0]


def obter_nome_navio_da_pasta(pasta: Path) -> str:
    nome_limpo = re.sub(r"^\d+[\s\-_]*", "", pasta.name, flags=re.IGNORECASE).strip()
    return nome_limpo if nome_limpo else "NAVIO NÃO IDENTIFICADO"


def abrir_workbooks(pasta_faturamentos: Path):
    root = tk.Tk()
    root.withdraw()

    pasta_navio_str = filedialog.askdirectory(title="Selecione a pasta do NAVIO (onde está o 1.xlsx)")
    if not pasta_navio_str:
        print("❌ Seleção cancelada pelo usuário.")
        root.destroy()
        return None

    pasta_navio = Path(pasta_navio_str)
    pasta_cliente = pasta_navio.parent
    nome_cliente = pasta_cliente.name.strip()

    arquivos_1 = list(pasta_navio.glob("1*.xls*"))
    if not arquivos_1:
        raise FileNotFoundError(f"Nenhum arquivo iniciando com '1' encontrado em:\n{pasta_navio}")

    arquivo1 = arquivos_1[0]

    arquivo2 = pasta_faturamentos / f"{nome_cliente}.xlsx"
    if not arquivo2.exists():
        raise FileNotFoundError(f"Arquivo do cliente não encontrado:\n{arquivo2}")

    # Abre o Excel
    app = xw.App(visible=False)
    app.api.Calculate()
    time.sleep(0.5)

    try:
        wb1 = app.books.open(str(arquivo1))
        wb2 = app.books.open(str(arquivo2))

        ws1 = wb1.sheets[0]
        nomes_abas = [s.name for s in wb2.sheets]

        if nome_cliente in nomes_abas:
            ws_front = wb2.sheets[nome_cliente]
        elif "FRONT VIGIA" in nomes_abas:
            ws_front = wb2.sheets["FRONT VIGIA"]
        else:
            raise RuntimeError(f"Aba não encontrada. Esperado: '{nome_cliente}' ou 'FRONT VIGIA'")

        root.destroy()
        return app, wb1, wb2, ws1, ws_front

    except Exception as e:
        try:
            if 'wb1' in locals(): wb1.close()
            if 'wb2' in locals(): wb2.close()
            app.quit()
        except:
            pass
        root.destroy()
        raise e


def copiar_para_temp_e_ler_excel(caminho_original: Path | str) -> pd.DataFrame:
    caminho_original = Path(caminho_original)
    if not caminho_original.exists():
        raise FileNotFoundError(f"Arquivo não encontrado: {caminho_original}")

    with tempfile.TemporaryDirectory() as temp_dir:
        caminho_temp = Path(temp_dir) / caminho_original.name
        print(f"Copiando {caminho_original.name} para temporário...")
        shutil.copy2(caminho_original, caminho_temp)
        print(f"Lendo {caminho_temp} com pandas...")
        return pd.read_excel(caminho_temp, engine="openpyxl")


def fechar_workbooks(app, wb1=None, wb2=None, arquivo_saida=None):
    try:
        if wb1:
            wb1.save()
            wb1.close()
        if wb2:
            if not arquivo_saida:
                raise RuntimeError("arquivo_saida obrigatório para wb2")
            wb2.save(str(arquivo_saida))
            wb2.close()
    finally:
        if app:
            app.quit()


# ==============================
# LICENÇA E DATA
# ==============================

def data_online():
    context = ssl.create_default_context(cafile=certifi.where())
    req = urllib.request.Request("https://www.cloudflare.com", headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req, context=context, timeout=5) as r:
        data_str = r.headers["Date"]
    dt_utc = datetime.strptime(data_str, "%a, %d %b %Y %H:%M:%S %Z").replace(tzinfo=timezone.utc)
    dt_local = dt_utc.astimezone()
    return dt_utc, dt_local


def validar_licenca():
    hoje_utc, hoje_local = data_online()
    limite = datetime(hoje_utc.year, hoje_utc.month, 30, tzinfo=timezone.utc)
    if hoje_utc > limite:
        sys.exit("⛔ Licença expirada")
    print(f"📅 Data local: {hoje_local.date()}")


# ==============================
# CLASSE 1: FATURAMENTO COMPLETO
# ==============================

class FaturamentoCompleto:

    def executar(self):
        print("🚀 Iniciando Faturamento Completo...")
        validar_licenca()

        pasta_faturamentos = obter_pasta_faturamentos()
        resultado = abrir_workbooks(pasta_faturamentos)

        if not resultado:
            return

        app, wb1, wb2, ws1, ws_front = resultado

        try:
            pasta_navio = Path(wb1.fullname).parent
            dn = self._obter_dn_da_pasta(pasta_navio)
            nome_navio = self._obter_nome_navio_da_pasta(pasta_navio)

            print(f"📋 DN: {dn}")
            print(f"🚢 Navio: {nome_navio}")

            # FRONT VIGIA
            data_inicio, data_fim = self._processar_front(ws1, ws_front)

            # REPORT VIGIA
            ws_resumo = wb1.sheets["Resumo"]
            periodos = self._obter_periodos(ws_resumo)
            ws_report = wb2.sheets["REPORT VIGIA"]

            self._inserir_linhas_report(ws_report, linha_inicial=22, periodos=periodos)

            ciclos_linha = self._gerar_coluna_E_ajustada(ws1, periodos)
            self._preencher_coluna_E_por_ciclos(ws_report, ciclos_linha)

            valores_por_ciclo = self._mapear_valores_por_ciclo(ws1)
            self._preencher_coluna_G_por_ciclo(ws_report, ciclos_linha, valores_por_ciclo)

            self._montar_datas_report_vigia(ws_report, ws_resumo, periodos=periodos)

            # Financeiro
            self._MMO(wb1, wb2)
            self._escrever_nf(wb2, nome_navio, dn)
            self._OC(str(wb1.fullname), wb2)
            self._credit_note(wb2, f"DN: {dn}/{datetime.now().year}")
            self._quitacao(wb2, f"DN: {dn}/{datetime.now().year}")

            self._arredondar_para_baixo_50(ws_front)
            self._cargonave(ws_front)  # só chama se precisar

            # Salva final
            pasta_saida = Path(wb1.fullname).parent
            arquivo_saida = pasta_saida / "3.xlsx"
            fechar_workbooks(app, wb1, wb2, arquivo_saida)

            print("✅ Faturamento Completo concluído com sucesso!")

        except Exception as e:
            print(f"❌ Erro: {e}")
        finally:
            try:
                app.quit()
            except:
                pass

    # ==============================
    # MÉTODOS PRIVADOS DA CLASSE
    # ==============================

    def _obter_dn_da_pasta(self, pasta: Path) -> str:
        numeros = re.findall(r"\d+", pasta.name)
        if not numeros:
            raise ValueError("DN não identificada")
        return numeros[0]

    def _obter_nome_navio_da_pasta(self, pasta: Path) -> str:
        nome_limpo = re.sub(r"^\d+[\s\-_]*", "", pasta.name, flags=re.IGNORECASE).strip()
        return nome_limpo if nome_limpo else "NAVIO NÃO IDENTIFICADO"

    def _data_por_extenso(self, valor):
        if isinstance(valor, (datetime, date)):
            data = valor if isinstance(valor, datetime) else datetime(valor.year, valor.month, valor.day)
        elif isinstance(valor, str):
            try:
                data = datetime.strptime(valor, "%d/%m/%Y")
            except:
                return ""
        else:
            return ""
        return data.strftime("%d de %B de %Y")

    def _obter_datas_extremos(self, ws_resumo):
        last_row = ws_resumo.used_range.last_cell.row
        valores = ws_resumo.range(f"B1:B{last_row}").value

        datas = []
        hoje = date.today()

        for v in valores:
            if v in (None, "", "Total"):
                continue
            if isinstance(v, datetime):
                d = v.date()
                if d == hoje:
                    continue
                datas.append(d)
            elif isinstance(v, str):
                v = v.strip()
                try:
                    datas.append(datetime.strptime(v, "%d/%m/%Y").date())
                except:
                    pass  # tenta outros formatos se precisar

        if not datas:
            return None, None
        return min(datas), max(datas)

    def _processar_front(self, ws1, ws_front):
        meses = ["", "janeiro", "fevereiro", "março", "abril", "maio", "junho",
                 "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"]

        hoje = datetime.now()
        ws_front.range("C39").value = f"Santos, {hoje.day} de {meses[hoje.month]} de {hoje.year}"

        data_min, data_max = self._obter_datas_extremos(ws1)

        if data_min:
            ws_front.range("D16").value = self._data_por_extenso(data_min)
        if data_max:
            ws_front.range("D17").value = self._data_por_extenso(data_max)

        return data_min, data_max

    def _inserir_linhas_report(self, ws_report, linha_inicial, periodos):
        if periodos <= 1:
            return
        row_height = ws_report.api.Rows(linha_inicial).RowHeight
        for i in range(periodos - 1):
            destino = linha_inicial + 1 + i
            ws_report.api.Rows(destino).Insert()
            ws_report.api.Rows(linha_inicial).Copy(ws_report.api.Rows(destino))
            ws_report.api.Rows(destino).RowHeight = row_height

    def _obter_periodos(self, ws_resumo):
        valores = ws_resumo.range("AA:AA").value
        valores = [v for v in valores if v is not None]
        try:
            ultimo = str(valores[-1]).replace("R$", "").replace(",", ".").strip()
            return int(float(ultimo))
        except:
            return 1

    def _gerar_coluna_E_ajustada(self, ws1, periodos, coluna_horario="C"):
        horario_para_ciclo = {"06h": "06x12", "06H": "06x12", "12h": "12x18", "12H": "12x18",
                              "18h": "18x24", "18H": "18x24", "00h": "00x06", "00H": "00x06"}
        sequencia_padrao = ["06x12", "12x18", "18x24", "00x06"]

        segundo_valor = str(ws1.range(f"{coluna_horario}3").value or "").strip().upper()

        if segundo_valor == "TOTAL" or segundo_valor not in horario_para_ciclo:
            primeiro_ciclo = "00x06"
        else:
            primeiro_ciclo = horario_para_ciclo[segundo_valor]

        idx = sequencia_padrao.index(primeiro_ciclo)
        sequencia = sequencia_padrao[idx:] + sequencia_padrao[:idx]

        return list(cycle(sequencia))[:periodos]

    def _preencher_coluna_E_por_ciclos(self, ws_report, ciclos_linha, linha_inicial=22):
        for i, ciclo in enumerate(ciclos_linha):
            ws_report.range(f"E{linha_inicial + i}").value = ciclo

    def _mapear_valores_por_ciclo(self, ws1, coluna_horario="C", coluna_valor="Z"):
        horario_para_ciclo = {"06h": "06x12", "12h": "12x18", "18h": "18x24", "00h": "00x06"}
        sequencia_ciclos = ["06x12", "12x18", "18x24", "00x06"]

        last_row = ws1.used_range.last_cell.row
        horarios = ws1.range(f"{coluna_horario}1:{coluna_horario}{last_row}").value
        valores = ws1.range(f"{coluna_valor}1:{coluna_valor}{last_row}").value

        horarios = [str(h).strip().upper() if h is not None else "" for h in horarios]

        valores_por_ciclo = {c: [] for c in sequencia_ciclos}

        for h, v in zip(horarios, valores):
            if h in horario_para_ciclo:
                ciclo = horario_para_ciclo[h]
                valores_por_ciclo[ciclo].append(v)

        return valores_por_ciclo

    def _preencher_coluna_G_por_ciclo(self, ws_report, ciclos_linha, valores_por_ciclo, coluna="G", linha_inicial=22):
        indices = {c: 0 for c in valores_por_ciclo}
        for i, ciclo in enumerate(ciclos_linha):
            linha = linha_inicial + i
            lista = valores_por_ciclo.get(ciclo, [])
            idx = indices[ciclo]
            valor = lista[idx] if idx < len(lista) else None
            indices[ciclo] += 1

            cel = ws_report.range(f"{coluna}{linha}")
            cel.value = valor
            try:
                cel.api.NumberFormat = "R$ #.##0,00"
                cel.api.HorizontalAlignment = -4152  # right
                cel.api.Font.Size = 18
            except:
                pass

    def _montar_datas_report_vigia(self, ws_report, ws_resumo, linha_inicial=22, periodos=None):
        if periodos is None:
            raise ValueError("periodos obrigatório")

        data_inicio, _ = self._obter_datas_extremos(ws_resumo)
        if not data_inicio:
            raise ValueError("Data início não encontrada")

        data_atual = data_inicio

        for i in range(periodos):
            linha = linha_inicial + i
            ciclo = ws_report.range(f"E{linha}").value

            if not ciclo:
                break

            ws_report.range(f"C{linha}").value = data_atual

            if str(ciclo).strip().upper() == "00x06":
                data_atual += timedelta(days=1)

    # ===== FUNÇÕES FINANCEIRAS E AJUSTES =====

    def _OC(self, wb1_path, wb2):
        ws = wb2.sheets["FRONT VIGIA"]
        if str(ws["G16"].value).strip().upper() == "O.C.:": 
            ws["H16"].value = input("OC: ")

    def _credit_note(self, wb, valor_c21):
        if "Credit Note" in [s.name for s in wb.sheets]:
            wb.sheets["Credit Note"]["C21"].value = valor_c21

    def _obter_proxima_nf(self, pasta_nfs):
        if not os.path.exists(pasta_nfs):
            return 1
        numeros = [int(re.match(r"(\d+)", f).group(1)) for f in os.listdir(pasta_nfs) if re.match(r"(\d+)", f) and f.lower().endswith(".pdf")]
        return max(numeros) + 1 if numeros else 1

    def _colar_nf(self, ws, celula, numero_nf):
        ws[celula].value = f"NF.: {numero_nf}"

    def _MMO(self, wb1, wb2):
        try:
            ws_report = wb2.sheets["REPORT VIGIA"]
        except:
            return

        if str(ws_report["E25"].value).strip().upper() != "MMO":
            return

        try:
            ws_resumo = wb1.sheets["Resumo"]
        except:
            return

        valores_g = ws_resumo.range("G1:G1000").value
        valores_limpos = [v for v in valores_g if v not in (None, "")]
        if not valores_limpos:
            return

        try:
            texto = str(valores_limpos[-1]).replace("R$", "").replace(".", "").replace(",", ".").strip()
            valor = float(texto)
        except:
            valor = 0.0

        ws_report["F25"].value = valor
        ws_report["F25"].number_format = "#,##0.00"

    def _cargonave(self, ws):
        valor = ws.range("C9").value
        return str(valor).strip().upper() == "A/C AGÊNCIA MARÍTIMA CARGONAVE LTDA."

    def _arredondar_para_baixo_50(self, ws_front):
        if not self._cargonave(ws_front):
            return
        valor = ws_front.range("E37").value
        if valor is None:
            return
        try:
            resultado = (int(valor) // 50) * 50
            ws_front.range("H28").value = resultado
        except:
            pass

    def _escrever_nf(self, wb_faturamento, nome_navio, dn):
        ws_nf = None
        for sheet in wb_faturamento.sheets:
            if sheet.name.strip().lower() == "nf":
                ws_nf = sheet
                break
        if not ws_nf:
            return

        ano = datetime.now().year
        texto = f"SERVIÇO PRESTADO DE ATENDIMENTO/APOIO AO M/V {nome_navio}\nDN {dn}/{ano}"

        cel = ws_nf.range("A1")
        cel.value = texto
        ws_nf.range("A1:E2").merge()
        cel.api.HorizontalAlignment = -4108
        cel.api.VerticalAlignment = -4108
        cel.api.WrapText = True
        cel.api.Font.Bold = True
        cel.api.Font.Size = 14

    def _quitacao(self, wb, valor_c21):
        if "Quitação" not in [s.name for s in wb.sheets]:
            return

        ws = wb.sheets["Quitação"]
        ws["C22"].value = valor_c21

        pasta_nfs = r"C:\Users\Carol\SANPORT LOGÍSTICA PORTUÁRIA LTDA\Central de Documentos - Documentos\2.2 CONTABILIDADE 2025\12 - DEZEMBRO"
        proxima_nf = self._obter_proxima_nf(pasta_nfs)
        self._colar_nf(ws, "H22", proxima_nf)



# ==============================
# CLASSE 2: FATURAMENTO DE ACORDO
# ========================


class FaturamentoDeAcordo:      
    def executar(self):
        print("🚀 Iniciando Faturamento De Acordo...")

        try:
            pasta_faturamentos = obter_pasta_faturamentos()
            resultado = abrir_workbooks(pasta_faturamentos)

            if not resultado:
                return

            app, wb1, wb2, ws1, ws_front = resultado

            try:
                pasta_navio = Path(wb1.fullname).parent
                dn = obter_dn_da_pasta(pasta_navio)
                nome_navio = obter_nome_navio_da_pasta(pasta_navio)

                print(f"📋 DN: {dn}")
                print(f"🚢 Navio: {nome_navio}")

                # Preenchimento FRONT VIGIA (seu código antigo aqui)
                hoje = datetime.now()
                meses = ["", "janeiro", "fevereiro", "março", "abril", "maio", "junho",
                         "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"]
                data_extenso = f"{hoje.day} de {meses[hoje.month]} de {hoje.year}"

                ws_front.range("D15").value = nome_navio
                ws_front.range("C21").value = f"DN: {dn}/{hoje.year}"
                ws_front.range("D16").value = data_extenso
                ws_front.range("D17").value = data_extenso
                ws_front.range("D18").value = "-"
                ws_front.range("C26").value = f'  DE ACORDO ( M/V "{nome_navio}" )'
                ws_front.range("C27").value = '  VOY SA02325'
                ws_front.range("G27").value = None

                cliente_c9 = str(ws_front.range("C9").value or "").strip()
                if "Unimar Agenciamentos" in cliente_c9:
                    ws_front.range("G26").value = 400

                try:
                    valor = float(ws_front.range("C35").value or 0)
                    ws_front.range("C35").value = valor + 20
                except:
                    ws_front.range("C35").value = 20

                ws_front.range("C39").value = f"Santos, {data_extenso}"

                # Remove outras abas
                for sheet in list(wb2.sheets):
                    if sheet.name != ws_front.name:
                        sheet.delete()

                # Salva
                desktop = Path.home() / "Desktop"
                arquivo_final = desktop / f"3 - DN_{dn}.xlsx"
                wb2.save(str(arquivo_final))

                print(f"\n✅ Faturamento De Acordo concluído!")
                print(f"Arquivo salvo em: {arquivo_final}")

            finally:
                fechar_workbooks(app, wb1, wb2)

        except Exception as e:
            print(f"❌ Erro: {e}")


# ==============================
# MENU PRINCIPAL
# ==============================

class CentralSanport:
    def __init__(self):
        self.completo = FaturamentoCompleto()
        self.de_acordo = FaturamentoDeAcordo()
        self.opcoes = [
            "FATURAMENTO",
            "DE ACORDO",
            "SAIR DO PROGRAMA"
        ]

    def limpar_tela(self):
        os.system('cls' if os.name == 'nt' else 'clear')

    def mostrar_menu(self, selecionado=0):
        self.limpar_tela()

        print("╔" + "═" * 62 + "╗")
        print(f"║{' 🚢 CENTRAL DE PROCESSOS - SANPORT 🚢 '.center(60)}║")
        print("╚" + "═" * 62 + "╝\n")

        for i, opcao in enumerate(self.opcoes):
            if i == selecionado:
                print(f"          ►► {opcao} ◄◄")
            else:
                print(f"              {opcao}")

        print("\n" + "═" * 64)
        print("   ↑ ↓ = Navegar     ENTER = Selecionar     0 = Sair")
        print("═" * 64)

    def rodar(self):
        selecionado = 0
        self.mostrar_menu(selecionado)

        while True:
            key = msvcrt.getch()

            # Trata prefixos comuns de teclas especiais
            if key in (b'\xe0', b'\x00'):
                key2 = msvcrt.getch()
                if key2 == b'H':  # Seta cima
                    selecionado = max(0, selecionado - 1)
                    self.mostrar_menu(selecionado)
                elif key2 == b'P':  # Seta baixo
                    selecionado = min(len(self.opcoes) - 1, selecionado + 1)
                    self.mostrar_menu(selecionado)
                continue

            # Enter (vários códigos possíveis)
            if key in (b'\r', b'\n'):
                self.limpar_tela()

                if selecionado == 0:
                    print("╔" + "═" * 62 + "╗")
                    print("║" + " INICIANDO FATURAMENTO... ".center(60) + "║")
                    print("╚" + "═" * 62 + "╝\n")
                    try:
                        self.completo.executar()
                    except Exception as e:
                        print(f"❌ ERRO: {e}\n")

                elif selecionado == 1:
                    print("╔" + "═" * 62 + "╗")
                    print("║" + " INICIANDO DE ACORDO... ".center(60) + "║")
                    print("╚" + "═" * 62 + "╝\n")
                    try:
                        self.de_acordo.executar()
                    except Exception as e:
                        print(f"❌ ERRO: {e}\n")

                elif selecionado == 2:
                    self.limpar_tela()
                    print("\n" + "╔" + "═" * 62 + "╗")
                    print("║" + " OBRIGADO POR USAR O SISTEMA! ".center(60) + "║")
                    print("║" + " Até logo, capitão! 🚢 ".center(60) + "║")
                    print("╚" + "═" * 62 + "╝\n")
                    time.sleep(2)
                    break

                input("\nPressione ENTER para continuar...")
                self.mostrar_menu(selecionado)

            # Tecla 0 para sair rápido
            elif key == b'0':
                self.limpar_tela()
                print("\n" + "╔" + "═" * 62 + "╗")
                print("║" + " Saindo do programa... ".center(60) + "║")
                print("╚" + "═" * 62 + "╝")
                time.sleep(1)
                break

            # Qualquer outra tecla: ignora ou atualiza menu
            else:
                self.mostrar_menu(selecionado)



if __name__ == "__main__":
    CentralSanport().rodar()