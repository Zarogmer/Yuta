# ==============================
# IMPORTS LIMPOS (sem duplicados)
# ==============================
import sys
import re
import ssl
import certifi
import urllib.request
from datetime import datetime, date, timedelta, timezone
from pathlib import Path
from tkinter import Tk, filedialog
import xlwings as xw
import pandas as pd
from itertools import cycle
import msvcrt
import os
import holidays
from datetime import datetime



feriados_br = holidays.Brazil()


# ==============================
# FUNÇÕES AUXILIARES GLOBAIS
# ==============================


def obter_pasta_faturamentos(pasta_cliente: str | Path, nome_cliente: str) -> Path:
    """
    Localiza o arquivo de faturamento do cliente dentro da pasta FATURAMENTOS.
    Mantém o nome da função para compatibilidade com o código existente.
    """
    pasta_cliente = Path(pasta_cliente)

    if not pasta_cliente.exists():
        raise FileNotFoundError(f"Pasta base não encontrada: {pasta_cliente}")

    # 🔹 Garante que estamos na pasta FATURAMENTOS
    if pasta_cliente.name.upper() != "FATURAMENTOS":
        pasta_faturamentos = pasta_cliente.parent
    else:
        pasta_faturamentos = pasta_cliente

    if not pasta_faturamentos.exists():
        raise FileNotFoundError("❌ Pasta FATURAMENTOS não encontrada")

    nome_cliente = nome_cliente.strip().upper()

    # 🔹 Procura o arquivo do cliente
    for arq in pasta_faturamentos.glob("*.xlsx"):
        if arq.stem.strip().upper() == nome_cliente:
            print(f"📂 Arquivo de faturamento encontrado: {arq.name}\n")
            return arq

    raise FileNotFoundError(
        f"❌ Arquivo de faturamento '{nome_cliente}.xlsx' não encontrado em {pasta_faturamentos}"
    )

def selecionar_arquivo_navio() -> str | None:
    """
    Abre uma janela para o usuário selecionar o arquivo NAVIO (.xlsx).
    A janela aparece em primeiro plano.
    """
    root = Tk()
    root.lift()                         # coloca a janela na frente
    root.attributes("-topmost", True)   # força ficar em primeiro plano
    root.focus_force()                  # força foco
    root.update()                       # aplica imediatamente
    caminho = filedialog.askopenfilename(
        title="Selecione o arquivo do NAVIO",
        filetypes=[("Arquivos Excel", "*.xlsx")]
    )
    root.destroy()  # fecha a janela principal depois da seleção

    if not caminho:
        return None

    print(f"📂 Arquivo NAVIO selecionado: {Path(caminho).name}")
    return caminho

def abrir_workbooks():


    # 1️⃣ Seleciona arquivo NAVIO
    arquivo_navio = selecionar_arquivo_navio()
    if not arquivo_navio:
        print("❌ Nenhum arquivo do NAVIO selecionado")
        return None

    pasta_navio = Path(arquivo_navio).parent
    # nome do cliente = pasta acima do navio (ex: UNIMAR)
    nome_cliente = pasta_navio.parent.name.strip().upper()

    # pasta FATURAMENTOS direto no Desktop
    pasta_faturamentos = Path.home() / "Desktop" / "FATURAMENTOS"

    try:
        arquivo_cliente = obter_pasta_faturamentos(
            pasta_cliente=pasta_faturamentos,
            nome_cliente=nome_cliente
        )
    except Exception as e:
        print(e)
        return None

    # 3️⃣ Abre Excel
    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False

    wb_navio = app.books.open(str(arquivo_navio))
    wb_cliente = app.books.open(str(arquivo_cliente))

    ws_navio = wb_navio.sheets["Resumo"]
    ws_front = wb_cliente.sheets["FRONT VIGIA"]

    return app, wb_navio, wb_cliente, ws_navio, ws_front


def obter_dn_da_pasta(pasta: Path) -> str:
    """
    Extrai o número DN do nome da pasta.
    Se não encontrar, retorna '0000' e exibe aviso.
    """
    numeros = re.findall(r"\d+", pasta.name)
    if not numeros:
        print(
            f"⚠️ Não foi possível identificar o DN no nome da pasta '{pasta.name}', usando '0000' como padrão"
        )
        return "0000"
    return numeros[0]

def obter_nome_navio_da_pasta(pasta: Path) -> str:
    nome_limpo = re.sub(r"^\d+[\s\-_]*", "", pasta.name, flags=re.IGNORECASE).strip()
    return nome_limpo if nome_limpo else "NAVIO NÃO IDENTIFICADO"


def fechar_workbooks(app=None, wb_navio=None, wb_cliente=None, arquivo_saida: Path | None = None):
    """
    Fecha e salva workbooks de forma defensiva.
    - app: instância xlwings.App (ou None)
    - wb_navio: workbook do navio (opcional)
    - wb_cliente: workbook do cliente (opcional) -> salvo em arquivo_saida se fornecido
    - arquivo_saida: Path onde salvar o Excel e gerar PDF (opcional)
    """
    # salva/expor e fecha com muitos try/except para evitar exceções RPC
    try:
        if wb_cliente and arquivo_saida:
            try:
                # tenta salvar o workbook do cliente
                wb_cliente.save(str(arquivo_saida))
            except Exception as e:
                print(f"⚠️ Erro ao salvar Excel: {e}")

            try:
                # tenta exportar para PDF (pode falhar se Excel estiver em estado inválido)
                pdf_saida = arquivo_saida.with_suffix(".pdf")
                wb_cliente.api.ExportAsFixedFormat(0, str(pdf_saida))
            except Exception as e:
                print(f"⚠️ Erro ao exportar PDF: {e}")

            try:
                print(f"💾 Arquivo Excel salvo em: {arquivo_saida}")
                print(f"📑 PDF gerado em: {pdf_saida}")
            except Exception:
                pass

        # fecha workbooks individualmente, cada um em try/except
        try:
            if wb_navio:
                wb_navio.close()
        except Exception as e:
            print(f"⚠️ Erro ao fechar wb_navio: {e}")

        try:
            if wb_cliente:
                wb_cliente.close()
        except Exception as e:
            print(f"⚠️ Erro ao fechar wb_cliente: {e}")

    finally:
        # encerra a app somente se ela existir e parecer válida
        try:
            if app:
                # proteção extra: algumas versões de xlwings/COM podem lançar se já encerrado
                app.quit()
        except Exception as e:
            print(f"⚠️ Erro ao encerrar app Excel: {e}")



# ==============================
# LICENÇA E DATA
# ==============================


def data_online():
    context = ssl.create_default_context(cafile=certifi.where())
    req = urllib.request.Request(
        "https://www.cloudflare.com", headers={"User-Agent": "Mozilla/5.0"}
    )
    with urllib.request.urlopen(req, context=context, timeout=5) as r:
        data_str = r.headers["Date"]
    dt_utc = datetime.strptime(data_str, "%a, %d %b %Y %H:%M:%S %Z").replace(
        tzinfo=timezone.utc
    )
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



class FaturamentoCompleto:
    def __init__(self, g_logic=1):
        self.app = None
        self.wb1 = None
        self.wb2 = None
        self.ws1 = None
        self.ws_front = None
        self.nome_navio = None
        self.g_logic = g_logic


    # ---------------- fluxo principal ----------------
    def executar(self):
        print("🚀 Iniciando execução...")
        resultado = abrir_workbooks()  # função externa esperada
        if not resultado:
            raise SystemExit("❌ Erro ou pasta inválida")
        self.app, self.wb1, self.wb2, self.ws1, self.ws_front = resultado
        print("📂 Workbooks abertos com sucesso!")
        try:
            self.processar()
        except Exception as e:
            print(f"❌ Erro durante o processamento: {e}")
            # tenta salvar/fechar com segurança usando a função global
            try:
                pasta_saida = Path(self.wb1.fullname).parent if self.wb1 else None
                arquivo_saida = (pasta_saida / "3.xlsx") if pasta_saida else None
                fechar_workbooks(self.app, self.wb1, self.wb2, arquivo_saida)
            except Exception as e2:
                print(f"⚠️ Erro ao fechar após falha: {e2}")
            # re-levanta para o menu tratar ou encerra conforme seu fluxo
            raise


    def processar(self):
        """
        Fluxo principal do faturamento:
        1) Preenche FRONT VIGIA
        2) Atualiza o REPORT VIGIA (E, G e C)
        """

        try:
            print("DEBUG hasattr FRONT:", hasattr(self, "preencher_front_vigia"))


            # ---------- FRONT ----------
            self.preencher_front_vigia()

            # ---------- REPORT ----------
            ws_report = self.wb2.sheets["REPORT VIGIA"]

            # 1️⃣ quantidade de períodos (APENAS PARA INSERIR LINHAS)
            qtd_periodos = self.obter_periodos(self.ws1)
            print("DEBUG: qtd_periodos =", qtd_periodos)

            self.inserir_linhas_report(
                ws_report,
                linha_inicial=22,
                periodos=qtd_periodos
            )

            # 2️⃣ COLUNA E → LISTA DE PERÍODOS
            periodos = self.preencher_coluna_E(
                ws_report,
                linha_inicial=22,
                debug=True
            )
            print("DEBUG: períodos E =", periodos)

            # 3️⃣ COLUNA G → VALORES (BASEADO NA E)
            self.preencher_coluna_G(
                ws_report,
                self.ws1,              # ws_resumo
                linha_inicial=22,
                periodos=periodos,     # 🔥 lista
                debug=True
            )

            # 4️⃣ COLUNA C → DATAS (USA QUANTIDADE DA E)
            self.montar_datas_report_vigia(
                ws_report,
                self.ws1,
                linha_inicial=22,
                periodos=len(periodos)  # ⚠️ INT
            )

            print("✅ REPORT VIGIA atualizado com sucesso!")

        except Exception as e:
            print(f"❌ Erro ao atualizar REPORT: {e}")
            raise

        # ---------- SALVA E FECHA ----------
        try:
            pasta_saida = Path(self.wb1.fullname).parent if self.wb1 else None
            arquivo_saida = (pasta_saida / "3.xlsx") if pasta_saida else None

            fechar_workbooks(self.app, self.wb1, self.wb2, arquivo_saida)

            print(f"💾 Arquivo Excel salvo em: {arquivo_saida}")

        except Exception as e:
            print(f"⚠️ Erro ao salvar/fechar workbooks: {e}")
            raise




# ===== FRONT ======#

    def preencher_front_vigia(self):
        """
        Preenche a aba FRONT VIGIA com:
        - Nome do navio
        - DN
        - Warehouse / Berço (entrada do usuário)
        - Data inicial e final
        - Rodapé com data atual
        """
        try:
            # Pasta do navio
            pasta_navio = Path(self.wb1.fullname).parent
            dn = obter_dn_da_pasta(pasta_navio)
            self.nome_navio = obter_nome_navio_da_pasta(pasta_navio)
            ano = datetime.now().year
            texto_dn = f"DN: {dn}/{ano}"

            # ---------------- FRONT ----------------
            try:
                # Nome do navio
                self.ws_front.range("D15").value = self.nome_navio
                # DN
                self.ws_front.range("C21").value = texto_dn
                # Warehouse / Berço
                self.ws_front.range("D18").value = input("WAREHOUSE / BERÇO: ").upper()
            except Exception as e:
                print(f"⚠️ Erro ao preencher FRONT (nome/DN/berço): {e}")

            # ---------------- DATAS ----------------
            data_min, data_max = self.obter_datas_extremos(self.ws1)
            try:
                if data_min:
                    self.ws_front.range("D16").value = self.data_por_extenso(data_min)
                if data_max:
                    self.ws_front.range("D17").value = self.data_por_extenso(data_max)
            except Exception as e:
                print(f"⚠️ Erro ao preencher datas no FRONT: {e}")

            # ---------------- RODAPÉ ----------------
            try:
                hoje = datetime.now()
                meses = ["", "janeiro","fevereiro","março","abril","maio","junho",
                        "julho","agosto","setembro","outubro","novembro","dezembro"]
                self.ws_front.range("C39").value = f"Santos, {hoje.day} de {meses[hoje.month]} de {hoje.year}"
            except Exception as e:
                print(f"⚠️ Erro ao preencher rodapé do FRONT: {e}")

            print("✅ FRONT VIGIA preenchido com sucesso!")

        except Exception as e:
            print(f"❌ Erro geral ao preencher FRONT VIGIA: {e}")
            raise



#==================== REPORT =====================#

    def inserir_linhas_report(self, ws_report, linha_inicial, periodos):
        if periodos <= 1:
            return
        app = ws_report.book.app
        app.screen_updating = False
        app.display_alerts = False
        app.enable_events = False
        app.calculation = 'manual'
        try:
            linha_modelo = linha_inicial
            for i in range(periodos - 1):
                destino = linha_inicial + 1 + i
                ws_report.api.Rows(destino).Insert()
                ws_report.api.Rows(linha_modelo).Copy(ws_report.api.Rows(destino))
        finally:
            app.calculation = 'automatic'
            app.enable_events = True
            app.display_alerts = True
            app.screen_updating = True


# ===== LINHA E=====#


    def gerar_ciclos_coluna_E(self, ws_resumo, linha_inicial=2):
        """
        Gera a lista de períodos para a coluna E do REPORT, baseada na data mais antiga.
        Sequência: 06x12 -> 12x18 -> 18x24 -> 00x06
        """
        sequencia_padrao = ["06x12", "12x18", "18x24", "00x06"]

        # 1️⃣ Encontrar a primeira data válida (mais antiga que hoje)
        last_row = ws_resumo.used_range.last_cell.row
        valores = ws_resumo.range(f"B{linha_inicial}:B{last_row}").value
        hoje = date.today()
        primeira_linha_data = None

        for idx, v in enumerate(valores):
            if v in (None, "", "Total"):
                continue
            if isinstance(v, datetime):
                d = v.date()
            elif isinstance(v, str):
                try:
                    d = datetime.strptime(v.strip(), "%d/%m/%Y").date()
                except:
                    continue
            else:
                continue

            if d < hoje:
                primeira_linha_data = linha_inicial + idx
                break

        if primeira_linha_data is None:
            return []  # nenhuma data antiga encontrada

        # 2️⃣ Contar espaços vazios ou "Total" antes da próxima data
        contador_vazio = 0
        for i in range(primeira_linha_data + 1, last_row + 1):
            valor = ws_resumo.range(f"B{i}").value
            if valor in (None, "", "Total"):
                contador_vazio += 1
            else:
                break

        # 3️⃣ Definir primeiro período
        if contador_vazio >= 4:
            primeiro_periodo = "06x12"
        elif contador_vazio == 3:
            primeiro_periodo = "12x18"
        elif contador_vazio == 2:
            primeiro_periodo = "18x24"
        else:
            primeiro_periodo = "00x06"

        # 4️⃣ Sequência cíclica
        idx_inicio = sequencia_padrao.index(primeiro_periodo)
        sequencia_ciclica = sequencia_padrao[idx_inicio:] + sequencia_padrao[:idx_inicio]

        # 5️⃣ Gerar lista completa de períodos
        total_periodos = self.obter_periodos(ws_resumo)
        ciclos = [sequencia_ciclica[i % 4] for i in range(total_periodos)]

        return ciclos


    
    def preencher_coluna_E(self, ws_report, linha_inicial=22, debug=False):
        """
        Preenche a coluna E do REPORT VIGIA com os períodos gerados.
        """
        try:
            ciclos = self.gerar_ciclos_coluna_E(self.ws1)
            for idx, p in enumerate(ciclos):
                ws_report.range(f"E{linha_inicial + idx}").value = p
            if debug:
                print("DEBUG: Coluna E preenchida com:", ciclos)
            return ciclos
        except Exception as e:
            print(f"❌ Erro ao preencher coluna E: {e}")
            raise

 

    # ===== LINHA G=====#

    def normalizar_periodo(self, valor_c):
        if not valor_c:
            return None

        s = str(valor_c).strip().lower()
        if s.startswith("06"):
            return "06x12"
        if s.startswith("12"):
            return "12x18"
        if s.startswith("18"):
            return "18x24"
        if s.startswith("00"):
            return "00x06"
        return None
        
    def gerar_valores_coluna_G(self, ws_resumo, periodos_E, debug=False):
        mapa = self.extrair_valores_por_periodo(ws_resumo, debug=debug)
        contadores = {k: 0 for k in mapa}
        valores_g = []

        for p in periodos_E:
            if p in mapa and contadores[p] < len(mapa[p]):
                valor = mapa[p][contadores[p]]
                contadores[p] += 1
            else:
                valor = 0.0

            valores_g.append(valor)

        if debug:
            print("DEBUG G FINAL:", valores_g)

        return valores_g


    def preencher_coluna_G(self, ws_report, ws_resumo, linha_inicial=22, periodos=None, debug=False):
        """
        Preenche a coluna G seguindo EXATAMENTE a ordem da coluna E
        e formatando como moeda com 2 casas decimais.
        """

        if not periodos:
            raise ValueError("periodos (lista da coluna E) é obrigatório")

        valores = self.gerar_valores_coluna_G(
            ws_resumo,
            periodos,
            debug=debug
        )

        if debug:
            print("DEBUG G FINAL:", valores)

        for i, valor in enumerate(valores):
            cell = ws_report.range(f"G{linha_inicial + i}")
            cell.value = valor              # número cru, sem arredondar
            cell.api.NumberFormatLocal = 'R$ #.##0,00'




    def extrair_valores_por_periodo(self, ws_resumo, debug=False):
        last_row = ws_resumo.used_range.last_cell.row

        mapa = {
            "00x06": [],
            "06x12": [],
            "12x18": [],
            "18x24": []
        }

        for i in range(2, last_row + 1):
            c = ws_resumo.range(f"C{i}").value
            z = ws_resumo.range(f"Z{i}").value

            if not c or z is None:
                continue

            s_c = str(c).strip().lower()
            if s_c.startswith("total"):
                continue

            periodo = self.normalizar_periodo(s_c)
            if not periodo:
                continue

            try:
                # ✅ CASO 1: já é número no Excel
                if isinstance(z, (int, float)):
                    valor = float(z)

                # ✅ CASO 2: veio como texto "R$ 1.144,70"
                else:
                    valor = (
                        str(z)
                        .replace("R$", "")
                        .replace(".", "")
                        .replace(",", ".")
                        .strip()
                    )
                    valor = float(valor)

            except:
                continue


            mapa[periodo].append(valor)

            if debug:
                print(f"DEBUG MAPA: {periodo} += {valor}")

        return mapa
    
    def extrair_numero_excel(self, z):
        """
        Garante conversão correta de valores do Excel
        independente de vir como float ou string pt-BR.
        """

        # 👉 Caso 1: Excel já entregou número
        if isinstance(z, (int, float)):
            return float(z)

        # 👉 Caso 2: Veio como texto (ex: "1.144,70")
        s = str(z).strip()

        if not s:
            raise ValueError("valor vazio")

        s = (
            s.replace("R$", "")
            .replace(" ", "")
            .replace(".", "")
            .replace(",", ".")
        )

        return float(s)





    # ===== LINHA C =====#

    def montar_datas_report_vigia(self, ws_report, ws_resumo, linha_inicial=22, periodos=None):
        if periodos is None:
            raise ValueError("É necessário informar 'periodos' para preencher as datas")
        data_inicio, _ = self.obter_datas_extremos(ws_resumo)
        if not data_inicio:
            raise ValueError("Não foi possível determinar a data inicial na aba RESUMO")
        data_atual = data_inicio
        for i in range(periodos):
            linha = linha_inicial + i
            ciclo = ws_report.range(f"E{linha}").value
            if ciclo in (None, ""):
                break
            ws_report.range(f"C{linha}").value = data_atual
            if isinstance(ciclo, str) and ciclo.strip().lower() == "00x06":
                data_atual += timedelta(days=1)
        return periodos

# ===== ABAS ESPECIFICAS =====#

    def MMO(self, arquivo1, wb2):
        """
        Processa MMO sem abrir arquivo na rede (zero permission denied).
        wb2 é o wb_navio (tem "Resumo")
        Escreve em "REPORT VIGIA" do wb2
        """
        print("   Iniciando MMO...")

        try:
            ws_report = wb2.sheets["REPORT VIGIA"]
        except:
            print("   ⚠️ Aba 'REPORT VIGIA' não encontrada. Pulando MMO.")
            return

        if str(ws_report.range("E25").value).strip().upper() != "MMO":
            print("   MMO não necessário (E25 != 'MMO').")
            return

        try:
            ws_resumo = wb2.sheets["Resumo"]
        except:
            print("   ⚠️ Aba 'Resumo' não encontrada. Pulando MMO.")
            return

        print("   Lendo coluna G...")
        valores_g = ws_resumo.range("G1:G1000").value
        valores_limpos = [v for v in valores_g if v is not None]

        if not valores_limpos:
            print("   Coluna G vazia. Pulando MMO.")
            return

        ultimo_valor = valores_limpos[-1]

        try:
            texto = str(ultimo_valor).replace("R$", "").replace(" ", "").strip()
            texto = texto.replace(".", "").replace(",", ".")
            ultimo_float = float(texto)
        except:
            print(f"   Erro ao converter '{ultimo_valor}'. Usando 0.")
            ultimo_float = 0.0

        ws_report.range("F25").value = ultimo_float
        ws_report.range("F25").api.NumberFormat = "#,##0.00"

        print(f"   ✅ MMO concluído: R$ {ultimo_float:,.2f} escrito em F25")

    def cargonave(self, ws):
        valor_c9 = ws.range("C9").value
        return str(valor_c9).strip().upper() == "A/C AGÊNCIA MARÍTIMA CARGONAVE LTDA."

    def arredondar_para_baixo_50(self, ws_front_vigia):
        if not self.cargonave(ws_front_vigia):
            return
        valor = ws_front_vigia.range("E37").value
        if valor is None:
            return
        try:
            resultado = (int(valor) // 50) * 50
        except:
            return
        ws_front_vigia.range("H28").value = resultado


# ===== DATAS / UTILITARIOS =====#

    def data_por_extenso(self, valor):
        if isinstance(valor, datetime):
            data = valor
        elif isinstance(valor, date):
            data = datetime(valor.year, valor.month, valor.day)
        elif isinstance(valor, str):
            try:
                data = datetime.strptime(valor, "%d/%m/%Y")
            except:
                return ""
        else:
            return ""
        return data.strftime("%d de %B de %Y")

    def obter_datas_extremos(self, ws_resumo):
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
                continue
            if isinstance(v, str):
                s = v.strip()
                try:
                    d = datetime.strptime(s, "%d/%m/%Y").date()
                    if d != hoje:
                        datas.append(d)
                    continue
                except:
                    pass
        if not datas:
            return None, None
        return min(datas), max(datas)

    def obter_periodos(self, ws_resumo):
        """
        Retorna o último valor válido da coluna AA como inteiro.
        """
        last_row = ws_resumo.used_range.last_cell.row
        # Lê toda a coluna AA
        valores = ws_resumo.range(f"AA1:AA{last_row}").value

        if not valores:
            return 1  # padrão

        # Garante que 'valores' seja uma lista
        if not isinstance(valores, list):
            valores = [valores]

        # Procura o último valor não vazio
        for v in reversed(valores):
            if v is not None and v != "":
                try:
                    return int(float(v))
                except:
                    continue

        return 1


# ==============================
# CLASSE 2: FATURAMENTO DE ACORDO
# ==============================


class FaturamentoDeAcordo:

    def executar(self):
        print("🚀 Iniciando Faturamento De Acordo...")

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

            hoje = datetime.now()
            meses = [
                "",
                "janeiro",
                "fevereiro",
                "março",
                "abril",
                "maio",
                "junho",
                "julho",
                "agosto",
                "setembro",
                "outubro",
                "novembro",
                "dezembro",
            ]
            data_extenso = f"{hoje.day} de {meses[hoje.month]} de {hoje.year}"

            ws_front.range("D15").value = nome_navio
            ws_front.range("C21").value = f"DN: {dn}/{hoje.year}"
            ws_front.range("D16").value = data_extenso
            ws_front.range("D17").value = data_extenso
            ws_front.range("D18").value = "-"
            ws_front.range("C26").value = f'  DE ACORDO ( M/V "{nome_navio}" )'
            ws_front.range("C27").value = "  VOY SA02325"
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

            print("\n✅ Faturamento De Acordo concluído!")
            print(f"📁 Arquivo salvo em: {arquivo_final}")

        finally:
            try:
                wb1.close()
                wb2.close()
                app.quit()
            except:
                pass


# ==============================
# CLASSE 3: Fazer Ponto
# ==============================


class ProgramaCopiarPeriodo:
    PERIODOS_MENU = {"1": "06h", "2": "12h", "3": "18h", "4": "00h"}

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
        "24h": "00h",
    }

    EQUIVALENTES = {
        "06h": ["06h", "12h"],
        "12h": ["12h", "06h"],
        "18h": ["18h", "00h"],
        "00h": ["00h", "18h"],
    }

    BLOCOS = {"06h": 1, "12h": 1, "18h": 2, "00h": 2}

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
            if (
                isinstance(valor, (datetime, date))
                and valor.strftime("%d/%m/%Y") == data_str
            ):
                return i
            elif valor == data_str:
                return i
        raise Exception(f"Data {data_str} não encontrada.")

    def encontrar_total_data(self, linha_data):
        i = linha_data + 1
        while True:
            valor_c = self.ws.range(f"C{i}").value
            valor_a = self.ws.range(f"A{i}").value
            if (
                isinstance(valor_a, str)
                and self.normalizar_texto(valor_a) == "totalgeral"
            ):
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

                if (
                    isinstance(valor_a, str)
                    and self.normalizar_texto(valor_a) == "totalgeral"
                ):
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
            if (
                isinstance(valor_a, str)
                and self.normalizar_texto(valor_a) == "totalgeral"
            ):
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


# ==============================
# MENU PRINCIPAL
# ==============================


class CentralSanport:
    def __init__(self):
        self.completo = FaturamentoCompleto()
        self.de_acordo = FaturamentoDeAcordo()

        self.opcoes = ["FATURAMENTO", "DE ACORDO", "FAZER PONTO", "SAIR DO PROGRAMA"]

    # =========================
    # UTILITÁRIOS
    # =========================
    def limpar_tela(self):
        os.system("cls" if os.name == "nt" else "clear")

    def limpar_buffer_teclado(self):
        while msvcrt.kbhit():
            msvcrt.getch()

    def pausar_e_voltar(self, selecionado):
        print("\n🔁 Pressione ENTER para voltar ao menu...")
        while True:
            key = msvcrt.getch()
            if key in (b"\r", b"\n"):
                self.limpar_buffer_teclado()
                self.mostrar_menu(selecionado)
                return

    # =========================
    # MENU
    # =========================
    def mostrar_menu(self, selecionado):
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
        print("   ↑ ↓ = Navegar     ENTER = Selecionar")
        print("═" * 64)

    # =========================
    # EXECUÇÃO PRINCIPAL
    # =========================

    # =========================
    # Dentro da classe CentralSanport
    # =========================

    def rodar(self):
        selecionado = 0
        self.mostrar_menu(selecionado)

        while True:
            key = msvcrt.getch()

            # SETAS
            if key in (b"\xe0", b"\x00"):
                key = msvcrt.getch()

                if key == b"H":  # ↑
                    selecionado = max(0, selecionado - 1)
                    self.mostrar_menu(selecionado)

                elif key == b"P":  # ↓
                    selecionado = min(len(self.opcoes) - 1, selecionado + 1)
                    self.mostrar_menu(selecionado)

                continue

            # ENTER
            if key in (b"\r", b"\n"):
                self.limpar_tela()

                # ----------------------------
                # FATURAMENTO
                # ----------------------------
                if selecionado == 0:
                    print("╔" + "═" * 62 + "╗")
                    print("║" + " INICIANDO FATURAMENTO... ".center(60) + "║")
                    print("╚" + "═" * 62 + "╝\n")

                    try:
                        # Chama o FaturamentoCompleto
                        # NÃO precisa do input() dentro da classe
                        self.completo.executar()

                    except Exception as e:
                        print(f"\n❌ ERRO NO FATURAMENTO: {e}")

                    # Apenas aqui espera o ENTER para voltar
                    self.pausar_e_voltar(selecionado)

                # ----------------------------
                # DE ACORDO
                # ----------------------------
                elif selecionado == 1:
                    print("╔" + "═" * 62 + "╗")
                    print("║" + " INICIANDO DE ACORDO... ".center(60) + "║")
                    print("╚" + "═" * 62 + "╝\n")

                    try:
                        self.de_acordo.executar()
                    except Exception as e:
                        print(f"\n❌ ERRO: {e}")

                    self.pausar_e_voltar(selecionado)

                # ----------------------------
                # FAZER PONTO
                # ----------------------------
                elif selecionado == 2:
                    print("╔" + "═" * 62 + "╗")
                    print("║" + " INICIANDO FAZER PONTO... ".center(60) + "║")
                    print("╚" + "═" * 62 + "╝\n")

                    try:
                        pasta_faturamentos = obter_pasta_faturamentos()
                        if not pasta_faturamentos:
                            print("⛔ Operação cancelada")
                            self.pausar_e_voltar(selecionado)
                            continue

                        resultado = abrir_workbooks(pasta_faturamentos)
                        if not resultado:
                            print("⛔ Abertura cancelada")
                            self.pausar_e_voltar(selecionado)
                            continue

                        app, wb1, wb2, ws1, ws_front = resultado

                        programa = ProgramaCopiarPeriodo(
                            app=app, wb=wb2, ws=ws_front, debug=False
                        )
                        programa.executar()

                        arquivo_saida = Path(wb2.fullname).with_name(
                            "1_atualizado.xlsx"
                        )
                        fechar_workbooks(app, wb1, wb2, arquivo_saida)

                        print("\n✅ FAZER PONTO FINALIZADO COM SUCESSO")

                    except Exception as e:
                        print(f"\n❌ ERRO NO FAZER PONTO: {e}")

                    self.pausar_e_voltar(selecionado)

                # ----------------------------
                # SAIR
                # ----------------------------
                elif selecionado == 3:
                    self.limpar_tela()
                    print("\n👋 Saindo do programa...")
                    break


if __name__ == "__main__":
    CentralSanport().rodar()