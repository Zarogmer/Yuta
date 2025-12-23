# 1️⃣ – Importações
import xlwings as xw
from datetime import datetime, timedelta, timezone
from openpyxl.styles import Alignment
import pandas as pd
import os
import sys
import urllib.request
import ssl
import certifi
from pathlib import Path
import tkinter as tk
from tkinter import filedialog
from datetime import date
import re
import locale
from itertools import cycle

# =========================
# Funções utilitárias
# =========================

def abrir_workbooks():
    """
    Abre:
    - 1.xlsx da pasta do NAVIO (selecionada pelo usuário)
    - XLSX do CLIENTE em Desktop/FATURAMENTOS/<CLIENTE>.xlsx
    """

    # --- Seleção da pasta do navio ---
    root = tk.Tk()
    root.withdraw()
    pasta_navio = filedialog.askdirectory(
        title="Selecione a pasta do NAVIO (onde está o 1.xlsx)"
    )

    if not pasta_navio:
        print("Nenhuma pasta selecionada. Encerrando.")
        return None, None, None, None, None

    pasta_navio = Path(pasta_navio)
    pasta_cliente = pasta_navio.parent   # PASTA PAI
    nome_cliente = pasta_cliente.name

    arquivos_1 = list(pasta_navio.glob("1*.xls*"))

    if not arquivos_1:
        raise FileNotFoundError(
            f"Nenhum arquivo começando com '1' encontrado em:\n{pasta_navio}"
        )

    arquivo1 = arquivos_1[0]

    pasta_faturamentos = Path.home() / "Desktop" / "FATURAMENTOS"
    arquivo2 = pasta_faturamentos / f"{nome_cliente}.xlsx"

    app = xw.App(visible=False)
    wb1 = wb2 = None

    try:
        # --- Valida arquivos ---
        if not arquivo1.exists():
            raise FileNotFoundError(f"Arquivo 1.xlsx não encontrado:\n{arquivo1}")

        if not arquivo2.exists():
            raise FileNotFoundError(
                f"Arquivo de faturamento do cliente não encontrado:\n{arquivo2}"
            )

        # --- Abre workbooks ---
        wb1 = app.books.open(arquivo1)
        wb2 = app.books.open(arquivo2)

        ws1 = wb1.sheets[0]

        nomes_abas = [s.name for s in wb2.sheets]

        # Prioridade: aba com nome do cliente
        if nome_cliente in nomes_abas:
            ws_front = wb2.sheets[nome_cliente]

        # Fallback: FRONT VIGIA
        elif "FRONT VIGIA" in nomes_abas:
            ws_front = wb2.sheets["FRONT VIGIA"]

        else:
            raise RuntimeError(
                f"Nenhuma aba válida encontrada no faturamento.\n"
                f"Esperado: '{nome_cliente}' ou 'FRONT VIGIA'."
            )

        return app, wb1, wb2, ws1, ws_front

    except Exception as e:
        print(f"Erro ao abrir os arquivos: {e}")

        if wb1:
            wb1.close()
        if wb2:
            wb2.close()

        app.quit()
        return None, None, None, None, None

    
def fechar_workbooks(app, wb1=None, wb2=None, arquivo_saida=None):
    """
    Salva wb1 normalmente e salva wb2 EXCLUSIVAMENTE como 3.xlsx
    na pasta do cliente (ex: WILSON). Nunca salva em FATURAMENTOS.
    """
    try:
        if wb1:
            wb1.save()
            wb1.close()

        if wb2:
            if not arquivo_saida:
                raise RuntimeError(
                    "Caminho de saída não informado. "
                    "wb2 NÃO será salvo para evitar salvar em FATURAMENTOS."
                )

            wb2.save(arquivo_saida)
            wb2.close()

    finally:
        if app:
            app.quit()

def obter_dn_da_pasta(caminho_arquivo):
    """
    Extrai números do nome da pasta do navio
    Ex: '123 - NAVIO' -> '123'
    """
    pasta = os.path.basename(os.path.dirname(caminho_arquivo))

    numeros = re.findall(r"\d+", pasta)

    if not numeros:
        return None

    return numeros[0]  # primeiro bloco numérico


# ===== Licença =====#


def data_online():
    context = ssl.create_default_context(cafile=certifi.where())

    req = urllib.request.Request(
        "https://www.cloudflare.com",
        headers={"User-Agent": "Mozilla/5.0"}
    )

    with urllib.request.urlopen(req, context=context, timeout=5) as r:
        data_str = r.headers["Date"]

    dt_utc = datetime.strptime(
        data_str, "%a, %d %b %Y %H:%M:%S %Z"
    ).replace(tzinfo=timezone.utc)

    dt_local = dt_utc.astimezone()
    return dt_utc, dt_local


def validar_licenca():
    hoje_utc, hoje_local = data_online()

    limite = datetime(hoje_utc.year, hoje_utc.month, 30, tzinfo=timezone.utc)
    if hoje_utc > limite:
        sys.exit("⛔ Licença expirada")

    print(f"📅 Data local: {hoje_local.date()}")



def data_por_extenso(valor):
    if isinstance(valor, datetime):
        data = valor

    elif isinstance(valor, date):
        data = datetime(valor.year, valor.month, valor.day)

    elif isinstance(valor, str):
        try:
            data = datetime.strptime(valor, "%d/%m/%Y")
        except:
            return ""  # não inventa data

    else:
        return ""  # nunca usa datetime.now()

    return data.strftime("%d de %B de %Y")




def processar_front(ws1, ws_front):
    """
    Atualiza somente a aba FRONT VIGIA
    """

    # data atual por extenso (rodapé)
    meses = ["", "janeiro", "fevereiro", "março", "abril", "maio", "junho",
             "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"]

    hoje = datetime.now()
    ws_front.range("C39").value = (
        f"Santos, {hoje.day} de {meses[hoje.month]} de {hoje.year}"
    )

    # pega datas extremas do RESUMO
    data_min, data_max = obter_datas_extremos(ws1)

    # mostra no FRONT
    if data_min:
        ws_front.range("D16").value = data_por_extenso(data_min)

    if data_max:
        ws_front.range("D17").value = data_por_extenso(data_max)

    # 👉 retorna as datas para o main
    return data_min, data_max





# ===== DATA INICIAL E FINAL DO FRONT =====#

MESES_EN = {
    1: "JAN", 2: "FEB", 3: "MAR", 4: "APR",
    5: "MAY", 6: "JUN", 7: "JUL", 8: "AUG",
    9: "SEP", 10: "OCT", 11: "NOV", 12: "DEC"
}

# ===== ABAS ESPECIFICAS =====#


def OC(arquivo1, wb2):
    ws = wb2.sheets["FRONT VIGIA"]
    if str(ws["G16"].value).strip().upper() == "O.C.:":
        ws["H16"].value = input("OC: ")



def obter_nome_navio_da_pasta(caminho_arquivo):
    """
    Ex: '123 - NAVIO' -> 'NAVIO'
    """
    pasta = os.path.basename(os.path.dirname(caminho_arquivo))

    if "-" in pasta:
        return pasta.split("-", 1)[1].strip()

    return pasta.strip()

def obter_aba_nf_opcional(wb):
    for sheet in wb.sheets:
        nome = sheet.name.strip().lower()
        if nome == "nf" or nome.startswith("nf") or "nota" in nome:
            return sheet
    return None

def escrever_nf(wb_faturamento, nome_navio, dn):
    # tenta localizar aba NF
    ws_nf = None
    for sheet in wb_faturamento.sheets:
        if sheet.name.strip().lower() == "nf":
            ws_nf = sheet
            break

    if ws_nf is None:
        print("⚠️ Aba NF não encontrada — seguindo sem escrever NF")
        return  # NÃO quebra o programa

    ano = datetime.now().year

    texto = (
        f"SERVIÇO PRESTADO DE ATENDIMENTO/APOIO AO M/V {nome_navio}\n"
        f"DN {dn}/{ano}"
    )

    # escreve na primeira célula
    cel = ws_nf.range("A1")
    cel.value = texto

    # mescla para ficar bonito
    ws_nf.range("A1:E2").merge()

    # formatação
    cel.api.HorizontalAlignment = -4108  # center
    cel.api.VerticalAlignment = -4108
    cel.api.WrapText = True
    cel.api.Font.Name = "Calibri"
    cel.api.Font.Size = 14
    cel.api.Font.Bold = True

    print("✅ Texto da NF escrito com sucesso")

def main():
    print("🚀 Iniciando execução...")

    # ========== 1 - Licença ========== #

    validar_licenca()

    # ========= 2 – Abrir arquivos =========
    app, wb1, wb2, ws1, ws_front = abrir_workbooks()
    if not all([app, wb1, wb2, ws1, ws_front]):
        sys.exit("❌ Erro ao abrir workbooks")

    print("📂 Workbooks abertos")

    # ========= 3 – DN e Navio =========
    dn = obter_dn_da_pasta(wb1.fullname)
    if not dn:
        sys.exit("❌ DN não identificada pela pasta")

    nome_navio = obter_nome_navio_da_pasta(wb1.fullname)
    ano_atual = datetime.now().year
    texto_dn = f"DN: {dn}/{ano_atual}"

    # FRONT VIGIA
    ws_front.range("D15").value = nome_navio
    ws_front.range("C21").value = texto_dn

    berco = input("WAREHOUSE / BERÇO: ").strip().upper()
    ws_front["D18"].value = berco

    # ========= 4 – FRONT (OBRIGATÓRIO PRIMEIRO) =========
    print("⚙️ Processando FRONT VIGIA...")
    data_inicio, data_fim = processar_front(ws1, ws_front)
    if not data_inicio or not data_fim:
        sys.exit("❌ Datas extremas inválidas no RESUMO")
    print(f"📆 Datas extremas: {data_inicio} → {data_fim}")

    # ========= 5 – MMO =========
    print("⚙️ Processando MMO...")
    MMO(wb1.fullname, wb2)

    # ========= 6 – NF =========
    escrever_nf(wb2, nome_navio, dn)

    # ===== 7 – REPORT VIGIA =====
    ws_resumo = wb1.sheets["Resumo"]
    ws_report = wb2.sheets["REPORT VIGIA"]

    # 1️⃣ Obter quantidade de períodos
    periodos = obter_periodos(wb1.fullname)

    # 2️⃣ Inserir linhas extras se necessário
    inserir_linhas_report(ws_report, linha_inicial=22, periodos=periodos)

    # 3️⃣ Gerar lista de ciclos confiável (coluna E)
    ciclos_linha = gerar_coluna_E_ajustada(ws1, periodos, coluna_horario="C")

    # 4️⃣ Preencher coluna E do REPORT VIGIA
    preencher_coluna_E_por_ciclos(ws_report, ciclos_linha, linha_inicial=22)


    # 5️⃣ Mapear valores do wb1 por ciclo (coluna G)
    valores_por_ciclo = mapear_valores_por_ciclo(ws1, coluna_horario="C", coluna_valor="Z")

    # 6️⃣ Preencher coluna G do REPORT VIGIA
    preencher_coluna_G_por_ciclo(ws_report, ciclos_linha, valores_por_ciclo, coluna="G", linha_inicial=22)

    # 7️⃣ Preencher coluna C (datas) respeitando ciclos 00x06
    montar_datas_report_vigia(
        ws_report=ws_report,
        ws_resumo=ws_resumo,
    linha_inicial=22,
    periodos=periodos
)



    # 7️⃣ Garantir que data_inicio seja datetime
    if not isinstance(data_inicio, datetime):
        try:
            data_inicio = pd.to_datetime(data_inicio)
        except Exception as e:
            raise ValueError(f"data_inicio inválida: {data_inicio}") from e


    # ========= 8 – Financeiro =========
    OC(str(wb1.fullname), wb2)
    credit_note(wb2, texto_dn)
    quitacao(wb2, texto_dn)

    # ========= 10 – Ajustes finais =========
    arredondar_para_baixo_50(ws_front)
    cargonave(ws_front)

    pasta_saida = Path(wb1.fullname).parent
    arquivo_saida = pasta_saida / "3.xlsx"

    fechar_workbooks(app, wb1, wb2, arquivo_saida)

    print(f"✅ Processo finalizado: {arquivo_saida}")


if __name__ == "__main__":
    
    main()

# Fim do código
