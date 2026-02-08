# ==============================
# IMPORTS
# ==============================
import sys
import re
import ssl
import certifi
import urllib.request
import shutil
import tempfile
import pdfplumber
import os
import msvcrt

from pathlib import Path
from itertools import cycle
from datetime import datetime, date, timedelta, timezone

import tkinter as tk
from tkinter import Tk, filedialog

import pandas as pd
import xlwings as xw
import openpyxl
from copy import copy  # para copiar estilos
from openpyxl.styles import Font
from tempfile import gettempdir
import shutil
import holidays
from docx import Document
from num2words import num2words
import comtypes.client
import unicodedata
import locale
import unicodedata

from docx import Document
from shutil import copy2
from num2words import num2words
from datetime import datetime
import calendar


from pdf2image import convert_from_path
import pytesseract



# ==============================
# FERIADOS
# ==============================
feriados_br = holidays.Brazil()

feriados_personalizados = [
    date(2025, 1, 1),
    date(2025, 4, 21),
    date(2025, 5, 1),
    # ... outros feriados locais
]

for d in feriados_personalizados:
    feriados_br[d] = "Feriado personalizado"


# ==============================
# FUNÇÕES AUXILIARES GLOBAIS
# ==============================

# ---------------------------
# 1️⃣ Copiar arquivo para pasta temporária e ler Excel
# ---------------------------
def copiar_para_temp_xlwings(caminho_original: Path) -> Path:
    if not caminho_original.exists():
        raise FileNotFoundError(f"Arquivo não encontrado: {caminho_original}")

    temp_dir = Path(tempfile.mkdtemp(prefix="faturamento_"))
    caminho_temp = temp_dir / caminho_original.name

    print(f"📄 Copiando para local temporário:")
    print(f"   {caminho_original.name}")
    shutil.copy2(caminho_original, caminho_temp)

    return caminho_temp



def copiar_para_temp_word(caminho_original: Path) -> Path:
    if not caminho_original.exists():
        raise FileNotFoundError(f"Arquivo Word não encontrado: {caminho_original}")

    temp_dir = Path(tempfile.mkdtemp(prefix="recibo_"))
    caminho_temp = temp_dir / caminho_original.name

    print(f"📄 Copiando modelo Word para temporário:")
    print(f"   {caminho_original.name}")
    shutil.copy2(caminho_original, caminho_temp)

    return caminho_temp



# ---------------------------
# 2️⃣ Localizar pasta FATURAMENTOS automaticamente
# ---------------------------
def obter_pasta_faturamentos() -> Path:
    print("\n=== BUSCANDO PASTA FATURAMENTOS AUTOMATICAMENTE ===")

    bases = [
        Path.home() / "SANPORT LOGÍSTICA PORTUÁRIA LTDA",
        Path.home() / "OneDrive - SANPORT LOGÍSTICA PORTUÁRIA LTDA",
    ]

    for base in bases:
        if base.exists():
            candidatos = list(base.rglob("FATURAMENTOS"))
            for c in candidatos:
                if "01. FATURAMENTOS" in c.parent.as_posix():
                    print(f"✅ Pasta FATURAMENTOS encontrada em:\n   {c}")
                    return c

    raise FileNotFoundError("Pasta FATURAMENTOS não localizada")



# ---------------------------
# 3️⃣ Abrir workbooks NAVIO e cliente com xlwings

# ---------------------------


#================DE ACORDO====================#

def abrir_workbooks_de_acordo(pasta_faturamentos: Path, pasta_navio: Path):
    pasta_cliente = pasta_navio.parent
    nome_cliente = pasta_cliente.name.strip()

    caminho_cliente_rede = pasta_faturamentos / f"{nome_cliente}.xlsx"
    if not caminho_cliente_rede.exists():
        raise FileNotFoundError(f"Faturamento não encontrado: {caminho_cliente_rede}")

    caminho_cliente_local = copiar_para_temp_xlwings(caminho_cliente_rede)

    app = xw.App(visible=False, add_book=False)
    wb = None

    try:
        wb = app.books.open(str(caminho_cliente_local))

        nomes_abas = [s.name for s in wb.sheets]
        if nome_cliente in nomes_abas:
            ws_front = wb.sheets[nome_cliente]
        elif "FRONT VIGIA" in nomes_abas:
            ws_front = wb.sheets["FRONT VIGIA"]
        else:
            raise RuntimeError("Aba FRONT não encontrada")

        return app, wb, ws_front

    except Exception:
        if wb:
            wb.close()
        app.quit()
        raise



def montar_nome_faturamento(dn: str, nome_navio: str) -> str:
    """
    Ex: dn=1, nome_navio='SANPORT'
    -> 'FATURAMENTO - ND 001 - MV SANPORT'
    """
    nd_formatado = str(dn).zfill(3)
    return f"FATURAMENTO - ND {nd_formatado} - MV {nome_navio}"


def escrever_de_acordo_nf(wb, nome_navio, dn, ano):
    """
    Escreve o texto DE ACORDO na aba NF (A1:E2).
    """

    ws_nf = None
    for sheet in wb.sheets:
        if sheet.name.strip().lower() == "nf":
            ws_nf = sheet
            break

    if ws_nf is None:
        print("⚠️ Aba NF não encontrada (DE ACORDO).")
        return

    texto = (
        f'SERVIÇO DE ATENDIMENTO/APOIO NO "DE ACORDO" '
        f'DA RAP DO {nome_navio} DN {dn}/{ano}'
    )


    rng = ws_nf.range("A1:E2")

    # segurança: desfaz merge anterior
    if rng.api.MergeCells:
        rng.api.UnMerge()

    rng.merge()
    rng.value = texto

    cel = ws_nf.range("A1")
    cel.api.HorizontalAlignment = -4108  # Center
    cel.api.VerticalAlignment = -4108
    cel.api.WrapText = True
    cel.api.Font.Bold = True
    cel.api.Font.Size = 14


def obter_nome_navio_da_pasta(pasta_navio: Path) -> str:
    """
    Ex: '054 - sanport' -> 'SANPORT'
    """
    nome = re.sub(r"^\s*\d+\s*[-–—]?\s*", "", pasta_navio.name).strip()
    return nome.upper() if nome else "NAVIO NÃO IDENTIFICADO"


#====================================================================================#



#===================SISTEMA=========================================#


def abrir_workbooks(pasta_faturamentos: Path):
    caminho_navio_rede = selecionar_arquivo_navio()
    if not caminho_navio_rede:
        raise FileNotFoundError("Arquivo do NAVIO não selecionado")

    caminho_navio_rede = Path(caminho_navio_rede)
    pasta_navio = caminho_navio_rede.parent
    pasta_cliente = pasta_navio.parent
    nome_cliente = pasta_cliente.name.strip()

    caminho_cliente_rede = pasta_faturamentos / f"{nome_cliente}.xlsx"
    if not caminho_cliente_rede.exists():
        raise FileNotFoundError(
            f"Arquivo de faturamento não encontrado:\n{caminho_cliente_rede}"
        )

    # 🔥 COPIA AMBOS PARA LOCAL
    caminho_navio_local = copiar_para_temp_xlwings(caminho_navio_rede)
    caminho_cliente_local = copiar_para_temp_xlwings(caminho_cliente_rede)

    app = xw.App(visible=False, add_book=False)
    wb1 = wb2 = None

    try:
        wb1 = app.books.open(str(caminho_navio_local))
        wb2 = app.books.open(str(caminho_cliente_local))

        ws1 = wb1.sheets[0]
        nomes_abas = [s.name for s in wb2.sheets]

        if nome_cliente in nomes_abas:
            ws_front = wb2.sheets[nome_cliente]
        elif "FRONT VIGIA" in nomes_abas:
            ws_front = wb2.sheets["FRONT VIGIA"]
        else:
            raise RuntimeError("Aba FRONT não encontrada")

        # ✅ RETURN PADRONIZADO (5 valores)
        return app, wb1, wb2, ws1, ws_front, pasta_navio


    except Exception:
        if wb1:
            wb1.close()
        if wb2:
            wb2.close()
        app.quit()
        raise


def selecionar_pasta_navio() -> Path:
    root = Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    pasta = filedialog.askdirectory(title="Selecione a pasta do NAVIO")

    root.destroy()

    if not pasta:
        raise RuntimeError("Nenhuma pasta de navio selecionada")

    pasta = Path(pasta)
    print(f"📁 Pasta do navio selecionada: {pasta.name}")
    return pasta


def obter_nome_navio(pasta_navio: Path, caminho_navio: Path | None = None) -> str:
    """
    Prioridade:
    1) Nome no arquivo
    2) Nome da pasta
    """
    if caminho_navio:
        nome_arquivo = obter_nome_navio_de_arquivo(caminho_navio)
        if nome_arquivo:
            return nome_arquivo

    return obter_nome_navio_da_pasta(pasta_navio)




def escrever_nf_faturamento_completo(wb_faturamento, nome_navio, dn, celula="A1", area_merge="A1:E10"):
    ws_nf = None
    for sheet in wb_faturamento.sheets:
        if sheet.name.strip().lower() == "nf":
            ws_nf = sheet
            break

    if ws_nf is None:
        print("⚠️ Aba NF não encontrada.")
        return False

    ano = datetime.now().strftime("%y")

    texto = f"SERVIÇO PRESTADO DE ATENDIMENTO/APOIO AO M/V {nome_navio}\nDN {dn}/{ano}"

    rng = ws_nf.range(area_merge)

    # ✅ desfaz merges com segurança (mesmo se a área tiver merge parcial)
    try:
        rng.api.UnMerge()
    except Exception:
        pass

    rng.merge()
    rng.value = texto

    cel = ws_nf.range(celula)
    cel.api.HorizontalAlignment = -4108  # xlCenter
    cel.api.VerticalAlignment = -4108    # xlCenter
    cel.api.WrapText = True
    cel.api.Font.Bold = True
    cel.api.Font.Size = 12

    print("✅ NF preenchida (A1:E10)")
    return True




def obter_dn_da_pasta(pasta_navio: Path) -> str:
    """
    Extrai o DN do início do nome da pasta.
    Ex: '054 - SANPORT' -> '054'
    """
    match = re.match(r"^\s*(\d+)", pasta_navio.name)
    if not match:
        print(
            f"⚠️ DN não encontrado no início da pasta "
            f"'{pasta_navio.name}', usando '0000'"
        )
        return "0000"

    return match.group(1)


def obter_nome_navio_de_arquivo(caminho_navio: Path) -> str:
    """
    Ex: 'FATURAMENTO - ND 001 - MV HOS REMINGTON.xlsx'
    -> 'MV HOS REMINGTON'
    """
    nome = re.sub(
        r"^.*?(?:DN|ND)\s*\d+\s*[-–—]?\s*",
        "",
        caminho_navio.stem,
        flags=re.IGNORECASE
    ).strip()

    return nome.upper() if nome else "NAVIO NÃO IDENTIFICADO"



def fechar_workbooks(app=None, wb_navio=None, wb_cliente=None, arquivo_saida: Path | None = None):
    try:
        if wb_navio and arquivo_saida:
            if arquivo_saida.exists():
                arquivo_saida.unlink()
            wb_navio.save(str(arquivo_saida))
            print(f"💾 Arquivo Excel salvo em: {arquivo_saida}")
    except Exception as e:
        print(f"⚠️ Erro ao salvar wb_navio: {e}")

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

    try:
        if app:
            app.quit()
    except Exception as e:
        print(f"⚠️ Erro ao fechar Excel: {e}")


def selecionar_arquivo_navio() -> str | None:
    root = Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    root.update_idletasks()

    caminho = filedialog.askopenfilename(
        title="Selecione o arquivo do NAVIO",
        filetypes=[("Arquivos Excel", "*.xlsx")]
    )

    root.destroy()

    if not caminho:
        return None

    print(f"📂 Arquivo NAVIO selecionado: {Path(caminho).name}")
    return caminho

def salvar_excel_com_nome(wb, pasta_saida: Path, nome_base: str) -> Path:
    """
    Salva SEM usar SaveAs (evita erro Excel COM).
    """
    caminho_final = pasta_saida / f"{nome_base}.xlsx"

    # 🧠 Se existir, apaga
    if caminho_final.exists():
        caminho_final.unlink()

    # 🔥 ESSENCIAL: SaveCopyAs (não SaveAs)
    wb.api.SaveCopyAs(str(caminho_final))

    return caminho_final

def obter_modelo_word_cargonave(pasta_faturamentos: Path, cliente: str) -> Path:
    caminhos_teste = [
        pasta_faturamentos / cliente,
        pasta_faturamentos / "CARGONAVE",  # fallback
    ]

    for caminho in caminhos_teste:
            arquivos = list(caminho.glob("RECIBO - YUTA.docx"))
            if arquivos:
                return arquivos[0]

    raise FileNotFoundError(f"Modelo Word não encontrado em {caminhos_teste}")


def gerar_pdf(caminho_excel, pasta_saida, nome_base, ws=None):
    app = xw.App(visible=False, add_book=False)
    wb = app.books.open(str(caminho_excel))

    try:
        caminho_pdf = pasta_saida / f"{nome_base}.pdf"

        if ws is not None:
            ws.api.ExportAsFixedFormat(Type=0, Filename=str(caminho_pdf))
        else:
            wb.api.ExportAsFixedFormat(Type=0, Filename=str(caminho_pdf))

        print(f"📄 PDF gerado: {caminho_pdf}")
        return caminho_pdf

    finally:
        wb.close()
        app.quit()





def gerar_pdf_workbook_inteiro(wb, pasta_saida: Path, nome_base: str) -> Path:
    caminho_pdf = pasta_saida / f"{nome_base}.pdf"

    if caminho_pdf.exists():
        caminho_pdf.unlink()

    wb.api.ExportAsFixedFormat(
        Type=0,  # PDF
        Filename=str(caminho_pdf),
        Quality=0,
        IncludeDocProperties=True,
        IgnorePrintAreas=False,  # respeita área de impressão de cada aba
        OpenAfterPublish=False
    )

    return caminho_pdf


def gerar_pdf_faturamento_completo(wb, pasta_saida: Path, nome_base: str) -> Path:
    caminho_pdf = pasta_saida / f"{nome_base}.pdf"

    if caminho_pdf.exists():
        caminho_pdf.unlink()

    # 🔒 Oculta aba NF (se existir)
    aba_nf = None
    for ws in wb.sheets:
        if ws.name.strip().upper() == "NF":
            aba_nf = ws
            ws.api.Visible = False
            break

    # 🔥 Remove qualquer Print_Area escondido
    try:
        for nome in list(wb.api.Names):
            if nome.Name.lower() == "print_area":
                nome.Delete()
    except:
        pass

    # 📄 Exporta workbook inteiro
    wb.api.ExportAsFixedFormat(
        Type=0,  # PDF
        Filename=str(caminho_pdf),
        Quality=0,
        IncludeDocProperties=True,
        IgnorePrintAreas=True,
        OpenAfterPublish=False
    )

    # 🔓 Reexibe NF
    if aba_nf:
        aba_nf.api.Visible = True

    return caminho_pdf



def extrair_identidade_navio(pasta_navio: Path) -> tuple[str, str]:
    """
    Retorna (dn, nome_navio) a partir da pasta do navio
    Ex: '123 - UNIMAR' -> ('123', 'UNIMAR')
    """
    dn = obter_dn_da_pasta(pasta_navio)
    nome_navio = obter_nome_navio_da_pasta(pasta_navio)
    return dn, nome_navio


#===================FATURAMENTO SÃO SEBASTIÃO=========================================#


def gerar_pdf_do_wb_aberto(wb, pasta_saida, nome_base, ignorar_abas=("nf",)):
    caminho_pdf = Path(pasta_saida) / f"{nome_base}.pdf"

    # 1) se existir e estiver aberto, já avisa o motivo
    if caminho_pdf.exists():
        try:
            caminho_pdf.unlink()
        except Exception as e:
            raise RuntimeError(f"PDF está aberto/travado e não pode ser sobrescrito: {caminho_pdf}") from e

    app = wb.app
    app.api.DisplayAlerts = False

    # 2) guarda visibilidade, oculta as que não devem sair no PDF
    vis_orig = {}
    for sh in wb.sheets:
        nome_norm = sh.name.strip().lower()
        vis_orig[sh.name] = sh.api.Visible
        if nome_norm in {x.strip().lower() for x in ignorar_abas}:
            sh.api.Visible = False  # oculta NF

    try:
        # 3) ativa uma aba visível (Excel odeia export sem sheet ativa)
        aba_ativa = None
        for sh in wb.sheets:
            if sh.api.Visible:  # True / -1
                aba_ativa = sh
                break
        if aba_ativa:
            aba_ativa.activate()

        # 4) exporta o workbook (sem as abas ocultas)
        wb.api.ExportAsFixedFormat(
            Type=0,  # xlTypePDF
            Filename=str(caminho_pdf),
            Quality=0,  # xlQualityStandard
            IncludeDocProperties=True,
            IgnorePrintAreas=False,
            OpenAfterPublish=False
        )

        print(f"📄 PDF gerado: {caminho_pdf}")
        return caminho_pdf

    finally:
        # 5) restaura visibilidade original
        for sh in wb.sheets:
            if sh.name in vis_orig:
                sh.api.Visible = vis_orig[sh.name]




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

    # 🔥 define uma data fixa de expiração: 5 de janeiro de 2026
    limite = datetime(2026, 2, 25, tzinfo=timezone.utc)

    if hoje_utc > limite:
        sys.exit("⛔ Licença expirada")

    print(f"📅 Data local: {hoje_local.date()}")
