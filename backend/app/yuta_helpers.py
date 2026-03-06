# ==============================
# IMPORTS
# ==============================
import os
import re
import ssl
import sys
import tempfile
import unicodedata
import urllib.request

import certifi
import holidays

import pdfplumber
import pytesseract
import xlwings as xw
from datetime import date, datetime, timedelta, timezone
from pathlib import Path
import shutil
from tempfile import gettempdir
from tkinter import Tk, filedialog

import comtypes.client
import openpyxl
from docx import Document
from num2words import num2words
from pdf2image import convert_from_path

from config_manager import obter_caminho_base_faturamentos



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
def _tentar_forcar_download_onedrive(caminho: Path) -> bool:
    """
    Tenta forçar o download de um arquivo OneDrive que pode estar apenas na nuvem.
    Retorna True se conseguiu acessar o arquivo, False caso contrário.
    """
    try:
        # Método 1: Tenta abrir para forçar download
        print(f"   Tentando forçar download: {caminho.name}")
        if caminho.exists():
            with open(caminho, 'rb') as f:
                f.read(1024)  # Lê 1KB para garantir
            print(f"   ✅ Arquivo acessível")
            return True
        else:
            # Método 2: Usa attrib do Windows para forçar download
            print(f"   ⚠️ Arquivo não disponível localmente, tentando forçar download...")
            import subprocess
            try:
                # Remove atributo P (pinned/unpinned) para forçar disponibilidade
                subprocess.run(
                    ['attrib', '-U', str(caminho)],
                    capture_output=True,
                    timeout=10,
                    check=False
                )
                # Tenta acessar novamente
                import time
                time.sleep(1)
                if caminho.exists():
                    print(f"   ✅ Arquivo baixado com sucesso")
                    return True
            except Exception as e:
                print(f"   ⚠️ Não foi possível forçar download: {e}")

            return False
    except (FileNotFoundError, OSError) as e:
        print(f"   ❌ Erro ao acessar: {e}")
        return False

def copiar_para_temp_xlwings(caminho_original: Path) -> Path:
    # Primeiro, tenta forçar download se for OneDrive
    if "OneDrive" in str(caminho_original) or "SANPORT" in str(caminho_original):
        print(f"🔄 Verificando sincronização OneDrive...")
        _tentar_forcar_download_onedrive(caminho_original)
    
    print(f"🔍 Procurando arquivo: {caminho_original.name}")
    print(f"🔍 Caminho completo: {caminho_original}")
    
    if not caminho_original.exists():
        # Tenta encontrar arquivo com nome similar (problema de codificação)
        pasta_pai = caminho_original.parent
        nome_procurado = caminho_original.name
        stem_procurado = caminho_original.stem

        def _norm_nome(s: str) -> str:
            s = unicodedata.normalize("NFKD", str(s))
            s = s.encode("ASCII", "ignore").decode("ASCII")
            s = s.replace("_", " ").replace("-", " ")
            s = re.sub(r"\s+", " ", s).strip().lower()
            return s
        
        print(f"⚠️ Arquivo não encontrado com nome exato")
        print(f"🔍 Arquivos .xlsx na pasta (como Python vê):")
        
        encontrado = None
        candidatos_xlsx = []
        if pasta_pai.exists():
            for item in pasta_pai.iterdir():
                if item.is_file() and item.suffix == '.xlsx':
                    print(f"   - {item.name}")
                    candidatos_xlsx.append(item)

            # 1) Match exato por nome normalizado (mais seguro)
            alvo_norm = _norm_nome(stem_procurado)
            for item in candidatos_xlsx:
                if _norm_nome(item.stem) == alvo_norm:
                    encontrado = item
                    break

            # 2) Fallback: todos os tokens do alvo presentes no candidato
            if not encontrado:
                tokens_alvo = [t for t in alvo_norm.split(" ") if t]
                candidatos_token = [
                    item for item in candidatos_xlsx
                    if all(t in _norm_nome(item.stem) for t in tokens_alvo)
                ]
                if candidatos_token:
                    # escolhe o nome mais próximo em tamanho do solicitado
                    encontrado = min(
                        candidatos_token,
                        key=lambda item: abs(len(_norm_nome(item.stem)) - len(alvo_norm))
                    )

            if encontrado:
                print(f"   ✅ Arquivo correspondente encontrado: {encontrado.name}")
        
        if encontrado:
            caminho_original = encontrado
        elif not caminho_original.exists():
            raise FileNotFoundError(
                f"\n❌ Arquivo não encontrado: {nome_procurado}\n"
                f"📂 Caminho: {caminho_original}\n\n"
                "🔧 SOLUÇÃO:\n"
                "   O arquivo está apenas na nuvem do OneDrive.\n"
                "   Para resolver, faça um dos seguintes:\n\n"
                "   1. Abra o arquivo no Excel (clique duas vezes)\n"
                "   2. Aguarde o OneDrive baixar o arquivo\n"
                "   3. Feche o Excel e execute o processo novamente\n\n"
                "   OU\n\n"
                "   1. Clique com botão direito no arquivo\n"
                "   2. Selecione 'Sempre manter neste dispositivo'\n"
                "   3. Execute o processo novamente\n"
            )

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
    r"""
    Localiza a pasta FATURAMENTOS usando o sistema de configuração.
    Retorna a pasta com os modelos (ex: ...\Central de Documentos - 01. FATURAMENTOS\FATURAMENTOS)
    """
    print("\n=== BUSCANDO PASTA FATURAMENTOS AUTOMATICAMENTE ===")

    try:
        # Usa o sistema de configuração centralizado
        caminho_base = obter_caminho_base_faturamentos()
        # Os modelos ficam na subpasta FATURAMENTOS dentro da pasta base
        caminho = caminho_base / "FATURAMENTOS"
        
        if not caminho.exists():
            # Fallback: se não existir a subpasta, usa a pasta base
            caminho = caminho_base
            
        print(f"✅ Pasta FATURAMENTOS encontrada em:\n   {caminho}")
        return caminho
    except FileNotFoundError:
        # Fallback: tenta o método antigo
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

    # 🔍 Busca flexível do arquivo do cliente na pasta de modelos
    pasta_modelos = _resolver_pasta_modelos_faturamento(pasta_faturamentos)
    caminho_cliente_rede = localizar_arquivo_cliente(pasta_modelos, nome_cliente)

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


def localizar_arquivo_cliente(pasta_faturamentos: Path, nome_cliente: str) -> Path:
    """
    Localiza o arquivo .xlsx do cliente na pasta de faturamentos.
    Faz busca flexível para encontrar mesmo com nomes diferentes.
    
    Ex: Se a pasta é "WILLIAMS", pode encontrar:
    - WILLIAMS.xlsx
    - WILLIAMS (PSS).xlsx
    - WILLIAMS - Porto.xlsx
    """
    import unicodedata
    
    def normalizar(texto: str) -> str:
        """Remove acentos, espaços, parênteses e caracteres especiais"""
        texto = texto.upper().strip()
        texto = unicodedata.normalize("NFKD", texto)
        texto = "".join(ch for ch in texto if not unicodedata.combining(ch))
        texto = re.sub(r"\([^)]*\)", "", texto)  # Remove parênteses
        texto = re.sub(r"[^A-Z0-9]+", "", texto)  # Remove não-alfanuméricos
        return texto
    
    # 1) Tentativa direta: nome exato
    caminho_direto = pasta_faturamentos / f"{nome_cliente}.xlsx"
    if caminho_direto.exists():
        return caminho_direto
    
    # 2) Busca flexível: normaliza e compara
    nome_normalizado = normalizar(nome_cliente)
    
    # Pega o nome base (sem parênteses)
    nome_base = re.sub(r"\s*\([^)]*\)\s*", " ", nome_cliente).strip()
    base_normalizado = normalizar(nome_base)
    
    candidatos = []
    
    for arquivo in pasta_faturamentos.glob("*.xlsx"):
        if arquivo.name.startswith("~"):  # Ignora arquivos temporários
            continue
        
        arquivo_norm = normalizar(arquivo.stem)
        
        # Match exato (normalizado)
        if arquivo_norm == nome_normalizado:
            candidatos.append(arquivo)
            continue
        
        # Match parcial (contém o nome base)
        if base_normalizado and base_normalizado in arquivo_norm:
            candidatos.append(arquivo)
    
    if len(candidatos) == 1:
        arquivo_encontrado = candidatos[0]
        if arquivo_encontrado.name != f"{nome_cliente}.xlsx":
            print(f"📎 Arquivo encontrado: {arquivo_encontrado.name} (cliente: {nome_cliente})")
        return arquivo_encontrado
    
    if len(candidatos) > 1:
        # Se há múltiplos candidatos, prefere o mais curto (mais específico)
        candidatos.sort(key=lambda p: len(p.name))
        arquivo_encontrado = candidatos[0]
        print(f"📎 Múltiplos arquivos encontrados, usando: {arquivo_encontrado.name}")
        return arquivo_encontrado
    
    # Não encontrado
    arquivos_disponiveis = [f.name for f in pasta_faturamentos.glob('*.xlsx') if not f.name.startswith('~')]
    raise FileNotFoundError(
        f"Arquivo de faturamento não encontrado para o cliente '{nome_cliente}'.\n"
        f"Pasta de faturamentos: {pasta_faturamentos}\n"
        f"Procurado: {nome_cliente}.xlsx\n"
        f"Arquivos disponíveis: {', '.join(arquivos_disponiveis[:10])}"
        + ("..." if len(arquivos_disponiveis) > 10 else "")
    )


def _resolver_pasta_modelos_faturamento(pasta_faturamentos: Path) -> Path:
    """
    Garante o caminho correto da pasta onde ficam os modelos dos clientes.
    """
    if pasta_faturamentos.name.upper() == "FATURAMENTOS":
        return pasta_faturamentos

    subpasta = pasta_faturamentos / "FATURAMENTOS"
    if subpasta.exists():
        return subpasta

    return pasta_faturamentos


def abrir_workbooks(pasta_faturamentos: Path, caminho_navio_rede: Path | str | None = None):
    if not caminho_navio_rede:
        caminho_navio_rede = selecionar_arquivo_navio()
        if not caminho_navio_rede:
            raise FileNotFoundError("Arquivo do NAVIO não selecionado")

    caminho_navio_rede = Path(caminho_navio_rede)
    pasta_navio = caminho_navio_rede.parent
    pasta_cliente = pasta_navio.parent
    nome_cliente = pasta_cliente.name.strip()

    # 🔍 Busca flexível do arquivo do cliente na pasta de modelos
    pasta_modelos = _resolver_pasta_modelos_faturamento(pasta_faturamentos)
    caminho_cliente_rede = localizar_arquivo_cliente(pasta_modelos, nome_cliente)

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

        # ✅ RETURN PADRONIZADO
        return app, wb1, wb2, ws1, ws_front, pasta_navio, caminho_navio_rede


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

def obter_modelo_word_cargonave(pasta_faturamentos: Path, cliente: str = "CARGONAVE") -> Path:
    """
    Localiza o modelo de recibo CARGONAVE com busca flexível.
    Suporta .doc e .docx e tenta pastas alternativas para rodar em PCs diferentes.
    """
    caminhos_teste = [
        pasta_faturamentos / cliente,
        pasta_faturamentos / "CARGONAVE",
        pasta_faturamentos,
        pasta_faturamentos.parent / "CARGONAVE",
    ]

    padroes = [
        "RECIBO - YUTA.doc",
        "RECIBO - YUTA.docx",
        "*RECIBO*YUTA*.doc",
        "*RECIBO*YUTA*.docx",
    ]

    vistos = set()
    candidatos = []

    for caminho in caminhos_teste:
        if not caminho.exists() or not caminho.is_dir():
            continue

        chave = str(caminho.resolve()).upper()
        if chave in vistos:
            continue
        vistos.add(chave)

        for padrao in padroes:
            candidatos.extend(caminho.glob(padrao))

        # fallback recursivo (limitado ao necessário)
        if not candidatos:
            for padrao in padroes:
                candidatos.extend(caminho.rglob(padrao))

        if candidatos:
            # prefere correspondência exata de nome
            exatos = [
                p for p in candidatos
                if p.name.strip().upper() in {"RECIBO - YUTA.DOC", "RECIBO - YUTA.DOCX"}
            ]
            alvo = sorted(exatos or candidatos, key=lambda p: p.stat().st_mtime, reverse=True)[0]
            print(f"✅ Modelo Word encontrado: {alvo}")
            return alvo

    raise FileNotFoundError(
        "Modelo Word não encontrado. Pastas verificadas: "
        + " | ".join(str(p) for p in caminhos_teste)
    )


def gerar_pdf(caminho_excel, pasta_saida, nome_base, ws=None):
    app = xw.App(visible=False, add_book=False)
    wb = app.books.open(str(caminho_excel))

    try:
        caminho_pdf = pasta_saida / f"{nome_base}.pdf"

        if ws is not None:
            try:
                ajustar_layout_pdf_por_aba(ws)
            except Exception:
                pass
            ws.api.ExportAsFixedFormat(Type=0, Filename=str(caminho_pdf))
        else:
            try:
                ajustar_layout_todas_abas_visiveis_no_wb(wb)
            except Exception:
                pass
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

    try:
        ajustar_layout_todas_abas_visiveis_no_wb(wb)
    except Exception:
        pass

    wb.api.ExportAsFixedFormat(
        Type=0,  # PDF
        Filename=str(caminho_pdf),
        Quality=0,
        IncludeDocProperties=True,
        IgnorePrintAreas=False,  # respeita área de impressão de cada aba
        OpenAfterPublish=False
    )

    return caminho_pdf


def ajustar_layout_report_vigia(ws_report):
    """
    Padroniza o layout de impressão da aba REPORT VIGIA para evitar corte no PDF.
    """
    xlPortrait = 1
    xlPaperA4 = 9

    try:
        ultima_linha, ultima_coluna = _detectar_area_util_planilha(
            ws_report,
            min_linhas=40,
            min_colunas=14,
            max_linhas_scan=260,
            max_colunas_scan=40,
        )

        area = ws_report.api.Range(
            ws_report.api.Cells(1, 1),
            ws_report.api.Cells(ultima_linha, ultima_coluna),
        )

        page_setup = ws_report.api.PageSetup
        page_setup.Zoom = False
        page_setup.FitToPagesWide = 1
        page_setup.FitToPagesTall = False
        page_setup.Orientation = xlPortrait
        page_setup.PaperSize = xlPaperA4
        page_setup.CenterHorizontally = True
        page_setup.CenterVertically = False
        page_setup.PrintArea = area.Address
    except Exception as e:
        print(f"⚠️ Não foi possível ajustar layout do REPORT VIGIA: {e}")


def ajustar_layout_report_vigia_no_wb(wb):
    for ws in wb.sheets:
        if ws.name.strip().upper() == "REPORT VIGIA":
            ajustar_layout_report_vigia(ws)
            break


def _detectar_area_util_planilha(
    ws,
    min_linhas=40,
    min_colunas=8,
    max_linhas_scan=220,
    max_colunas_scan=40,
):
    """
    Detecta a área útil por conteúdo real (evita UsedRange inflado por formatação).
    """
    try:
        valores = ws.range((1, 1), (max_linhas_scan, max_colunas_scan)).value
        if not isinstance(valores, list):
            valores = [[valores]]

        ultima_linha = 1
        ultima_coluna = 1

        for i, linha in enumerate(valores, start=1):
            if not isinstance(linha, list):
                linha = [linha]
            for j, valor in enumerate(linha, start=1):
                if valor not in (None, ""):
                    ultima_linha = max(ultima_linha, i)
                    ultima_coluna = max(ultima_coluna, j)

        # pequeno padding para não cortar bordas/rodapé por 1-2 células
        ultima_linha = min(ultima_linha + 2, max_linhas_scan)
        ultima_coluna = min(ultima_coluna + 1, max_colunas_scan)

        return max(ultima_linha, min_linhas), max(ultima_coluna, min_colunas)
    except Exception:
        return min_linhas, min_colunas


def ajustar_layout_front_vigia(ws_front):
    """
    Padroniza a impressão da FRONT VIGIA para reduzir variação entre máquinas/usuários.
    """
    xlPortrait = 1
    xlPaperA4 = 9

    try:
        ultima_linha, ultima_coluna = _detectar_area_util_planilha(
            ws_front,
            min_linhas=45,
            min_colunas=14,
            max_linhas_scan=180,
            max_colunas_scan=30,
        )

        area = ws_front.api.Range(
            ws_front.api.Cells(1, 1),
            ws_front.api.Cells(ultima_linha, ultima_coluna),
        )

        page_setup = ws_front.api.PageSetup
        page_setup.Zoom = False
        page_setup.FitToPagesWide = 1
        page_setup.FitToPagesTall = False
        page_setup.Orientation = xlPortrait
        page_setup.PaperSize = xlPaperA4
        page_setup.CenterHorizontally = True
        page_setup.CenterVertically = False
        page_setup.PrintArea = area.Address
    except Exception as e:
        print(f"⚠️ Não foi possível ajustar layout da FRONT VIGIA: {e}")


def ajustar_layout_front_vigia_no_wb(wb):
    for ws in wb.sheets:
        if ws.name.strip().upper() == "FRONT VIGIA":
            ajustar_layout_front_vigia(ws)
            break


def ajustar_layout_planilha_generica(
    ws,
    min_linhas=40,
    min_colunas=14,
    max_linhas_scan=260,
    max_colunas_scan=40,
):
    xlPortrait = 1
    xlPaperA4 = 9

    try:
        ultima_linha, ultima_coluna = _detectar_area_util_planilha(
            ws,
            min_linhas=min_linhas,
            min_colunas=min_colunas,
            max_linhas_scan=max_linhas_scan,
            max_colunas_scan=max_colunas_scan,
        )

        area = ws.api.Range(
            ws.api.Cells(1, 1),
            ws.api.Cells(ultima_linha, ultima_coluna),
        )

        page_setup = ws.api.PageSetup
        page_setup.Zoom = False
        page_setup.FitToPagesWide = 1
        page_setup.FitToPagesTall = False
        page_setup.Orientation = xlPortrait
        page_setup.PaperSize = xlPaperA4
        page_setup.CenterHorizontally = True
        page_setup.CenterVertically = False
        page_setup.PrintArea = area.Address
    except Exception as e:
        print(f"⚠️ Não foi possível ajustar layout da aba '{ws.name}': {e}")


def _normalizar_nome_aba_layout(nome: str) -> str:
    texto = unicodedata.normalize("NFKD", str(nome or ""))
    texto = texto.encode("ASCII", "ignore").decode("ASCII")
    texto = re.sub(r"\s+", " ", texto).strip().upper()
    return texto


def ajustar_layout_pdf_por_aba(ws):
    nome = _normalizar_nome_aba_layout(getattr(ws, "name", ""))

    if "REPORT VIGIA" in nome:
        ajustar_layout_report_vigia(ws)
        return

    if "FRONT VIGIA" in nome:
        ajustar_layout_front_vigia(ws)
        return

    ajustar_layout_planilha_generica(ws)


def ajustar_layout_todas_abas_visiveis_no_wb(wb, ignorar_abas=()):
    ignorar_norm = {
        _normalizar_nome_aba_layout(nome)
        for nome in (ignorar_abas or ())
    }

    for ws in wb.sheets:
        try:
            if not bool(ws.api.Visible):
                continue
        except Exception:
            pass

        nome_norm = _normalizar_nome_aba_layout(getattr(ws, "name", ""))
        if nome_norm in ignorar_norm:
            continue

        ajustar_layout_pdf_por_aba(ws)


def ajustar_layout_abas_estrategicas_no_wb(wb):
    ajustar_layout_todas_abas_visiveis_no_wb(wb)


def gerar_pdf_faturamento_completo(wb, pasta_saida: Path, nome_base: str, apenas_front=False) -> Path:
    caminho_pdf = pasta_saida / f"{nome_base}.pdf"

    if caminho_pdf.exists():
        caminho_pdf.unlink()

    # ✅ Se apenas_front=True, exporta só FRONT VIGIA
    if apenas_front:
        # Encontra a aba FRONT VIGIA
        ws_front = None
        for ws in wb.sheets:
            if ws.name.strip().upper() == "FRONT VIGIA":
                ws_front = ws
                break
        
        if ws_front:
            ajustar_layout_pdf_por_aba(ws_front)
            # Exporta apenas essa aba
            ws_front.api.ExportAsFixedFormat(
                Type=0,  # PDF
                Filename=str(caminho_pdf),
                Quality=0,
                IncludeDocProperties=True,
                IgnorePrintAreas=False,
                OpenAfterPublish=False
            )
            return caminho_pdf
        else:
            raise RuntimeError("Aba FRONT VIGIA não encontrada para exportar PDF")

    # 🔒 Oculta aba NF (se existir)
    aba_nf = None
    for ws in wb.sheets:
        if ws.name.strip().upper() == "NF":
            aba_nf = ws
            ws.api.Visible = False
            break

    ajustar_layout_todas_abas_visiveis_no_wb(wb, ignorar_abas=("NF",))

    # 📄 Exporta workbook inteiro
    wb.api.ExportAsFixedFormat(
        Type=0,  # PDF
        Filename=str(caminho_pdf),
        Quality=0,
        IncludeDocProperties=True,
        IgnorePrintAreas=False,
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


def gerar_pdf_do_wb_aberto(wb, pasta_saida, nome_base, ignorar_abas=("nf",), apenas_front=False):
    caminho_pdf = Path(pasta_saida) / f"{nome_base}.pdf"

    # 1) se existir e estiver aberto, já avisa o motivo
    if caminho_pdf.exists():
        try:
            caminho_pdf.unlink()
        except Exception as e:
            raise RuntimeError(f"PDF está aberto/travado e não pode ser sobrescrito: {caminho_pdf}") from e

    app = wb.app
    app.api.DisplayAlerts = False

    # ✅ Se apenas_front=True, exporta só FRONT VIGIA
    if apenas_front:
        # Encontra a aba FRONT VIGIA
        ws_front = None
        for sh in wb.sheets:
            if sh.name.strip().upper() == "FRONT VIGIA":
                ws_front = sh
                break
        
        if ws_front:
            ajustar_layout_pdf_por_aba(ws_front)
            # Ativa e exporta apenas essa aba
            ws_front.activate()
            ws_front.api.ExportAsFixedFormat(
                Type=0,  # xlTypePDF
                Filename=str(caminho_pdf),
                Quality=0,  # xlQualityStandard
                IncludeDocProperties=True,
                IgnorePrintAreas=False,
                OpenAfterPublish=False
            )
            print(f"📄 PDF gerado (apenas FRONT VIGIA): {caminho_pdf}")
            return caminho_pdf
        else:
            raise RuntimeError("Aba FRONT VIGIA não encontrada para exportar PDF")

    # 2) guarda visibilidade, oculta as que não devem sair no PDF
    vis_orig = {}
    for sh in wb.sheets:
        nome_norm = sh.name.strip().lower()
        vis_orig[sh.name] = sh.api.Visible
        if nome_norm in {x.strip().lower() for x in ignorar_abas}:
            sh.api.Visible = False  # oculta NF

    ajustar_layout_todas_abas_visiveis_no_wb(wb, ignorar_abas=ignorar_abas)

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

    # 🔥 define uma data fixa de expiração: 30 de março de 2026
    limite = datetime(2026, 3, 30, tzinfo=timezone.utc)

    if hoje_utc > limite:
        sys.exit("⛔ Licença expirada")

    print(f"📅 Data local: {hoje_local.date()}")
