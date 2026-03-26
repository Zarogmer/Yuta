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
from pypdf import PdfReader, PdfWriter, Transformation

from backend.app.config_manager import obter_caminho_base_faturamentos
from backend.app.utils.path_utils import poppler_paths_candidatos



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
# FUNÃ‡Ã•ES AUXILIARES GLOBAIS
# ==============================

# ---------------------------
# 1ï¸âƒ£ Copiar arquivo para pasta temporÃ¡ria e ler Excel
# ---------------------------
def _tentar_forcar_download_onedrive(caminho: Path) -> bool:
    """
    Tenta forÃ§ar o download de um arquivo OneDrive que pode estar apenas na nuvem.
    Retorna True se conseguiu acessar o arquivo, False caso contrÃ¡rio.
    """
    try:
        # MÃ©todo 1: Tenta abrir para forÃ§ar download
        print(f"   Tentando forÃ§ar download: {caminho.name}")
        if caminho.exists():
            with open(caminho, 'rb') as f:
                f.read(1024)  # LÃª 1KB para garantir
            print(f"   âœ… Arquivo acessÃ­vel")
            return True
        else:
            # MÃ©todo 2: Usa attrib do Windows para forÃ§ar download
            print(f"   âš ï¸ Arquivo nÃ£o disponÃ­vel localmente, tentando forÃ§ar download...")
            import subprocess
            try:
                # Remove atributo P (pinned/unpinned) para forÃ§ar disponibilidade
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
                    print(f"   âœ… Arquivo baixado com sucesso")
                    return True
            except Exception as e:
                print(f"   âš ï¸ NÃ£o foi possÃ­vel forÃ§ar download: {e}")

            return False
    except (FileNotFoundError, OSError) as e:
        print(f"   âŒ Erro ao acessar: {e}")
        return False

def copiar_para_temp_xlwings(caminho_original: Path) -> Path:
    # Primeiro, tenta forÃ§ar download se for OneDrive
    if "OneDrive" in str(caminho_original) or "SANPORT" in str(caminho_original):
        print(f"ðŸ”„ Verificando sincronizaÃ§Ã£o OneDrive...")
        _tentar_forcar_download_onedrive(caminho_original)
    
    print(f"ðŸ” Procurando arquivo: {caminho_original.name}")
    print(f"ðŸ” Caminho completo: {caminho_original}")
    
    if not caminho_original.exists():
        # Tenta encontrar arquivo com nome similar (problema de codificaÃ§Ã£o)
        pasta_pai = caminho_original.parent
        nome_procurado = caminho_original.name
        stem_procurado = caminho_original.stem

        def _norm_nome(s: str) -> str:
            s = unicodedata.normalize("NFKD", str(s))
            s = s.encode("ASCII", "ignore").decode("ASCII")
            s = s.replace("_", " ").replace("-", " ")
            s = re.sub(r"\s+", " ", s).strip().lower()
            return s
        
        print(f"âš ï¸ Arquivo nÃ£o encontrado com nome exato")
        print(f"ðŸ” Arquivos .xlsx na pasta (como Python vÃª):")
        
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
                    # escolhe o nome mais prÃ³ximo em tamanho do solicitado
                    encontrado = min(
                        candidatos_token,
                        key=lambda item: abs(len(_norm_nome(item.stem)) - len(alvo_norm))
                    )

            if encontrado:
                print(f"   âœ… Arquivo correspondente encontrado: {encontrado.name}")
        
        if encontrado:
            caminho_original = encontrado
        elif not caminho_original.exists():
            raise FileNotFoundError(
                f"\nâŒ Arquivo nÃ£o encontrado: {nome_procurado}\n"
                f"ðŸ“‚ Caminho: {caminho_original}\n\n"
                "ðŸ”§ SOLUÃ‡ÃƒO:\n"
                "   O arquivo estÃ¡ apenas na nuvem do OneDrive.\n"
                "   Para resolver, faÃ§a um dos seguintes:\n\n"
                "   1. Abra o arquivo no Excel (clique duas vezes)\n"
                "   2. Aguarde o OneDrive baixar o arquivo\n"
                "   3. Feche o Excel e execute o processo novamente\n\n"
                "   OU\n\n"
                "   1. Clique com botÃ£o direito no arquivo\n"
                "   2. Selecione 'Sempre manter neste dispositivo'\n"
                "   3. Execute o processo novamente\n"
            )

    temp_dir = Path(tempfile.mkdtemp(prefix="faturamento_"))
    caminho_temp = temp_dir / caminho_original.name

    print(f"ðŸ“„ Copiando para local temporÃ¡rio:")
    print(f"   {caminho_original.name}")
    shutil.copy2(caminho_original, caminho_temp)

    return caminho_temp



def copiar_para_temp_word(caminho_original: Path) -> Path:
    if not caminho_original.exists():
        raise FileNotFoundError(f"Arquivo Word nÃ£o encontrado: {caminho_original}")

    temp_dir = Path(tempfile.mkdtemp(prefix="recibo_"))
    caminho_temp = temp_dir / caminho_original.name

    print(f"ðŸ“„ Copiando modelo Word para temporÃ¡rio:")
    print(f"   {caminho_original.name}")
    shutil.copy2(caminho_original, caminho_temp)

    return caminho_temp



# ---------------------------
# 2ï¸âƒ£ Localizar pasta FATURAMENTOS automaticamente
# ---------------------------
def obter_pasta_faturamentos() -> Path:
    r"""
    Localiza a pasta FATURAMENTOS usando o sistema de configuraÃ§Ã£o.
    Retorna a pasta com os modelos (ex: ...\Central de Documentos - 01. FATURAMENTOS\FATURAMENTOS)
    """
    print("\n=== BUSCANDO PASTA FATURAMENTOS AUTOMATICAMENTE ===")

    try:
        # Usa o sistema de configuraÃ§Ã£o centralizado
        caminho_base = obter_caminho_base_faturamentos()
        # Os modelos ficam na subpasta FATURAMENTOS dentro da pasta base
        caminho = caminho_base / "FATURAMENTOS"
        
        if not caminho.exists():
            # Fallback: se nÃ£o existir a subpasta, usa a pasta base
            caminho = caminho_base
            
        print(f"âœ… Pasta FATURAMENTOS encontrada em:\n   {caminho}")
        return caminho
    except FileNotFoundError:
        # Fallback: tenta o mÃ©todo antigo
        bases = [
            Path.home() / "SANPORT LOGÃSTICA PORTUÃRIA LTDA",
            Path.home() / "OneDrive - SANPORT LOGÃSTICA PORTUÃRIA LTDA",
        ]

        for base in bases:
            if base.exists():
                candidatos = list(base.rglob("FATURAMENTOS"))
                for c in candidatos:
                    if "01. FATURAMENTOS" in c.parent.as_posix():
                        print(f"âœ… Pasta FATURAMENTOS encontrada em:\n   {c}")
                        return c

        raise FileNotFoundError("Pasta FATURAMENTOS nÃ£o localizada")



# ---------------------------
# 3ï¸âƒ£ Abrir workbooks NAVIO e cliente com xlwings

# ---------------------------


#================DE ACORDO====================#

def abrir_workbooks_de_acordo(pasta_faturamentos: Path, pasta_navio: Path):
    pasta_cliente = pasta_navio.parent
    nome_cliente = pasta_cliente.name.strip()

    # ðŸ” Busca flexÃ­vel do arquivo do cliente na pasta de modelos
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
            raise RuntimeError("Aba FRONT nÃ£o encontrada")

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
        print("âš ï¸ Aba NF nÃ£o encontrada (DE ACORDO).")
        return

    texto = (
        f'SERVIÇO DE ATENDIMENTO/APOIO NO "DE ACORDO" '
        f'DA RAP DO {nome_navio} DN {dn}/{ano}'
    )


    rng = ws_nf.range("A1:E2")

    # seguranÃ§a: desfaz merge anterior
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
    nome = re.sub(r"^\s*\d+\s*[-â€“â€”]?\s*", "", pasta_navio.name).strip()
    return nome.upper() if nome else "NAVIO NÃƒO IDENTIFICADO"


#====================================================================================#



#===================SISTEMA=========================================#


def localizar_arquivo_cliente(pasta_faturamentos: Path, nome_cliente: str) -> Path:
    """
    Localiza o arquivo .xlsx do cliente na pasta de faturamentos.
    Faz busca flexÃ­vel para encontrar mesmo com nomes diferentes.
    
    Ex: Se a pasta Ã© "WILLIAMS", pode encontrar:
    - WILLIAMS.xlsx
    - WILLIAMS (PSS).xlsx
    - WILLIAMS - Porto.xlsx
    """
    import unicodedata
    
    def normalizar(texto: str) -> str:
        """Remove acentos, espaÃ§os, parÃªnteses e caracteres especiais"""
        texto = texto.upper().strip()
        texto = unicodedata.normalize("NFKD", texto)
        texto = "".join(ch for ch in texto if not unicodedata.combining(ch))
        texto = re.sub(r"\([^)]*\)", "", texto)  # Remove parÃªnteses
        texto = re.sub(r"[^A-Z0-9]+", "", texto)  # Remove nÃ£o-alfanumÃ©ricos
        return texto
    
    # 1) Tentativa direta: nome exato
    caminho_direto = pasta_faturamentos / f"{nome_cliente}.xlsx"
    if caminho_direto.exists():
        return caminho_direto
    
    # 2) Busca flexÃ­vel: normaliza e compara
    nome_normalizado = normalizar(nome_cliente)
    
    # Pega o nome base (sem parÃªnteses)
    nome_base = re.sub(r"\s*\([^)]*\)\s*", " ", nome_cliente).strip()
    base_normalizado = normalizar(nome_base)
    
    candidatos = []
    
    for arquivo in pasta_faturamentos.glob("*.xlsx"):
        if arquivo.name.startswith("~"):  # Ignora arquivos temporÃ¡rios
            continue
        
        arquivo_norm = normalizar(arquivo.stem)
        
        # Match exato (normalizado)
        if arquivo_norm == nome_normalizado:
            candidatos.append(arquivo)
            continue
        
        # Match parcial (contÃ©m o nome base)
        if base_normalizado and base_normalizado in arquivo_norm:
            candidatos.append(arquivo)
    
    if len(candidatos) == 1:
        arquivo_encontrado = candidatos[0]
        if arquivo_encontrado.name != f"{nome_cliente}.xlsx":
            print(f"ðŸ“Ž Arquivo encontrado: {arquivo_encontrado.name} (cliente: {nome_cliente})")
        return arquivo_encontrado
    
    if len(candidatos) > 1:
        # Se hÃ¡ mÃºltiplos candidatos, prefere o mais curto (mais especÃ­fico)
        candidatos.sort(key=lambda p: len(p.name))
        arquivo_encontrado = candidatos[0]
        print(f"ðŸ“Ž MÃºltiplos arquivos encontrados, usando: {arquivo_encontrado.name}")
        return arquivo_encontrado
    
    # NÃ£o encontrado
    arquivos_disponiveis = [f.name for f in pasta_faturamentos.glob('*.xlsx') if not f.name.startswith('~')]
    raise FileNotFoundError(
        f"Arquivo de faturamento nÃ£o encontrado para o cliente '{nome_cliente}'.\n"
        f"Pasta de faturamentos: {pasta_faturamentos}\n"
        f"Procurado: {nome_cliente}.xlsx\n"
        f"Arquivos disponÃ­veis: {', '.join(arquivos_disponiveis[:10])}"
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
            raise FileNotFoundError("Arquivo do NAVIO nÃ£o selecionado")

    caminho_navio_rede = Path(caminho_navio_rede)
    pasta_navio = caminho_navio_rede.parent
    pasta_cliente = pasta_navio.parent
    nome_cliente = pasta_cliente.name.strip()

    # ðŸ” Busca flexÃ­vel do arquivo do cliente na pasta de modelos
    pasta_modelos = _resolver_pasta_modelos_faturamento(pasta_faturamentos)
    caminho_cliente_rede = localizar_arquivo_cliente(pasta_modelos, nome_cliente)

    # ðŸ”¥ COPIA AMBOS PARA LOCAL
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
            raise RuntimeError("Aba FRONT nÃ£o encontrada")

        # âœ… RETURN PADRONIZADO
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
    print(f"ðŸ“ Pasta do navio selecionada: {pasta.name}")
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
        print("âš ï¸ Aba NF nÃ£o encontrada.")
        return False

    ano = datetime.now().strftime("%y")

    # Evita problema de codificacao no template/Excel.
    texto = f"SERVICO PRESTADO DE ATENDIMENTO/APOIO AO M/V {nome_navio}\nDN {dn}/{ano}"

    rng = ws_nf.range(area_merge)

    # âœ… desfaz merges com seguranÃ§a (mesmo se a Ã¡rea tiver merge parcial)
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

    print("âœ… NF preenchida (A1:E10)")
    return True


def _compactar_print_area_para_exportacao_final(
    ws,
    pad_linhas=4,
    pad_colunas=1,
    min_linhas=35,
    min_colunas=8,
    max_linhas_scan=500,
    max_colunas_scan=40,
):
    """
    Reduz espaco em branco no PDF final, aparando apenas a sobra ao final da
    PrintArea atual. Nao altera preview e preserva a origem do template.
    """
    try:
        ps = ws.api.PageSetup
        atual = str(ps.PrintArea or "").strip()

        if atual:
            area = ws.api.Range(atual)
            row1 = int(area.Row)
            col1 = int(area.Column)
            row2 = int(area.Row + area.Rows.Count - 1)
            col2 = int(area.Column + area.Columns.Count - 1)
        else:
            row1 = 1
            col1 = 1
            info = detectar_area_util_planilha(
                ws,
                min_linhas=min_linhas,
                min_colunas=min_colunas,
                max_linhas_scan=max_linhas_scan,
                max_colunas_scan=max_colunas_scan,
            )
            row2 = int(info["last_row"])
            col2 = int(info["last_col"])

        scan_row2 = min(row2, row1 + max_linhas_scan - 1)
        scan_col2 = min(col2, col1 + max_colunas_scan - 1)

        valores = _normalize_2d(ws.range((row1, col1), (scan_row2, scan_col2)).value)
        formulas = _normalize_2d(ws.range((row1, col1), (scan_row2, scan_col2)).formula)

        ultimo_row_rel = 0
        ultimo_col_rel = 0
        total_rows = max(0, scan_row2 - row1 + 1)
        total_cols = max(0, scan_col2 - col1 + 1)

        for i in range(total_rows):
            linha_vals = valores[i] if i < len(valores) else []
            linha_for = formulas[i] if i < len(formulas) else []
            for j in range(total_cols):
                valor = linha_vals[j] if j < len(linha_vals) else None
                formula = linha_for[j] if j < len(linha_for) else None
                if _tem_conteudo_celula(valor, formula):
                    ultimo_row_rel = max(ultimo_row_rel, i + 1)
                    ultimo_col_rel = max(ultimo_col_rel, j + 1)

        if ultimo_row_rel <= 0 or ultimo_col_rel <= 0:
            return

        novo_row2 = min(row2, max(row1 + min_linhas - 1, row1 + ultimo_row_rel - 1 + pad_linhas))
        novo_col2 = min(col2, max(col1 + min_colunas - 1, col1 + ultimo_col_rel - 1 + pad_colunas))

        if novo_row2 >= row1 and novo_col2 >= col1:
            ps.PrintArea = ws.api.Range(
                ws.api.Cells(row1, col1),
                ws.api.Cells(novo_row2, novo_col2),
            ).Address
            _log_layout_debug(ws, "compactado_exportacao_final", {
                "last_row": novo_row2,
                "last_col": novo_col2,
                "scan_rows": total_rows,
                "scan_cols": total_cols,
            })
    except Exception as e:
        print(f"Nao foi possivel compactar PrintArea final da aba '{ws.name}': {e}")


def compactar_layout_final_ws(ws):
    nome_norm = _normalizar_nome_aba_layout(getattr(ws, "name", ""))
    if "REPORT VIGIA" in nome_norm:
        ajustar_layout_report_vigia_final(ws)
        return

    # FRONT VIGIA tem bordas em celulas vazias (moldura do template).
    # Compactar por valores/formulas corta essas bordas, causando PDF
    # sem moldura em PCs com driver de impressora diferente.
    if "FRONT VIGIA" in nome_norm:
        return

    _compactar_print_area_para_exportacao_final(ws)


def compactar_layout_final_wb(wb, ignorar_abas=()):
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

        compactar_layout_final_ws(ws)


def _cm_para_pontos(valor_cm: float) -> float:
    return float(valor_cm) * 72.0 / 2.54


def _bbox_conteudo_imagem_preview(image, pad=18):
    try:
        gray = image.convert("L")
        mask = gray.point(lambda p: 255 if p < 245 else 0)
        bbox = mask.getbbox()
        if not bbox:
            return None

        left, top, right, bottom = bbox
        left = max(0, left - pad)
        top = max(0, top - pad)
        right = min(image.width, right + pad)
        bottom = min(image.height, bottom + pad)
        if right <= left or bottom <= top:
            return None
        return (left, top, right, bottom)
    except Exception:
        return None


def _renderizar_paginas_pdf_para_bbox(caminho_pdf, dpi=160):
    erros = []

    try:
        paginas = convert_from_path(str(caminho_pdf), dpi=dpi)
        if paginas:
            return paginas
    except Exception as exc:
        erros.append(str(exc))

    for poppler_dir in poppler_paths_candidatos():
        try:
            paginas = convert_from_path(
                str(caminho_pdf),
                dpi=dpi,
                poppler_path=str(poppler_dir),
            )
            if paginas:
                return paginas
        except Exception as exc:
            erros.append(str(exc))

    if erros:
        print(f"Nao foi possivel renderizar PDF para recorte final: {erros[-1]}")
    return []


def normalizar_pdf_final_para_a4(caminho_pdf):
    """
    Reposiciona cada pagina gerada pelo Excel dentro de uma Letter limpa,
    usando o mesmo recorte visual do preview apenas para detectar o conteudo.
    O resultado continua sendo PDF; nao vira screenshot.
    Padrao: Letter (21.59 x 27.94 cm) com margens estreitas.
    """
    caminho_pdf = Path(caminho_pdf)
    if not caminho_pdf.exists():
        return False

    try:
        reader = PdfReader(str(caminho_pdf))
        if not reader.pages:
            return False

        paginas_img = _renderizar_paginas_pdf_para_bbox(caminho_pdf, dpi=160)
        if not paginas_img:
            return False

        writer = PdfWriter()
        # Letter: 21.59 x 27.94 cm
        letter_width = _cm_para_pontos(21.59)
        letter_height = _cm_para_pontos(27.94)
        margin_x = _cm_para_pontos(0.64)
        margin_y = _cm_para_pontos(1.10)
        usable_w = max(1.0, letter_width - (2 * margin_x))
        usable_h = max(1.0, letter_height - (2 * margin_y))

        for idx, page in enumerate(reader.pages):
            bbox_img = None
            if idx < len(paginas_img):
                bbox_img = _bbox_conteudo_imagem_preview(paginas_img[idx])

            if not bbox_img:
                writer.add_page(page)
                continue

            media_w = float(page.mediabox.width or 0)
            media_h = float(page.mediabox.height or 0)
            if media_w <= 0 or media_h <= 0:
                writer.add_page(page)
                continue

            img = paginas_img[idx]
            scale_x = media_w / float(img.width or 1)
            scale_y = media_h / float(img.height or 1)

            left_px, top_px, right_px, bottom_px = bbox_img
            x0 = max(0.0, left_px * scale_x)
            x1 = min(media_w, right_px * scale_x)
            y0 = max(0.0, media_h - (bottom_px * scale_y))
            y1 = min(media_h, media_h - (top_px * scale_y))

            src_w = max(1.0, x1 - x0)
            src_h = max(1.0, y1 - y0)
            scale = min(usable_w / src_w, usable_h / src_h)

            page.cropbox.lower_left = (x0, y0)
            page.cropbox.upper_right = (x1, y1)

            tx = margin_x + ((usable_w - (src_w * scale)) / 2.0) - (x0 * scale)
            ty = margin_y + ((usable_h - (src_h * scale)) / 2.0) - (y0 * scale)

            nova_pagina = writer.add_blank_page(width=letter_width, height=letter_height)
            nova_pagina.merge_transformed_page(
                page,
                Transformation().scale(scale, scale).translate(tx, ty),
                over=True,
            )

        caminho_tmp = caminho_pdf.with_suffix(".tmp.pdf")
        with open(caminho_tmp, "wb") as fp:
            writer.write(fp)

        caminho_tmp.replace(caminho_pdf)
        print(f"PDF final normalizado para Letter: {caminho_pdf}")
        return True
    except Exception as e:
        print(f"Nao foi possivel normalizar PDF final '{caminho_pdf}': {e}")
        return False


def _pagina_pdf_eh_report_vigia(page) -> bool:
    try:
        texto = (page.extract_text() or "").upper()
    except Exception:
        return False

    marcadores = (
        "COST OF WATCHMEN",
        "WATCHMEN",
        "GANGWAYMAN/WATCHMEN",
    )
    return any(marcador in texto for marcador in marcadores)


def _obter_dimensoes_papel_em_pontos(ws):
    app = ws.book.app
    ps = ws.api.PageSetup
    paper = int(getattr(ps, "PaperSize", 9) or 9)
    orient = int(getattr(ps, "Orientation", 1) or 1)
    largura_cm, altura_cm = _PAPER_DIMS_CM.get(paper, _PAPER_DIMS_CM[9])

    # xlLandscape = 2
    if orient == 2:
        largura_cm, altura_cm = altura_cm, largura_cm

    return (
        float(app.api.CentimetersToPoints(largura_cm)),
        float(app.api.CentimetersToPoints(altura_cm)),
    )


def ajustar_layout_report_vigia_final(ws_report):
    """
    Layout final do REPORT VIGIA preservando uma unica pagina e a borda completa.
    O preview permanece inalterado; esta regra vale so para o PDF final.
    """
    try:
        ajustar_layout_report_vigia(ws_report)

        ps = ws_report.api.PageSetup
        atual = str(ps.PrintArea or "").strip()
        if not atual:
            return

        area = ws_report.api.Range(atual)
        row1 = int(area.Row)
        col1 = int(area.Column)
        row2 = int(area.Row + area.Rows.Count - 1)
        col2 = int(area.Column + area.Columns.Count - 1)

        row2 = min(row2 + 2, row1 + 1199)
        col2 = min(max(col2, col1 + 8), col1 + 11)

        ps.PrintArea = ws_report.api.Range(
            ws_report.api.Cells(row1, col1),
            ws_report.api.Cells(row2, col2),
        ).Address
        ps.Zoom = False
        ps.FitToPagesWide = 1
        ps.FitToPagesTall = 1
        ps.CenterHorizontally = True
        ps.CenterVertically = False

        _log_layout_debug(ws_report, "report_vigia_final_preservado", {
            "last_row": row2,
            "last_col": col2,
            "scan_rows": row2 - row1 + 1,
            "scan_cols": col2 - col1 + 1,
        })
    except Exception as e:
        print(f"Nao foi possivel ajustar layout final do REPORT VIGIA: {e}")





def obter_dn_da_pasta(pasta_navio: Path) -> str:
    """
    Extrai o DN do inÃ­cio do nome da pasta.
    Ex: '054 - SANPORT' -> '054'
    """
    match = re.match(r"^\s*(\d+)", pasta_navio.name)
    if not match:
        print(
            f"âš ï¸ DN nÃ£o encontrado no inÃ­cio da pasta "
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
        r"^.*?(?:DN|ND)\s*\d+\s*[-â€“â€”]?\s*",
        "",
        caminho_navio.stem,
        flags=re.IGNORECASE
    ).strip()

    return nome.upper() if nome else "NAVIO NÃƒO IDENTIFICADO"



def fechar_workbooks(app=None, wb_navio=None, wb_cliente=None, arquivo_saida: Path | None = None):
    try:
        if wb_navio and arquivo_saida:
            if arquivo_saida.exists():
                arquivo_saida.unlink()
            wb_navio.save(str(arquivo_saida))
            print(f"ðŸ’¾ Arquivo Excel salvo em: {arquivo_saida}")
    except Exception as e:
        print(f"âš ï¸ Erro ao salvar wb_navio: {e}")

    try:
        if wb_navio:
            wb_navio.close()
    except Exception as e:
        print(f"âš ï¸ Erro ao fechar wb_navio: {e}")

    try:
        if wb_cliente:
            wb_cliente.close()
    except Exception as e:
        print(f"âš ï¸ Erro ao fechar wb_cliente: {e}")

    try:
        if app:
            app.quit()
    except Exception as e:
        print(f"âš ï¸ Erro ao fechar Excel: {e}")


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

    print(f"ðŸ“‚ Arquivo NAVIO selecionado: {Path(caminho).name}")
    return caminho

def salvar_excel_com_nome(wb, pasta_saida: Path, nome_base: str) -> Path:
    """
    Salva SEM usar SaveAs (evita erro Excel COM).
    """
    caminho_final = pasta_saida / f"{nome_base}.xlsx"

    # ðŸ§  Se existir, apaga
    if caminho_final.exists():
        caminho_final.unlink()

    # ðŸ”¥ ESSENCIAL: SaveCopyAs (nÃ£o SaveAs)
    wb.api.SaveCopyAs(str(caminho_final))

    return caminho_final

def obter_modelo_word_cargonave(pasta_faturamentos: Path, cliente: str = "CARGONAVE") -> Path:
    """
    Localiza o modelo de recibo CARGONAVE com busca flexÃ­vel.
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

        # fallback recursivo (limitado ao necessÃ¡rio)
        if not candidatos:
            for padrao in padroes:
                candidatos.extend(caminho.rglob(padrao))

        if candidatos:
            # prefere correspondÃªncia exata de nome
            exatos = [
                p for p in candidatos
                if p.name.strip().upper() in {"RECIBO - YUTA.DOC", "RECIBO - YUTA.DOCX"}
            ]
            alvo = sorted(exatos or candidatos, key=lambda p: p.stat().st_mtime, reverse=True)[0]
            print(f"âœ… Modelo Word encontrado: {alvo}")
            return alvo

    raise FileNotFoundError(
        "Modelo Word nÃ£o encontrado. Pastas verificadas: "
        + " | ".join(str(p) for p in caminhos_teste)
    )


def gerar_pdf(caminho_excel, pasta_saida, nome_base, ws=None, tipo_layout=None):
    app = xw.App(visible=False, add_book=False)
    wb = app.books.open(str(caminho_excel))

    try:
        caminho_pdf = pasta_saida / f"{nome_base}.pdf"

        if ws is not None:
            ws_export = None
            ws_alvo_nome = str(getattr(ws, "name", "")).strip().upper()
            for sh in wb.sheets:
                if str(sh.name).strip().upper() == ws_alvo_nome:
                    ws_export = sh
                    break

            if ws_export is None and len(wb.sheets) == 1:
                ws_export = wb.sheets[0]

            if ws_export is None:
                raise RuntimeError(
                    f"Aba alvo nao encontrada no workbook salvo: {getattr(ws, 'name', '<sem nome>')}"
                )

            try:
                aplicar_layout_pdf_especifico(ws_export, tipo_layout=tipo_layout)
            except Exception:
                pass

            try:
                compactar_layout_final_ws(ws_export)
            except Exception:
                pass

            ws_export.activate()
            ws_export.api.ExportAsFixedFormat(
                Type=0,
                Filename=str(caminho_pdf),
                Quality=0,
                IncludeDocProperties=True,
                IgnorePrintAreas=False,
                OpenAfterPublish=False,
            )
        else:
            try:
                ajustar_layout_todas_abas_visiveis_no_wb(wb)
            except Exception:
                pass

            try:
                compactar_layout_final_wb(wb)
            except Exception:
                pass

            wb.api.ExportAsFixedFormat(
                Type=0,
                Filename=str(caminho_pdf),
                Quality=0,
                IncludeDocProperties=True,
                IgnorePrintAreas=False,
                OpenAfterPublish=False,
            )

        try:
            normalizar_pdf_final_para_a4(caminho_pdf)
        except Exception:
            pass

        print(f"ðŸ“„ PDF gerado: {caminho_pdf}")
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

    try:
        compactar_layout_final_wb(wb)
    except Exception:
        pass

    wb.api.ExportAsFixedFormat(
        Type=0,  # PDF
        Filename=str(caminho_pdf),
        Quality=0,
        IncludeDocProperties=True,
        IgnorePrintAreas=False,  # respeita Ã¡rea de impressÃ£o de cada aba
        OpenAfterPublish=False
    )

    try:
        normalizar_pdf_final_para_a4(caminho_pdf)
    except Exception:
        pass

    return caminho_pdf


# Dimensoes de papel em cm  {xlPaperSize: (largura_cm, altura_cm)}
_PAPER_DIMS_CM = {
    1: (21.59, 27.94),   # Letter
    5: (21.59, 35.56),   # Legal
    7: (18.415, 26.67),  # Executive
    8: (29.7, 42.0),     # A3
    9: (21.0, 29.7),     # A4
    11: (21.0, 29.7),    # A4 Small
}


def _normalize_2d(values):
    if isinstance(values, list):
        if values and isinstance(values[0], list):
            return values
        return [values]
    return [[values]]


def _tem_conteudo_celula(valor, formula):
    if formula not in (None, ""):
        return True
    if valor is None:
        return False
    if isinstance(valor, str):
        return valor.strip() != ""
    return True


def detectar_area_util_planilha(
    ws,
    min_linhas=40,
    min_colunas=8,
    max_linhas_scan=400,
    max_colunas_scan=60,
):
    """
    Detecta a area com conteudo real (valor ou formula) sem confiar em UsedRange.
        if not reader.pages:
            return False
    Isso evita encolhimento do PDF por colunas/linhas fantasmas no PrintArea.
    """
    try:
        used = ws.api.UsedRange
        used_last_row = int(used.Row + used.Rows.Count - 1)
        used_last_col = int(used.Column + used.Columns.Count - 1)
    except Exception:
        used_last_row = min_linhas
        used_last_col = min_colunas

    scan_rows = max(min_linhas, min(max_linhas_scan, used_last_row + 8))
    scan_cols = max(min_colunas, min(max_colunas_scan, used_last_col + 4))

    valores = _normalize_2d(ws.range((1, 1), (scan_rows, scan_cols)).value)
    formulas = _normalize_2d(ws.range((1, 1), (scan_rows, scan_cols)).formula)

    ultima_linha = min_linhas
    ultima_coluna = min_colunas

    for i in range(scan_rows):
        linha_vals = valores[i] if i < len(valores) else []
        linha_for = formulas[i] if i < len(formulas) else []
        for j in range(scan_cols):
            valor = linha_vals[j] if j < len(linha_vals) else None
            formula = linha_for[j] if j < len(linha_for) else None
            if _tem_conteudo_celula(valor, formula):
                ultima_linha = max(ultima_linha, i + 1)
                ultima_coluna = max(ultima_coluna, j + 1)

    # Pequeno padding para evitar corte de borda/rodape
    ultima_linha = min(ultima_linha + 2, scan_rows)
    ultima_coluna = min(ultima_coluna + 1, scan_cols)

    return {
        "last_row": max(ultima_linha, min_linhas),
        "last_col": max(ultima_coluna, min_colunas),
        "scan_rows": scan_rows,
        "scan_cols": scan_cols,
        "used_last_row": used_last_row,
        "used_last_col": used_last_col,
    }


def limpar_planilha_para_exportacao(
    ws,
    min_linhas=40,
    min_colunas=8,
    max_linhas_scan=400,
    max_colunas_scan=60,
):
    """
    Limpa configuracoes antigas de impressao e redefine PrintArea pela area util.
    """
    info = detectar_area_util_planilha(
        ws,
        min_linhas=min_linhas,
        min_colunas=min_colunas,
        max_linhas_scan=max_linhas_scan,
        max_colunas_scan=max_colunas_scan,
    )

    area = ws.api.Range(
        ws.api.Cells(1, 1),
        ws.api.Cells(info["last_row"], info["last_col"]),
    )

    ps = ws.api.PageSetup
    try:
        ps.PrintArea = ""
    except Exception:
        pass

    ps.PrintArea = area.Address
    info["print_area"] = area.Address
    return info


def _log_layout_debug(ws, etapa, info=None):
    try:
        ps = ws.api.PageSetup
        paper = int(ps.PaperSize)
        fit_w = ps.FitToPagesWide
        fit_h = ps.FitToPagesTall
        zoom = ps.Zoom
        print(
            f"[PDF-DEBUG] etapa={etapa} aba='{ws.name}' paper={paper} "
            f"print_area='{ps.PrintArea}' zoom={zoom} fitW={fit_w} fitH={fit_h} "
            f"centerH={ps.CenterHorizontally} centerV={ps.CenterVertically}"
        )
        if info:
            print(
                f"[PDF-DEBUG] etapa={etapa} aba='{ws.name}' "
                f"scan={info.get('scan_rows')}x{info.get('scan_cols')} "
                f"used_last={info.get('used_last_row')}x{info.get('used_last_col')} "
                f"detected_last={info.get('last_row')}x{info.get('last_col')}"
            )
    except Exception as exc:
        print(f"[PDF-DEBUG] falha ao logar layout da aba '{ws.name}': {exc}")


def _expandir_print_area_segura(ws, extra_linhas=2, extra_colunas=1):
    """
    Expande levemente a PrintArea atual para evitar corte de borda/rodape
    em ambientes com diferenca de driver/motor de PDF do Excel.
    """
    try:
        ps = ws.api.PageSetup
        atual = str(ps.PrintArea or "").strip()
        if not atual:
            return

        area = ws.api.Range(atual)
        row1 = int(area.Row)
        col1 = int(area.Column)
        row2 = int(area.Row + area.Rows.Count - 1)
        col2 = int(area.Column + area.Columns.Count - 1)

        max_row = int(ws.api.Rows.Count)
        max_col = int(ws.api.Columns.Count)

        row2 = min(max_row, row2 + max(0, int(extra_linhas)))
        col2 = min(max_col, col2 + max(0, int(extra_colunas)))

        ps.PrintArea = ws.api.Range(ws.api.Cells(row1, col1), ws.api.Cells(row2, col2)).Address
    except Exception:
        pass


def _aplicar_page_setup_a4(
    ws,
    uma_pagina=True,
    preservar_print_area=False,
    min_linhas=None,
    min_colunas=8,
    max_linhas_scan=360,
    max_colunas_scan=60,
):
    """
    Aplica configuracao de impressao equivalente ao preview manual do Excel.

    Ajusta para A4 e recalcula PrintArea pela area real para evitar PDF reduzido.
    """
    xlPortrait = 1
    xlPaperLetter = 1  # Letter (21.59 x 27.94 cm)

    app = ws.book.app
    ps = ws.api.PageSetup

    info = None
    if not preservar_print_area:
        info = limpar_planilha_para_exportacao(
            ws,
            min_linhas=min_linhas if min_linhas is not None else (45 if uma_pagina else 40),
            min_colunas=min_colunas,
            max_linhas_scan=max_linhas_scan,
            max_colunas_scan=max_colunas_scan,
        )
    else:
        try:
            info = {"print_area": ps.PrintArea}
        except Exception:
            info = None

    # Margens estreitas (preset "Narrow" do Excel)
    margem_lr = app.api.CentimetersToPoints(0.64)
    margem_tb = app.api.CentimetersToPoints(1.91)
    margem_hf = app.api.CentimetersToPoints(0.76)

    ps.Orientation = xlPortrait
    ps.PaperSize = xlPaperLetter
    ps.TopMargin = margem_tb
    ps.BottomMargin = margem_tb
    ps.LeftMargin = margem_lr
    ps.RightMargin = margem_lr
    ps.HeaderMargin = margem_hf
    ps.FooterMargin = margem_hf

    # FitToPages exige Zoom desativado.
    ps.Zoom = False
    ps.FitToPagesWide = 1
    ps.FitToPagesTall = 1 if uma_pagina else False

    ps.CenterHorizontally = True
    # Centralizacao vertical costuma aumentar espaco em branco perceptivel.
    ps.CenterVertically = False

    _log_layout_debug(ws, "page_setup_aplicado", info)


def ajustar_layout_report_vigia(ws_report):
    """
    Configura layout de impressao A4 para aba REPORT VIGIA.
    Ajuste dinamico para sempre caber em 1 folha (reduz escala quando necessario).
    """
    try:
        app = ws_report.book.app
        ps = ws_report.api.PageSetup

        # Detecta area util real para evitar encolhimento excessivo no A4.
        info = limpar_planilha_para_exportacao(
            ws_report,
            min_linhas=40,
            min_colunas=8,
            max_linhas_scan=1200,
            # Limita colunas para manter apenas a area principal do report (borda esquerda).
            max_colunas_scan=10,
        )

        # Evita corte da ultima borda: adiciona folga pequena ao recorte.
        last_row_safe = min(int(info.get("last_row", 40)) + 4, int(info.get("scan_rows", 1200)))
        last_col_safe = min(int(info.get("last_col", 8)) + 1, int(info.get("scan_cols", 80)))
        area_segura = ws_report.api.Range(
            ws_report.api.Cells(1, 1),
            ws_report.api.Cells(last_row_safe, last_col_safe),
        )
        ps.PrintArea = area_segura.Address
        info["print_area"] = area_segura.Address

        # Margens estreitas (preset "Narrow" do Excel)
        margem_lr = app.api.CentimetersToPoints(0.64)
        margem_tb = app.api.CentimetersToPoints(1.91)
        margem_hf = app.api.CentimetersToPoints(0.76)

        xlPortrait = 1
        xlPaperLetter = 1  # Letter (21.59 x 27.94 cm)

        ps.Orientation = xlPortrait
        ps.PaperSize = xlPaperLetter
        ps.TopMargin = margem_tb
        ps.BottomMargin = margem_tb
        ps.LeftMargin = margem_lr
        ps.RightMargin = margem_lr
        ps.HeaderMargin = margem_hf
        ps.FooterMargin = margem_hf

        # Sempre caber em 1 pagina para evitar corte quando houver mais linhas.
        ps.Zoom = False
        ps.FitToPagesWide = 1
        ps.FitToPagesTall = 1
        ps.CenterHorizontally = True
        ps.CenterVertically = False

        print(
            f"[PDF-DEBUG] REPORT VIGIA used_range={info.get('used_last_row')}x{info.get('used_last_col')} "
            f"| detected={info.get('last_row')}x{info.get('last_col')} "
            f"| print_area={info.get('print_area')}"
        )
    except Exception as e:
        print(f"Nao foi possivel ajustar layout do REPORT VIGIA: {e}")


def ajustar_layout_report_vigia_no_wb(wb):
    for ws in wb.sheets:
        if ws.name.strip().upper() == "REPORT VIGIA":
            ajustar_layout_report_vigia(ws)
            break


def ajustar_layout_front_vigia(ws_front):
    """
    Configura layout de impressao A4 para aba FRONT VIGIA.
    Uma pagina: zoom preenche A4 garantindo que caiba em 1 folha.
    """
    try:
        ps = ws_front.api.PageSetup
        print_area_atual = str(ps.PrintArea or "").strip()

        # FRONT costuma ter bordas em celulas sem valor; preservar PrintArea evita corte no cliente.
        if print_area_atual:
            _aplicar_page_setup_a4(ws_front, uma_pagina=True, preservar_print_area=True)
            _expandir_print_area_segura(ws_front, extra_linhas=2, extra_colunas=1)
        else:
            # Fallback conservador para manter quadro completo quando o template vier sem PrintArea.
            ps.PrintArea = ws_front.api.Range("A1:J52").Address
            _aplicar_page_setup_a4(ws_front, uma_pagina=True, preservar_print_area=True)
            _expandir_print_area_segura(ws_front, extra_linhas=2, extra_colunas=1)
    except Exception as e:
        print(f"Nao foi possivel ajustar layout da FRONT VIGIA: {e}")


def ajustar_layout_quitacao_credit_note(ws):
    """
    Enquadra somente a folha principal da aba Quitacao/Credit Note.
    Ignora blocos auxiliares laterais para nao "puxar" conteudo alem da borda.
    """
    try:
        ps = ws.api.PageSetup
        print_area_atual = str(ps.PrintArea or "").strip()

        if print_area_atual:
            _aplicar_page_setup_a4(ws, uma_pagina=True, preservar_print_area=True)
            _expandir_print_area_segura(ws, extra_linhas=3, extra_colunas=1)
            _log_layout_debug(ws, "quitacao_credit_note_preservado", {"print_area": print_area_atual})
            return

        # Fallback fixo: mantem a folha principal mesmo sem PrintArea no modelo.
        ps.PrintArea = ws.api.Range("A1:J52").Address
        _aplicar_page_setup_a4(ws, uma_pagina=True, preservar_print_area=True)
        _expandir_print_area_segura(ws, extra_linhas=3, extra_colunas=1)
        _log_layout_debug(ws, "quitacao_credit_note_fallback", {"print_area": ps.PrintArea})
    except Exception as e:
        print(f"Nao foi possivel ajustar layout da aba '{ws.name}': {e}")


def ajustar_layout_de_acordo(ws):
    """
    Preserva a area de impressao do template DE ACORDO.
    Esses modelos podem ter bordas e molduras fora das celulas com valor,
    entao recalcular a area util costuma desalinha-las no PDF.
    """
    try:
        ps = ws.api.PageSetup
        print_area_atual = str(ps.PrintArea or "").strip()

        if print_area_atual:
            _aplicar_page_setup_a4(ws, uma_pagina=True, preservar_print_area=True)
            _expandir_print_area_segura(ws, extra_linhas=2, extra_colunas=1)
            _log_layout_debug(ws, "de_acordo_preservado", {"print_area": print_area_atual})
            return

        ps.PrintArea = ws.api.Range("A1:J52").Address
        _aplicar_page_setup_a4(ws, uma_pagina=True, preservar_print_area=True)
        _expandir_print_area_segura(ws, extra_linhas=2, extra_colunas=1)
        _log_layout_debug(ws, "de_acordo_fallback", {"print_area": ps.PrintArea})
    except Exception as e:
        print(f"Nao foi possivel ajustar layout DE ACORDO da aba '{ws.name}': {e}")


def ajustar_layout_front_vigia_no_wb(wb):
    for ws in wb.sheets:
        if ws.name.strip().upper() == "FRONT VIGIA":
            ajustar_layout_front_vigia(ws)
            break


def ajustar_layout_planilha_generica(ws, **_kwargs):
    """
    Configura layout A4 generico com zoom calculado.
    Para abas que nao sao FRONT ou REPORT VIGIA.
    """
    try:
        _aplicar_page_setup_a4(ws, uma_pagina=True)
    except Exception as e:
        print(f"Nao foi possivel ajustar layout da aba '{ws.name}': {e}")


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

    if (
        nome == "QUITACAO"
        or "QUITA" in nome
        or nome == "CREDIT NOTE"
    ):
        ajustar_layout_quitacao_credit_note(ws)
        return

    ajustar_layout_planilha_generica(ws)


def aplicar_layout_pdf_especifico(ws, tipo_layout=None):
    tipo_norm = _normalizar_nome_aba_layout(tipo_layout or "")

    if tipo_norm == "DE ACORDO":
        ajustar_layout_de_acordo(ws)
        return

    ajustar_layout_pdf_por_aba(ws)


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


def _restaurar_page_setup_para_visualizacao(ws):
    """
    Reseta PageSetup de uma aba para visualizacao normal no Excel,
    desfazendo FitToPages e restaurando PrintArea generosa.
    Isso evita que bordas fiquem cortadas em PCs com drivers de
    impressora diferentes.
    """
    try:
        app = ws.book.app
        ps = ws.api.PageSetup

        xlPortrait = 1
        xlPaperLetter = 1  # Letter (21.59 x 27.94 cm)

        # Margens estreitas (mesmo preset "Narrow")
        margem_lr = app.api.CentimetersToPoints(0.64)
        margem_tb = app.api.CentimetersToPoints(1.91)
        margem_hf = app.api.CentimetersToPoints(0.76)

        ps.Orientation = xlPortrait
        ps.PaperSize = xlPaperLetter
        ps.TopMargin = margem_tb
        ps.BottomMargin = margem_tb
        ps.LeftMargin = margem_lr
        ps.RightMargin = margem_lr
        ps.HeaderMargin = margem_hf
        ps.FooterMargin = margem_hf

        # Desativa FitToPages e usa zoom fixo para nao depender da impressora.
        ps.FitToPagesWide = False
        ps.FitToPagesTall = False
        ps.Zoom = 100

        ps.CenterHorizontally = True
        ps.CenterVertically = False

        # PrintArea generosa para incluir todas as bordas do template.
        ps.PrintArea = ws.api.Range("A1:J52").Address
    except Exception as e:
        print(f"Nao foi possivel restaurar layout de visualizacao da aba '{ws.name}': {e}")


def restaurar_layout_excel_para_visualizacao(wb, ignorar_abas=()):
    """
    Restaura PageSetup de todas as abas visiveis do workbook para
    visualizacao correta no Excel (sem FitToPages).
    Deve ser chamado APOS a geracao do PDF e ANTES de salvar o .xlsx.
    """
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

        _restaurar_page_setup_para_visualizacao(ws)


def gerar_pdf_faturamento_completo(
    wb,
    pasta_saida: Path,
    nome_base: str,
    apenas_front=False,
    aplicar_layout=True,
) -> Path:
    caminho_pdf = pasta_saida / f"{nome_base}.pdf"

    if caminho_pdf.exists():
        caminho_pdf.unlink()

    # âœ… Se apenas_front=True, exporta sÃ³ FRONT VIGIA
    if apenas_front:
        # Encontra a aba FRONT VIGIA
        ws_front = None
        for ws in wb.sheets:
            if ws.name.strip().upper() == "FRONT VIGIA":
                ws_front = ws
                break
        
        if ws_front:
            if aplicar_layout:
                ajustar_layout_pdf_por_aba(ws_front)
                compactar_layout_final_ws(ws_front)
            # Exporta apenas essa aba
            ws_front.api.ExportAsFixedFormat(
                Type=0,  # PDF
                Filename=str(caminho_pdf),
                Quality=0,
                IncludeDocProperties=True,
                IgnorePrintAreas=False,
                OpenAfterPublish=False
            )
            try:
                normalizar_pdf_final_para_a4(caminho_pdf)
            except Exception:
                pass
            return caminho_pdf
        else:
            raise RuntimeError("Aba FRONT VIGIA nÃ£o encontrada para exportar PDF")

    # ðŸ”’ Oculta aba NF (se existir)
    aba_nf = None
    for ws in wb.sheets:
        if ws.name.strip().upper() == "NF":
            aba_nf = ws
            ws.api.Visible = False
            break

    if aplicar_layout:
        ajustar_layout_todas_abas_visiveis_no_wb(wb, ignorar_abas=("NF",))
        compactar_layout_final_wb(wb, ignorar_abas=("NF",))

    # ðŸ“„ Exporta workbook inteiro
    wb.api.ExportAsFixedFormat(
        Type=0,  # PDF
        Filename=str(caminho_pdf),
        Quality=0,
        IncludeDocProperties=True,
        IgnorePrintAreas=False,
        OpenAfterPublish=False
    )

    try:
        normalizar_pdf_final_para_a4(caminho_pdf)
    except Exception:
        pass

    # ðŸ”“ Reexibe NF
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


#===================FATURAMENTO SÃƒO SEBASTIÃƒO=========================================#


def gerar_pdf_do_wb_aberto(wb, pasta_saida, nome_base, ignorar_abas=("nf",), apenas_front=False):
    caminho_pdf = Path(pasta_saida) / f"{nome_base}.pdf"

    # 1) se existir e estiver aberto, jÃ¡ avisa o motivo
    if caminho_pdf.exists():
        try:
            caminho_pdf.unlink()
        except Exception as e:
            raise RuntimeError(f"PDF estÃ¡ aberto/travado e nÃ£o pode ser sobrescrito: {caminho_pdf}") from e

    app = wb.app
    app.api.DisplayAlerts = False

    # âœ… Se apenas_front=True, exporta sÃ³ FRONT VIGIA
    if apenas_front:
        # Encontra a aba FRONT VIGIA
        ws_front = None
        for sh in wb.sheets:
            if sh.name.strip().upper() == "FRONT VIGIA":
                ws_front = sh
                break
        
        if ws_front:
            ajustar_layout_pdf_por_aba(ws_front)
            compactar_layout_final_ws(ws_front)
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
            try:
                normalizar_pdf_final_para_a4(caminho_pdf)
            except Exception:
                pass
            print(f"ðŸ“„ PDF gerado (apenas FRONT VIGIA): {caminho_pdf}")
            return caminho_pdf
        else:
            raise RuntimeError("Aba FRONT VIGIA nÃ£o encontrada para exportar PDF")

    # 2) guarda visibilidade, oculta as que nÃ£o devem sair no PDF
    vis_orig = {}
    for sh in wb.sheets:
        nome_norm = sh.name.strip().lower()
        vis_orig[sh.name] = sh.api.Visible
        if nome_norm in {x.strip().lower() for x in ignorar_abas}:
            sh.api.Visible = False  # oculta NF

    ajustar_layout_todas_abas_visiveis_no_wb(wb, ignorar_abas=ignorar_abas)
    compactar_layout_final_wb(wb, ignorar_abas=ignorar_abas)

    try:
        # 3) ativa uma aba visÃ­vel (Excel odeia export sem sheet ativa)
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

        try:
            normalizar_pdf_final_para_a4(caminho_pdf)
        except Exception:
            pass

        print(f"ðŸ“„ PDF gerado: {caminho_pdf}")
        return caminho_pdf

    finally:
        # 5) restaura visibilidade original
        for sh in wb.sheets:
            if sh.name in vis_orig:
                sh.api.Visible = vis_orig[sh.name]




# ==============================
# LICENÃ‡A E DATA
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

    # ðŸ”¥ define uma data fixa de expiraÃ§Ã£o: 30 de marÃ§o de 2026
    limite = datetime(2026, 3, 30, tzinfo=timezone.utc)

    if hoje_utc > limite:
        sys.exit("â›” LicenÃ§a expirada")

    print(f"ðŸ“… Data local: {hoje_local.date()}")

