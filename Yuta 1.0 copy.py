import xlwings as xw
from datetime import datetime, timedelta
from openpyxl.styles import Font, Alignment
import pandas as pd
import locale
import os
import sys
from datetime import date
import urllib.request
import sys
import ssl
import certifi
from datetime import datetime, timezone, date




# Abre Excel em background
app = xw.App(visible=False, add_book=False)

def main():

    print("==========Sanport==========")
    # Seu código principal aqui




def aplicar_fonte_xw(ws, celula, tamanho=15):
    """
    Aplica fonte Arial tamanho definido (padrão = 15)
    e centraliza a célula horizontal e verticalmente usando xlwings.
    """
    rng = ws.range(celula)
    rng.api.Font.Name = "Calibri"
    rng.api.Font.Size = tamanho
    rng.api.Font.Bold = False
    rng.api.Font.Italic = False
    rng.api.Font.Underline = False
    rng.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignCenter
    rng.api.VerticalAlignment = xw.constants.VAlign.xlVAlignCenter





def formatar_coluna_c_data(ws, linha_inicial, periodos):
    """
    Aplica formatação de data (DD/MM/YYYY) e alinhamento centralizado
    na coluna C a partir da linha_inicial pelo número de 'periodos'.
    """
    alinhamento = Alignment(horizontal="center", vertical="center")
    formato_data = "dd/mm/yyyy"



    for offset in range(periodos):
        linha_atual = linha_inicial + offset
        celula = ws[f"C{linha_atual}"]

        celula.number_format = formato_data
        celula.alignment = alinhamento













try:

    main()
    from pathlib import Path

    desktop = Path.home() / "Desktop"

    arquivo1 = (desktop / "1.xlsx").resolve()
    arquivo2 = (desktop / "2.xlsx").resolve()

    # Abre workbooks
    wb1 = app.books.open(arquivo1)
    wb2 = app.books.open(arquivo2)

    ws1 = wb1.sheets[0]              # primeira aba do 1.xlsx (ativa no código original)
    ws_front = wb2.sheets["FRONT VIGIA"]   # primeira aba sensível (preservar shapes)
    # ws_report e demais serão acessadas pelo nome mais abaixo quando existirem

    # =========================
    # 1 – Transferências diretas (usando xlwings na FRONT VIGIA)
    # =========================

    if "REPORT VIGIA" not in [s.name for s in wb2.sheets]:
        raise RuntimeError("Aba 'REPORT VIGIA' não encontrada em 2.xlsx")

    ws_report = wb2.sheets["REPORT VIGIA"]


    linha_origem = 22  # linha-base de onde você copia os estilos   
    # Captura altura da linha origem (RowHeight)
    row_height = ws_report.api.Rows(linha_origem).RowHeight
    # Localiza última célula preenchida da coluna AA no ws_report
    # Pega todas as células preenchidas na coluna AA manualmente
    # 1 – captura com end('down')
   

    df = pd.read_excel(arquivo1, sheet_name="Resumo", header=None)

    # Coluna AA = índice 26
    col_aa = df[26]

    # Remove vazios
    col_aa = col_aa.dropna()

    # Último valor preenchido
    ultimo = col_aa.iloc[-1]

    # Converte para inteiro
    periodos = int(float(str(ultimo).replace("R$", "").replace(",", ".").replace(" ", "")))


    # Inserir e copiar linha 22 para cada período (fazendo uma cópia real no Excel)
    for i in range(periodos):
        destino = linha_origem + 1 + i
        # Inserir linha em branco na posição destino
        ws_report.api.Rows(destino).Insert()
        # Copiar linha origem para destino (preserva fórmulas, formatação, merges)
        ws_report.api.Rows(linha_origem).Copy(ws_report.api.Rows(destino))
        # Ajustar altura
        ws_report.api.Rows(destino).RowHeight = row_height

    formatar_coluna_b_mes_extenso(ws_report, linha_inicial=22, periodos=periodos)

    
    # =========================
    # 8 – Formatar coluna C como data (DD/MM/YYYY) e preencher datas conforme ciclo
    # =========================

    df = pd.read_excel(arquivo1, sheet_name="Resumo")

    # ---- 1) coluna C ----
    coluna_c = df.iloc[:, 2].astype(str).str.strip()

    # ---- 2) cortar no primeiro "Total" ----
    lista_reduzida = []
    for valor in coluna_c:
        if valor.lower() == "total":
            break
        lista_reduzida.append(valor)

    # ---- 3) normalizar ----
    horarios = [h.lower() for h in lista_reduzida if h.lower() in ["00h", "06h", "12h", "18h"]]

    # ---- 4) ordem desejada (00h sempre por último) ----
    ordem = ["06h", "12h", "18h", "00h"]

    # ---- 5) pegar o menor horário ----
    horarios_validos = [h for h in ordem if h in horarios]

    if horarios_validos:
        menor_horario = horarios_validos[0]
    else:
        menor_horario = "06h"  # fallback

    # ---- 6) ciclo ----
    ciclo = ["06x12", "12x18", "18x00", "00x06"]
    mapa = {"06h": 0, "12h": 1, "18h": 2, "00h": 3}

    periodo_escolhido = ciclo[mapa[menor_horario]]

    indice = mapa[menor_horario]   # ← ISSO PRECISA EXISTIR




    # DESMESCLAR COLUNA E das linhas geradas (se necessário)
    # Percorrer merges e limpar apenas se intersecta as células alvo
    # xlwings não tem API direta para merged_ranges fácil; usamos COM:
    merged_items = ws_report.api.UsedRange.MergeCells  # True/False; para detalhes usamos MergeAreas
    # Simples approach: tentar unmerge nas células E{linha} para cada nova linha
    for i in range(periodos):
        linha = linha_origem + i
        ws_report.range(f"E{linha}").value = ciclo[(indice + i) % len(ciclo)]
        try:
            rng = ws_report.range(f"E{linha}")
            if rng.api.MergeCells:
                rng.api.UnMerge()
        except Exception:
            pass
            
    # Preencher coluna E com ciclo
    for i in range(periodos):
        linha = linha_origem + i
        ws_report.range(f"E{linha}").value = ciclo[(indice + i) % len(ciclo)]

    # Aplicar fonte Arial tamanho 20 na coluna E das linhas geradas
    for i in range(periodos):
        linha = linha_origem + i
        rng = ws_report.range(f"E{linha}")
        
        aplicar_fonte_xw(ws_report, f"E{linha}", tamanho=18)

    # =========================
    # 9 – Pegar valores da coluna Z do arquivo1 quando coluna C for "00h","06h","12h","18h"
    # =========================
    valores_validos = []
    filtros = ["00h", "06h", "12h", "18h"]

    # Ler colunas C e Z do wb1 até o fim (assumindo dados contínuos)
    dadosC = ws1.range("C2").options(expand='down').value  # lista
    dadosZ = ws1.range("Z2").options(expand='down').value

    # garantir listas
    if not isinstance(dadosC, list):
        dadosC = [dadosC] if dadosC is not None else []
    if not isinstance(dadosZ, list):
        dadosZ = [dadosZ] if dadosZ is not None else []

    # iterar paralelo
    for c_val, z_val in zip(dadosC, dadosZ):
        if c_val is None:
            break
        c_norm = str(c_val).strip().lower()
        if c_norm in filtros:
            valores_validos.append(z_val)

    # cortar para tamanho de periodos
    valores_validos = valores_validos[:periodos]

    # =========================
    # 10 – Colar na coluna G do REPORT VIGIA a partir de G22
    # =========================
    linha_destino = linha_origem
    for valor in valores_validos:
        cel = ws_report.range(f"G{linha_destino}")
        cel.value = valor
        cel.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignRight
        cel.api.VerticalAlignment = xw.constants.VAlign.xlVAlignCenter
        # Formato moeda brasileiro com 2 casas decimais
        cel.api.NumberFormat = 'R$ #.##0,00'
        cel.api.Font.Name = "Calibri"
        cel.api.Font.Size = 18
        linha_destino += 1

    # =========================
    # Ajustar as datas na coluna C com base no ciclo
    # =========================
    # data_inicial vem de wb2 FRONT VIGIA D16
    di_val = ws_front.range("D16").value
    if isinstance(di_val, datetime):
        data_inicial = di_val
    else:
        # tentar converter string dd/mm/yyyy
        try:
            data_inicial = datetime.strptime(str(di_val), "%d/%m/%Y")
        except Exception:
            data_inicial = datetime.now()

    data_atual = data_inicial
    for i in range(periodos):
        linha = linha_origem + i
        periodo = ws_report.range(f"E{linha}").value

    # Preenche com STRING por extenso em inglês
        ws_report.range(f"C{linha}").value = data_por_extenso(data_atual)
        aplicar_fonte_xw(ws_report, f"C{linha}", tamanho=20)

        # Incremento da data baseado no período
        if periodo in ["00x06"]:
            data_atual += timedelta(days=1)

#========================
    # 11 – Arredondar para baixo o valor em E37 e colocar em H28 (FRONT VIGIA)
    # =========================
    arredondar_para_baixo_50(ws_front)
    validar_c9(ws_front)
    # =========================
    # SALVAR E FECHAR
    # =========================


    desktop = (Path.home() / "Desktop").resolve()
    novo_arquivo = desktop / "2_ATUALIZADO.xlsx"

    wb2.save(str(novo_arquivo))

    print("Arquivo salvo em:", novo_arquivo)


finally:
    try:
        wb1.close()
    except Exception:
        pass
    try:
        wb2.close()
    except Exception:
        pass
    app.quit()
