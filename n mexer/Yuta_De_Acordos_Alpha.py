import xlwings as xw
from pathlib import Path
import tkinter as tk
from tkinter import filedialog
import re
import sys
from datetime import datetime

# =========================
# FUNÇÕES AUXILIARES
# =========================

def selecionar_pasta_com_arquivos():
    """Usuário seleciona a pasta que contém 1.xlsx e 2.xlsx"""
    root = tk.Tk()
    root.withdraw()
    pasta_str = filedialog.askdirectory(title="Selecione a pasta que contém 2.xlsx")
    if not pasta_str:
        print("❌ Nenhuma pasta selecionada. Encerrando.")
        sys.exit(0)
    
    pasta = Path(pasta_str)
    print(f"📂 Pasta selecionada: {pasta.name}")
    return pasta


def obter_dn_da_pasta(pasta: Path):
    numeros = re.findall(r"\d+", pasta.name)
    if not numeros:
        raise ValueError("Não foi possível identificar o DN no nome da pasta.")
    return numeros[0]


def obter_nome_navio_da_pasta(pasta: Path):
    # Remove número inicial + separadores como -, _, espaço
    nome_limpo = re.sub(r"^\d+[\s\-_]*", "", pasta.name, flags=re.IGNORECASE).strip()
    return nome_limpo if nome_limpo else "NAVIO NÃO IDENTIFICADO"


# =========================
# CLASSE FRONT VIGIA
# =========================

class FrontVigiaProcessor:
    def __init__(self, ws_front, dn, nome_navio):
        self.ws = ws_front
        self.dn = dn
        self.nome_navio = nome_navio

    def executar(self):
        hoje = datetime.now()
        meses = ["", "janeiro", "fevereiro", "março", "abril", "maio", "junho",
                 "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"]
        data_extenso = f"{hoje.day} de {meses[hoje.month]} de {hoje.year}"

        texto_dn = f"DN: {self.dn}/{hoje.year}"

        # Preenchimentos
        self.ws.range("D15").value = self.nome_navio                    # Nome do navio
        self.ws.range("C21").value = texto_dn                           # DN
        self.ws.range("D16").value = data_extenso                       # Data berthed
        self.ws.range("D17").value = data_extenso                       # Data sailed
        self.ws.range("D18").value = "-"                                # Berço
        self.ws.range("C26").value = f'  DE ACORDO ( M/V "{self.nome_navio}" )'
        self.ws.range("C27").value = f'  VOY SA02325'

        # === LIMPEZA DA CÉLULA G27 ===
        self.ws.range("G27").value = None   # ou .clear_contents() se quiser limpar formatação também
        print("🧹 Célula G27 limpa (conteúdo apagado)")


        # Regra especial Unimar (se o cliente for Unimar)
        cliente_c9 = str(self.ws.range("C9").value or "").strip()
        if "Unimar Agenciamentos" in cliente_c9:
            self.ws.range("G26").value = 400

        # Taxa administrativa +20 (C33)
        try:
            valor_atual = float(self.ws.range("C35").value or 0)
            self.ws.range("C35").value = valor_atual + 20
        except:
            self.ws.range("C35").value = 20

        # Rodapé
        self.ws.range("C39").value = f"Santos, {hoje.day} de {meses[hoje.month]} de {hoje.year}"

        print("✅ FRONT VIGIA preenchido com sucesso!")


# =========================
# PROGRAMA PRINCIPAL
# =========================

def main():
    print("🚀 Iniciando Gerador de FRONT VIGIA - SANPORT (versão Área de Trabalho)\n")

    try:
        # 1. Usuário seleciona a pasta com o 2.xlsx
        pasta_selecionada = selecionar_pasta_com_arquivos()

        # 2. Verifica apenas o arquivo necessário: 2.xlsx
        arquivo2 = pasta_selecionada / "2.xlsx"

        if not arquivo2.exists():
            raise FileNotFoundError("Arquivo '2.xlsx' não encontrado na pasta selecionada.")

        # 3. Extrai DN e nome do navio do nome da pasta
        dn = obter_dn_da_pasta(pasta_selecionada)
        nome_navio = obter_nome_navio_da_pasta(pasta_selecionada)
        print(f"📋 DN: {dn}")
        print(f"🚢 Navio: {nome_navio}")

        # 4. Abre APENAS o 2.xlsx
        app = xw.App(visible=False)
        wb2 = app.books.open(str(arquivo2))

        # 5. Seleciona a aba FRONT VIGIA
        if "FRONT VIGIA" not in [s.name for s in wb2.sheets]:
            raise RuntimeError("Aba 'FRONT VIGIA' não encontrada no arquivo 2.xlsx")
        
        ws_front = wb2.sheets["FRONT VIGIA"]

        # 6. Preenche a FRONT VIGIA
        processor = FrontVigiaProcessor(ws_front, dn, nome_navio)
        processor.executar()

        # 7. Prepara salvamento na Área de Trabalho
        desktop = Path.home() / "Desktop"
        arquivo_excel = desktop / f"3 - DN_{dn}.xlsx"
        arquivo_pdf = desktop / f"3 - DN_{dn}.pdf"

        # Remove todas as abas exceto FRONT VIGIA
        for sheet in list(wb2.sheets):
            if sheet.name != "FRONT VIGIA":
                sheet.delete()

        # === SALVA EM EXCEL ===
        wb2.save(str(arquivo_excel))
        print(f"📄 Excel salvo: {arquivo_excel.name}")

        # === EXPORTA PARA PDF ===
        ws_front.api.ExportAsFixedFormat(
            Type=0,  # xlTypePDF
            Filename=str(arquivo_pdf),
            Quality=0,  # xlQualityStandard
            IncludeDocProperties=True,
            IgnorePrintAreas=False,
            OpenAfterPublish=False
        )
        print(f"📄 PDF salvo: {arquivo_pdf.name}")

        # Fecha tudo
        wb2.close()
        app.quit()

        print(f"\n✅ CONCLUÍDO COM SUCESSO!")
        print(f"Arquivos gerados na Área de Trabalho:")
        print(f"   📊 {arquivo_excel.name}")
        print(f"   📄 {arquivo_pdf.name}")

    except Exception as e:
        print(f"\n❌ ERRO: {e}")
        try:
            if 'app' in locals():
                app.quit()
        except:
            pass
        sys.exit(1)

if __name__ == "__main__":
    main()