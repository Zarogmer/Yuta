from pathlib import Path
from tempfile import gettempdir

from yuta_helpers import (
    ajustar_layout_pdf_por_aba,
    abrir_workbooks_de_acordo,
    escrever_de_acordo_nf,
    fechar_workbooks,
    gerar_pdf,
    montar_nome_faturamento,
    obter_dn_da_pasta,
    obter_nome_navio,
    obter_pasta_faturamentos,
    salvar_excel_com_nome,
    selecionar_pasta_navio,
)
from yuta_helpers import datetime

from .criar_pasta import CriarPasta
from .email_rascunho import criar_rascunho_email_cliente


class FaturamentoDeAcordo:

    def __init__(self, usuario_nome: str | None = None):
        self.usuario_nome = (usuario_nome or "").strip()

    @staticmethod
    def limpar_celula_segura(ws, endereco):
        rng = ws.range(endereco)
        if rng.merge_cells:
            rng.merge_area.clear_contents()
        else:
            rng.clear_contents()

    @staticmethod
    def escrever_celula_segura(ws, endereco, valor):
        rng = ws.range(endereco)
        if rng.merge_cells:
            rng.merge_area.value = valor
        else:
            rng.value = valor


    @staticmethod
    def aplicar_regras(ws_front, regras):
        for celula, valor in regras.items():

            FaturamentoDeAcordo.limpar_celula_segura(ws_front, celula)

            if valor is not None:
                FaturamentoDeAcordo.escrever_celula_segura(ws_front, celula, valor)

        for extra in ("C27", "G27"):
            FaturamentoDeAcordo.limpar_celula_segura(ws_front, extra)

    # =========================
    # REGRAS POR CLIENTE
    # =========================
    REGRAS_CLIENTES = {
        "Unimar Agenciamentos": {
            "G26": 500,
            "C27": None,
            "C35": 23.85,
        },
        "A/C Delta Agenciamento Marítimo Ltda.": {
            "G26": 500,
            "C27": None,
        },
        "A/C NORTH STAR SUDESTE SERVIÇOS MARÍTIMOS LTDA.": {
            "G26": 500,
            "C27": None,
            "C28": None,
            "C29": None,
            "H28": None,
            "H29": None,
        },
    }

    # =========================
    # APLICA REGRAS
    # =========================
    @staticmethod
    def aplicar_regras_cliente(ws_front):
        cliente_c9 = str(ws_front.range("C9").value or "").strip()

        for nome_cliente, regras in FaturamentoDeAcordo.REGRAS_CLIENTES.items():
            if nome_cliente in cliente_c9:
                FaturamentoDeAcordo.aplicar_regras(ws_front, regras)
                print(f"🔧 Regras aplicadas para cliente: {nome_cliente}")
                return

        print("ℹ️ Nenhuma regra específica de cliente aplicada.")



    def executar(self, preview=False, selection=None):
        print("🚀 Iniciando execução (DE ACORDO)...")

        pasta_faturamentos = obter_pasta_faturamentos()
        if selection and isinstance(selection, dict) and selection.get("pasta_navio"):
            pasta_navio = Path(selection["pasta_navio"])
        else:
            pasta_navio = selecionar_pasta_navio()

        dn = obter_dn_da_pasta(pasta_navio)
        nome_navio = obter_nome_navio(pasta_navio)

        nome_base = montar_nome_faturamento(dn, nome_navio)

        app = wb = ws_front = None

        try:
            app, wb, ws_front = abrir_workbooks_de_acordo(
                pasta_faturamentos,
                pasta_navio
            )

            print(f"📋 DN: {dn}")
            print(f"🚢 Navio: {nome_navio}")
            escrever_de_acordo_nf(wb, nome_navio, dn, ano=datetime.now().year)

            hoje = datetime.now()
            meses = ["", "janeiro","fevereiro","março","abril","maio","junho",
                    "julho","agosto","setembro","outubro","novembro","dezembro"]
            data_extenso = f"{hoje.day} de {meses[hoje.month]} de {hoje.year}"

            # -------- PREENCHIMENTO FRONT --------
            ws_front.range("D15").value = nome_navio
            ws_front.range("C21").value = f"DN {str(dn).zfill(3)}/{hoje.strftime('%y')}"

            ws_front.range("D16").value = data_extenso
            ws_front.range("D17").value = data_extenso
            ws_front.range("D18").value = "-"
            ws_front.range("C26").value = f"DE ACORDO ( M/V {nome_navio} )"
            ws_front.range("C39").value = f" Santos, {data_extenso}"

            cliente_front = str(ws_front.range("C9").value or "").upper()
            if self.usuario_nome and "NORTH STAR" not in cliente_front:
                ws_front.range("C42").value = f"   {self.usuario_nome}"

            # 🔧 Regras por cliente
            self.aplicar_regras_cliente(ws_front)
            
            # ⚠️ NÃO atualiza planilha de controle aqui (será feito após preview)

            print("✅ Faturamento De Acordo concluído!")

            if preview:
                preview_pdf = self._export_preview_pdf(ws_front, nome_base)
                return {
                    "text": "",
                    "preview_pdf": str(preview_pdf) if preview_pdf else None,
                    "selection": {"pasta_navio": str(pasta_navio)},
                }

            # ✅ ATUALIZAR PLANILHA DE CONTROLE (só na execução final)
            self._atualizar_planilha_controle(pasta_navio, nome_navio, dn, data_extenso, ws_front)

            # ✅ SALVAR EXCEL (ainda dentro do try, com wb aberto)
            caminho_excel = salvar_excel_com_nome(
                wb=wb,
                pasta_saida=pasta_navio,
                nome_base=nome_base
            )
            print(f"💾 Excel salvo em: {caminho_excel}")

            # ✅ GERAR PDF (passando caminho_excel corretamente)
            caminho_pdf = gerar_pdf(
                caminho_excel=caminho_excel,
                pasta_saida=pasta_navio,
                nome_base=nome_base,
                ws=ws_front
            )
            print(f"📑 PDF FRONT salvo em: {caminho_pdf}")

            nome_cliente = pasta_navio.parent.name.strip()
            anexos = [caminho_pdf]  # ✅ Removido Excel dos anexos
            try:
                criar_rascunho_email_cliente(
                    nome_cliente,
                    anexos=anexos,
                    dn=str(dn),
                    navio=nome_navio,
                    usuario_nome=self.usuario_nome,
                )
                print("✅ Rascunho do Outlook criado com anexos.")
            except Exception as e:
                print(f"⚠️ Nao foi possivel criar rascunho do Outlook: {e}")

        finally:
            fechar_workbooks(app=app, wb_cliente=wb)

    def _atualizar_planilha_controle(self, pasta_navio: Path, nome_navio: str, dn: str, data_extenso: str, ws_front):
        """
        Atualiza a planilha de controle com informações do DE ACORDO.
        Preenche colunas B (data), C (serviço), D (ETA), E (ETB), F (cliente), G (navio), J (DN), K (valor total), O (ISS).
        """
        try:
            cliente = pasta_navio.parent.name.strip()
            
            # Obter data atual em formato dd/mm/yyyy
            from datetime import datetime
            data_hoje = datetime.now().strftime("%d/%m/%Y")
            
            # ✅ Ler valor da FRONT VIGIA para coluna K no controle
            # Regra DE ACORDO: usar G26
            valor_total = ws_front.range("G26").value
            celula_usada = "G26"
            
            print(f"📊 Lendo valores da FRONT VIGIA:")
            print(f"   {celula_usada} (Valor Total): {valor_total} (tipo: {type(valor_total)})")
            
            if valor_total is None:
                print(f"⚠️ AVISO: G26 está vazio! Verifique a planilha FRONT VIGIA.")
            
            # ✅ Abrir workbook de controle uma única vez
            criar_pasta = CriarPasta()
            caminho_planilha = criar_pasta._encontrar_planilha()
            from yuta_helpers import openpyxl
            wb_controle = openpyxl.load_workbook(caminho_planilha)
            
            try:
                # Para DE ACORDO, D16 e D17 são iguais (mesma data) - usar formato dd/mm/yyyy
                # Usar CriarPasta para gravar na planilha (reutilizando workbook)
                criar_pasta._gravar_planilha(
                    cliente=cliente,
                    navio=nome_navio,
                    dn=dn,
                    servico="DE ACORDO",
                    data=data_hoje,
                    eta=data_hoje,  # Mesmo dia
                    etb=data_hoje,  # Mesmo dia
                    mmo=valor_total,  # Valor de G26 vai para coluna K
                    wb_externo=wb_controle,
                    iss_formula=True,  # Cria fórmula =K{linha}*5% na coluna O
                    limpar_formulas_adm_cliente=True  # Limpa colunas N e P (ADM % e CLIENTE %)
                )
                
                # ✅ Salvar apenas uma vez
                criar_pasta.salvar_planilha_com_retry(wb_controle, caminho_planilha)
                print("✅ Planilha de controle atualizada com sucesso!")
            finally:
                # Fechar workbook
                wb_controle.close()
            
        except Exception as e:
            print(f"⚠️ Erro ao atualizar planilha de controle: {e}")
            # Não levanta exceção para não interromper o fluxo principal

    def _build_preview_text(self, nome_base, dn, nome_navio, data_extenso):
        linhas = [
            "PRE-VISUALIZACAO",
            "Processo: De Acordo",
            f"Nome base: {nome_base}",
            f"DN: {dn}",
            f"Navio: {nome_navio}",
            f"Data: {data_extenso}",
        ]
        return "\n".join(linhas)

    def _export_preview_pdf(self, ws_front, nome_base):
        caminho_pdf = Path(gettempdir()) / f"preview_{nome_base}.pdf"
        if caminho_pdf.exists():
            caminho_pdf.unlink()

        ajustar_layout_pdf_por_aba(ws_front)
        ws_front.activate()
        ws_front.api.ExportAsFixedFormat(
            Type=0,
            Filename=str(caminho_pdf),
            Quality=0,
            IncludeDocProperties=True,
            IgnorePrintAreas=False,
            OpenAfterPublish=False,
        )
        return caminho_pdf
