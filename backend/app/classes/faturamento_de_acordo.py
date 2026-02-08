from yuta_helpers import *


class FaturamentoDeAcordo:

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
            "C35": 25,
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



    def executar(self):
        print("🚀 Iniciando execução (DE ACORDO)...")

        pasta_faturamentos = obter_pasta_faturamentos()
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

            # 🔧 Regras por cliente
            self.aplicar_regras_cliente(ws_front)

            print("✅ Faturamento De Acordo concluído!")

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

        finally:
            fechar_workbooks(app=app, wb_cliente=wb)
