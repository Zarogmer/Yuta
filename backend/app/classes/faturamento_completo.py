from yuta_helpers import *
import pdfplumber
import re
from .email_rascunho import criar_rascunho_email_cliente
from .criar_pasta import CriarPasta


class FaturamentoCompleto:
    def __init__(self, g_logic=1):
        self.app = None
        self.wb1 = None
        self.wb2 = None
        self.ws1 = None
        self.ws_front = None
        self.nome_navio = None
        self.g_logic = g_logic
        self.pasta_saida_final = None
        self.dn = None
        self.pdf_path = None
        self.pasta_faturamentos = None  # <--- GUARDA PASTA AQUI



    def executar(self, preview=False, selection=None):
        print("🚀 Iniciando execução...")

        # 🔹 1️⃣ Buscar pasta FATURAMENTOS apenas 1x
        self.pasta_faturamentos = obter_pasta_faturamentos()
        caminho_navio_rede = None
        if selection and isinstance(selection, dict):
            caminho_navio_rede = selection.get("navio_path")

        resultado = abrir_workbooks(self.pasta_faturamentos, caminho_navio_rede)

        if not resultado:
            raise SystemExit("❌ Erro ou pasta inválida")

        (
            self.app,
            self.wb1,
            self.wb2,
            self.ws1,
            self.ws_front,
            pasta_navio_rede,
            caminho_navio_rede,
        ) = resultado

        self.pasta_saida_final = pasta_navio_rede

        # 🔹 Extrair DN e nome do navio
        self.dn, self.nome_navio = extrair_identidade_navio(pasta_navio_rede)

        # caminho PDF OGMO
        self.pdf_path = pasta_navio_rede / "FOLHAS OGMO.pdf"

        print(f"📌 DN: {self.dn}")
        print(f"🚢 NAVIO: {self.nome_navio}")
        print(f"📑 PDF OGMO: {self.pdf_path}")

        escrever_nf_faturamento_completo(self.wb2, self.nome_navio, self.dn)

        nome_base = f"FATURAMENTO - DN {self.dn} - MV {self.nome_navio}"

        try:
            if preview:
                total_periodos = self.processar_preview()
                preview_pdf = self._export_preview_pdf(nome_base)
                fechar_workbooks(self.app, self.wb1, self.wb2)
                return {
                    "text": "",
                    "preview_pdf": str(preview_pdf) if preview_pdf else None,
                    "selection": {"navio_path": str(caminho_navio_rede)},
                }

            self.processar()

            caminho_excel = pasta_navio_rede / f"{nome_base}.xlsx"
            caminho_pdf = pasta_navio_rede / f"{nome_base}.pdf"

            if caminho_excel.exists():
                caminho_excel.unlink()
            
            # ✅ Verifica se é cliente WILLIAMS (apenas FRONT VIGIA no PDF)
            nome_cliente = pasta_navio_rede.parent.name.strip()
            apenas_front = "WILLIAMS" in nome_cliente.upper()
            
            gerar_pdf_faturamento_completo(
                self.wb2,
                pasta_navio_rede,
                nome_base,
                apenas_front=apenas_front
            )

            # SALVAR EXCEL (local → rede)
            temp_excel = Path(gettempdir()) / f"{nome_base}.xlsx"
            if temp_excel.exists():
                temp_excel.unlink()

            self.wb2.save(str(temp_excel))
            shutil.copy2(temp_excel, caminho_excel)

            # PDF REPORT separado
            self.gerar_pdf_report_vigia_separado(
                pasta_navio_rede, self.dn, self.nome_navio
            )

            nome_cliente = pasta_navio_rede.parent.name.strip()
            caminho_report = (
                pasta_navio_rede
                / f"REPORT VIGIA - DN {self.dn} - MV {self.nome_navio}.pdf"
            )
            if nome_cliente.strip().upper() == "ROCHAMAR":
                anexos = [caminho_report] if caminho_report.exists() else []
            else:
                anexos = [caminho_pdf]  # ✅ Removido Excel dos anexos
                if caminho_report.exists():
                    anexos.append(caminho_report)
                if self.pdf_path and Path(self.pdf_path).exists():
                    anexos.append(self.pdf_path)
            try:
                adiantamento = self.obter_valor_cargonave()
                ws_report = self.wb2.sheets["REPORT VIGIA"]
                try:
                    self.wb2.app.calculate()
                except Exception:
                    pass
                atracacao_ini, atracacao_fim = self._obter_atracacao_report(ws_report)
                costs = ws_report.range("F24").value
                adm = ws_report.range("D25").value
                if nome_cliente.strip().upper() == "ROCHAMAR" and caminho_report.exists():
                    pdf_costs, pdf_adm = self._extrair_valores_report_pdf(caminho_report)
                    if pdf_costs is not None:
                        costs = pdf_costs
                    if pdf_adm is not None:
                        adm = pdf_adm
                criar_rascunho_email_cliente(
                    nome_cliente,
                    anexos=anexos,
                    dn=str(self.dn),
                    navio=self.nome_navio,
                    adiantamento=adiantamento,
                    atracacao_ini=atracacao_ini,
                    atracacao_fim=atracacao_fim,
                    costs=costs,
                    adm=adm,
                )
                print("✅ Rascunho do Outlook criado com anexos.")
            except Exception as e:
                print(f"⚠️ Nao foi possivel criar rascunho do Outlook: {e}")

            fechar_workbooks(self.app, self.wb1, self.wb2)

            print(f"💾 Excel salvo em: {caminho_excel}")
            print(f"📑 PDF FRONT salvo em: {caminho_pdf}")

        except Exception as e:
            print(f"❌ ERRO NO FATURAMENTO: {e}")
            fechar_workbooks(self.app, self.wb1, self.wb2)
            raise

    def _preview_title(self):
        return "Faturamento (Normal)"

    def _format_date(self, value):
        if isinstance(value, datetime):
            return value.strftime("%d/%m/%Y")
        if isinstance(value, date):
            return value.strftime("%d/%m/%Y")
        return str(value) if value is not None else ""

    def _format_brl(self, value):
        if value in (None, ""):
            return "R$ 0,00"
        try:
            num = float(value)
        except Exception:
            return str(value)

        texto = f"{num:,.2f}"
        texto = texto.replace(",", "X").replace(".", ",").replace("X", ".")
        return f"R$ {texto}"

    def _extrair_valores_report_pdf(self, caminho_pdf: Path):
        def parse_brl(txt: str) -> float | None:
            m = re.search(r"R\$\s*([0-9\.]+,[0-9]{2})", txt)
            if not m:
                return None
            try:
                return float(m.group(1).replace(".", "").replace(",", "."))
            except Exception:
                return None

        def extrair_valores_brl(txt: str) -> list[float]:
            vals = []
            for m in re.finditer(r"R\$\s*([0-9\.]+,[0-9]{2})", txt):
                try:
                    vals.append(float(m.group(1).replace(".", "").replace(",", ".")))
                except Exception:
                    continue
            return vals

        costs = None
        adm = None
        try:
            with pdfplumber.open(str(caminho_pdf)) as pdf:
                for page in pdf.pages:
                    texto = page.extract_text() or ""
                    if not texto:
                        continue
                    linhas = texto.splitlines()
                    custos_idx = None
                    for idx, line in enumerate(linhas):
                        line_up = line.upper()
                        if costs is None and "COST" in line_up:
                            costs = parse_brl(line)
                            if costs is not None:
                                custos_idx = idx
                        if adm is None and "ADM" in line_up:
                            adm = parse_brl(line)
                            if adm is None:
                                vals = extrair_valores_brl(line)
                                if vals:
                                    adm = vals[0]
                    if costs is not None and adm is None and custos_idx is not None:
                        for next_line in linhas[custos_idx + 1: custos_idx + 4]:
                            vals = extrair_valores_brl(next_line)
                            if vals:
                                adm = vals[0]
                                break
                    if costs is not None and adm is not None:
                        break
        except Exception:
            return None, None

        return costs, adm

    def _obter_atracacao_report(self, ws_report):
        datas = []
        valores = ws_report.range("C22:C500").value
        if not isinstance(valores, list):
            valores = [valores]
        for v in valores:
            if v in (None, ""):
                continue
            if isinstance(v, datetime):
                datas.append(v.date())
                continue
            if isinstance(v, date):
                datas.append(v)
                continue
            if isinstance(v, str):
                try:
                    d = datetime.strptime(v.strip(), "%d/%m/%Y").date()
                    datas.append(d)
                except Exception:
                    continue
        if not datas:
            return None, None
        return min(datas), max(datas)

    def _build_preview_text(self, total_periodos, nome_base):
        ws_report = self.wb2.sheets["REPORT VIGIA"]
        linhas = []
        linhas.append("PRE-VISUALIZACAO")
        linhas.append(f"Processo: {self._preview_title()}")
        linhas.append(f"Nome base: {nome_base}")
        linhas.append(f"DN: {self.dn}")
        linhas.append(f"Navio: {self.nome_navio}")
        linhas.append(f"Periodos: {total_periodos}")
        linhas.append("")
        linhas.append("DATA | PERIODO | VALOR")

        limite = min(total_periodos, 15)
        for i in range(limite):
            linha = 22 + i
            data = self._format_date(ws_report.range(f"C{linha}").value)
            periodo = ws_report.range(f"E{linha}").value or ""
            valor = self._format_brl(ws_report.range(f"G{linha}").value)
            linhas.append(f"{data} | {periodo} | {valor}")

        if total_periodos > limite:
            linhas.append(f"... {total_periodos - limite} linhas omitidas")

        return "\n".join(linhas)

    def _export_preview_pdf(self, nome_base):
        caminho_pdf = Path(gettempdir()) / f"preview_{nome_base}.pdf"
        if caminho_pdf.exists():
            caminho_pdf.unlink()

        try:
            self.wb2.app.calculation = "automatic"
            self.wb2.app.calculate()
        except Exception:
            pass

        aba_nf = None
        for ws in self.wb2.sheets:
            if ws.name.strip().upper() == "NF":
                aba_nf = ws
                ws.api.Visible = False
                break

        try:
            self.wb2.api.ExportAsFixedFormat(
                Type=0,
                Filename=str(caminho_pdf),
                Quality=0,
                IncludeDocProperties=True,
                IgnorePrintAreas=True,
                OpenAfterPublish=False,
            )
        finally:
            if aba_nf:
                aba_nf.api.Visible = True

        return caminho_pdf

    def processar_preview(self):
        # Front e report, sem gerar arquivos externos
        self.preencher_front_vigia()
        # ⚠️ NÃO atualiza planilha de controle no preview (evita duplicação)

        if "REPORT VIGIA" not in [s.name for s in self.wb2.sheets]:
            raise RuntimeError("Aba 'REPORT VIGIA' não encontrada")

        ws_report = self.wb2.sheets["REPORT VIGIA"]

        self.processar_MMO(self.wb1, self.wb2)

        qtd_periodos = self.obter_periodos(self.ws1)

        self.inserir_linhas_report(
            ws_report,
            linha_inicial=22,
            periodos=qtd_periodos
        )

        periodos = self.preencher_coluna_E(
            ws_report,
            linha_inicial=22,
            debug=True
        ) or []

        self.preencher_coluna_G(
            ws_report,
            self.ws1,
            linha_inicial=22,
            periodos=periodos,
            debug=False
        )

        self.montar_datas_report_vigia(
            ws_report,
            self.ws1,
            linha_inicial=22,
            periodos=len(periodos)
        )

        return len(periodos)


    def processar(self):
        # ---------- FRONT ----------
        self.preencher_front_vigia()

        # ---------- REPORT ----------
        if "REPORT VIGIA" not in [s.name for s in self.wb2.sheets]:
            raise RuntimeError("Aba 'REPORT VIGIA' não encontrada")

        ws_report = self.wb2.sheets["REPORT VIGIA"]

        self.processar_MMO(self.wb1, self.wb2)

        qtd_periodos = self.obter_periodos(self.ws1)

        self.inserir_linhas_report(
            ws_report,
            linha_inicial=22,
            periodos=qtd_periodos
        )

        periodos = self.preencher_coluna_E(
            ws_report,
            linha_inicial=22,
            debug=True
        )

        self.preencher_coluna_G(
            ws_report,
            self.ws1,
            linha_inicial=22,
            periodos=periodos,
            debug=True
        )

        self.montar_datas_report_vigia(
            ws_report,
            self.ws1,
            linha_inicial=22,
            periodos=len(periodos)
        )


        valor_arredondado = self.arredondar_para_baixo_50_se_cargonave()

        # 🔹 GERAR RECIBO CARGONAVE (Word + PDF)
        self.gerar_recibo_cargonave_word()


        # 🔹 GERAR PLANILHA DE CÁLCULO
        self.gerar_planilha_calculo_cargonave()

        self.gerar_planilha_calculo_conesul()

        print("✅ REPORT VIGIA atualizado com sucesso!")
        
        # ✅ Atualizar planilha de controle APÓS tudo estar pronto
        self.atualizar_planilha_controle()


    def escrever_cn_credit_note(self, texto_cn):
            ws_credit = None

            for sheet in self.wb2.sheets:
                if sheet.name.strip().lower() == "credit note":
                    ws_credit = sheet
                    break

            if ws_credit is None:
                print("ℹ️ Aba Credit Note não existe — seguindo fluxo.")
                return

            ws_credit.range("C21").value = texto_cn


    # ===== FRONT ======#
    def extrair_berco(self):
        """Extrai o valor do campo 'Berço' do PDF FOLHAS OGMO."""
        if not self.pdf_path or not Path(self.pdf_path).exists():
            print("⚠️ PDF FOLHAS OGMO não encontrado")
            return None

        with pdfplumber.open(str(self.pdf_path)) as pdf:
            for page in pdf.pages:
                words = page.extract_words()
                for w in words:
                    if w["text"] == "Berço":
                        x_ref = w["x0"]
                        y_ref = w["top"]

                        # pega palavras logo abaixo, alinhadas na mesma coluna
                        candidatos = [
                            wd for wd in words
                            if abs(wd["x0"] - x_ref) < 50 and wd["top"] > y_ref
                        ]

                        if candidatos:
                            candidatos.sort(key=lambda wd: wd["top"])
                            y_target = candidatos[0]["top"]

                            # junta todas as palavras dessa mesma linha
                            linha = [
                                wd["text"] for wd in candidatos
                                if abs(wd["top"] - y_target) < 5
                            ]
                            return " ".join(linha).strip()
        return None

    def preencher_front_vigia(self):
        try:
            ano_curto = datetime.now().strftime('%y')

            # FRONT VIGIA
            texto_dn = f"DN {self.dn}/{ano_curto}"
            self.ws_front.range("D15").value = self.nome_navio
            self.ws_front.range("C21").value = texto_dn

            # CREDIT NOTE (se existir)
            texto_cn = f"CN {self.dn}/{ano_curto}"
            self.escrever_cn_credit_note(texto_cn)


            # ======================

            # automatiza a leitura do BERÇO
            berco = self.extrair_berco()
            if berco:
                self.ws_front.range("D18").value = berco.upper()
            else:
                self.ws_front.range("D18").value = "NÃO ENCONTRADO"

            # -------- DATAS --------
            data_min, data_max = self.obter_datas_extremos(self.ws1)
            if data_min:
                self.ws_front.range("D16").value = self.data_por_extenso(data_min)
            if data_max:
                self.ws_front.range("D17").value = self.data_por_extenso(data_max)

            # -------- RODAPÉ --------
            hoje = datetime.now()
            meses = [
                "", "janeiro","fevereiro","março","abril","maio","junho",
                "julho","agosto","setembro","outubro","novembro","dezembro"
            ]
            self.ws_front.range("C39").value = (
                f"  Santos, {hoje.day} de {meses[hoje.month]} de {hoje.year}"
            )

            print("✅ FRONT VIGIA preenchido com sucesso!")

        except Exception as e:
            print(f"❌ Erro ao preencher FRONT VIGIA: {e}")
            raise

    def atualizar_planilha_controle(self):
        """
        Atualiza a planilha de controle com informações do faturamento VIGIA.
        Preenche colunas B (data), C (serviço), D (ETA), E (ETB), F (cliente), G (navio), J (DN), K (MMO/COSTS).
        """
        try:
            # Obter nome do cliente da pasta
            cliente = self.pasta_saida_final.parent.name.strip()
            
            # Obter data atual
            from datetime import datetime
            data_hoje = datetime.now().strftime("%d/%m/%Y")
            
            # Obter datas diretamente do RESUMO (NAVIO) em formato date
            data_min, data_max = self.obter_datas_extremos(self.ws1)
            
            # Formatar as datas como dd/mm/yyyy
            eta = data_min.strftime("%d/%m/%Y") if data_min else ""
            etb = data_max.strftime("%d/%m/%Y") if data_max else ""
            
            # Buscar valor de COSTS no REPORT VIGIA
            mmo = self._buscar_costs_report()
            
            # Usar CriarPasta para gravar na planilha
            criar_pasta = CriarPasta()
            criar_pasta._gravar_planilha(
                cliente=cliente,
                navio=self.nome_navio,
                dn=self.dn,
                servico="VIGIA",
                data=data_hoje,
                eta=eta,
                etb=etb,
                mmo=mmo
            )
            
        except Exception as e:
            print(f"⚠️ Erro ao atualizar planilha de controle: {e}")
    
    def _buscar_costs_report(self):
        """
        Busca o valor de COSTS no REPORT VIGIA dinamicamente.
        O valor do COSTS SEMPRE está na coluna G, independente de onde está a palavra 'COSTS'.
        Retorna formato brasileiro sem R$: 16.227,85
        """
        import re
        
        try:
            ws_report = self.wb2.sheets["REPORT VIGIA"]
            
            # Procura em um range amplo (linhas 1 a 150, todas as colunas)
            for linha in range(1, 151):
                for col_letra in ['C', 'D', 'E', 'F', 'G', 'H']:
                    try:
                        valor_celula = ws_report.range(f"{col_letra}{linha}").value
                        
                        # Verifica se contém COSTS ou MMO
                        if valor_celula and isinstance(valor_celula, str):
                            texto_upper = valor_celula.upper().strip()
                            
                            if "COSTS" in texto_upper or "MMO" in texto_upper:
                                # ✅ SEMPRE busca o valor na coluna G da MESMA linha
                                try:
                                    celula_g = ws_report.range(f"G{linha}")
                                    
                                    # Pega o TEXTO formatado (como aparece no Excel)
                                    valor_texto = None
                                    try:
                                        valor_texto = celula_g.api.Text
                                    except:
                                        valor_texto = str(celula_g.value) if celula_g.value else None
                                    
                                    if valor_texto and valor_texto.strip():
                                        # Remove formatação e converte
                                        try:
                                            valor_limpo = valor_texto.replace("R$", "").replace(" ", "").strip()
                                            # Se já está no formato brasileiro (1.234,56), converte
                                            if "," in valor_limpo:
                                                valor_limpo = valor_limpo.replace(".", "").replace(",", ".")
                                            valor_num = float(valor_limpo)
                                            
                                            # Retorna APENAS com vírgula decimal (sem ponto de milhar)
                                            resultado = f"{valor_num:.2f}".replace(".", ",")
                                            return resultado
                                        except Exception as e:
                                            # Se já estiver formatado, remove R$ e pontos
                                            texto_limpo = valor_texto.replace("R$", "").replace(".", "").strip()
                                            return texto_limpo
                                except:
                                    pass
                    except:
                        continue
            
            return ""
            
        except Exception as e:
            print(f"❌ Erro ao buscar COSTS: {e}")
            return ""


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

        # 2️⃣ Buscar periodos da primeira data (coluna C)
        data_ref = ws_resumo.range(f"B{primeira_linha_data}").value
        periodos_encontrados = []

        for i in range(primeira_linha_data, last_row + 1):
            valor_b = ws_resumo.range(f"B{i}").value
            valor_c = ws_resumo.range(f"C{i}").value

            if i != primeira_linha_data:
                if valor_b not in (None, "", "Total") and valor_b != data_ref:
                    break

            if isinstance(valor_c, str) and valor_c.strip().lower().startswith("total"):
                break

            periodo = self.normalizar_periodo(valor_c)
            if periodo:
                periodos_encontrados.append(periodo)

        if periodos_encontrados:
            # Escolhe o primeiro periodo pela ordem padrao (ignora a ordem da planilha)
            primeiro_periodo = min(
                periodos_encontrados,
                key=lambda p: sequencia_padrao.index(p),
            )
        else:
            primeiro_periodo = None

        # 3️⃣ Se nao for possivel, cai na heuristica de espacamento
        if not primeiro_periodo:
            contador_vazio = 0
            for i in range(primeira_linha_data + 1, last_row + 1):
                valor = ws_resumo.range(f"B{i}").value
                if valor in (None, "", "Total"):
                    contador_vazio += 1
                else:
                    break

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





    def gerar_pdf_report_vigia_separado(self, pasta_navio: Path, dn: str, navio: str):
        ws_report = self.wb2.sheets["REPORT VIGIA"]

        nome_pdf = f"REPORT VIGIA - DN {dn} - MV {navio}.pdf"
        caminho_pdf = pasta_navio / nome_pdf

        ws_report.api.ExportAsFixedFormat(
            Type=0,
            Filename=str(caminho_pdf),
            Quality=0,
            IncludeDocProperties=True,
            IgnorePrintAreas=False,
            OpenAfterPublish=False
        )

        print(f"📑 PDF REPORT VIGIA salvo em: {caminho_pdf}")



    def processar_MMO(self, wb_navio, wb_cliente):
        # ---------- REPORT VIGIA (CLIENTE) ----------
        try:
            ws_report = wb_cliente.sheets["REPORT VIGIA"]
        except:
            return

        if str(ws_report.range("E25").value).strip().upper() != "MMO":
            return

        # ---------- RESUMO (NAVIO) ----------
        try:
            ws_resumo = wb_navio.sheets["Resumo"]
        except:
            return

        # ---------- LÊ COLUNA G ----------
        valores = ws_resumo.range("G1:G1000").value
        valores_validos = [v for v in valores if v not in (None, "")]

        if not valores_validos:
            return

        ultimo_valor = valores_validos[-1]

        # ---------- CONVERSÃO CORRETA (IGUAL COLUNA G) ----------
        try:
            valor_float = float(ultimo_valor)
        except:
            print(f"   ⚠️ Valor inválido '{ultimo_valor}'. Usando 0.")
            valor_float = 0.0

        # 🔥 correção de escala (quando vem gigante)
        if valor_float > 1_000_000:
            valor_float = valor_float / 100

        # ---------- ESCREVE ----------
        celula = ws_report.range("F25")
        celula.value = valor_float
        celula.api.NumberFormatLocal = 'R$ #.##0,00'

        print(f"   ✅ MMO concluído → R$ {valor_float:,.2f}")



    def arredondar_para_baixo_50_se_cargonave(self):
        """
        Arredonda para baixo em múltiplos de 50.
        Somente para A/C AGÊNCIA MARÍTIMA CARGONAVE LTDA.
        Coloca o resultado em H28 do FRONT.
        """
        ws_front_vigia = self.ws_front
        valor_empresa = ws_front_vigia.range("C9").value
        if not valor_empresa:
            return

        if str(valor_empresa).strip().upper() != "A/C AGÊNCIA MARÍTIMA CARGONAVE LTDA.":
            return

        valor = ws_front_vigia.range("E37").value
        if valor is None:
            return

        try:
            resultado = (int(valor) // 50) * 50
        except (ValueError, TypeError):
            return

        ws_front_vigia.range("H28").value = resultado

        # Para gerar o Word, você pode usar esse mesmo valor:
        return resultado


    def gerar_recibo_cargonave_word(self):

        word = None
        doc = None

        try:
            # ==========================
            # 🔒 TRAVA DE SEGURANÇA
            # ==========================
            ws = self.ws_front

            empresa = ws.range("C9").value
            if not empresa or str(empresa).strip().upper() != "A/C AGÊNCIA MARÍTIMA CARGONAVE LTDA.":
                print("ℹ️ Recibo não gerado (empresa não é CARGONAVE).")
                return

            valor_h28 = ws.range("H28").value
            if valor_h28 in (None, "", 0):
                print("ℹ️ Recibo não gerado (adiantamento não acionado ou valor zero).")
                return

            # ==========================
            # 📄 MODELO WORD
            # ==========================
            pasta_modelos = self.pasta_faturamentos / "CARGONAVE"
            modelos = list(pasta_modelos.glob("RECIBO - YUTA.doc"))

            if not modelos:
                print(f"❌ Modelo Word não encontrado em {pasta_modelos}")
                return

            modelo_word = modelos[0]

            # ==========================
            # 📂 COPIAR PARA TEMP
            # ==========================
            temp_doc = Path(tempfile.gettempdir()) / f"RECIBO - {self.dn}.doc"
            shutil.copy2(modelo_word, temp_doc)

            # ==========================
            # 📝 ABRIR WORD
            # ==========================
            word = comtypes.client.CreateObject("Word.Application")
            word.Visible = False
            doc = word.Documents.Open(str(temp_doc))

            # ==========================
            # 💰 VALOR
            # ==========================
            valor = float(valor_h28)

            hoje = datetime.now()
            meses = [
                "", "janeiro", "fevereiro", "março", "abril", "maio", "junho",
                "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"
            ]

            data_extenso = f"Santos, {hoje.day} de {meses[hoje.month].capitalize()} de {hoje.year}"

            placeholders = {
                "{{VALOR}}": f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."),
                "{{VALOR_EXTENSO}}": num2words(valor, lang="pt_BR") + " reais",
                "{{DN}}": self.dn,
                "{{NAVIO}}": self.nome_navio,
                "{{DATA}}": data_extenso,
            }

            # ==========================
            # 🔁 SUBSTITUIR (TUDO NEGRITO)
            # ==========================
            find = doc.Content.Find

            for key, val in placeholders.items():
                find.ClearFormatting()
                find.Replacement.ClearFormatting()

                find.Text = key
                find.Replacement.Text = str(val)

                # 👉 FORÇA NEGRITO SEM EXCEÇÃO
                find.Replacement.Font.Bold = True

                find.Forward = True
                find.Wrap = 1  # wdFindContinue
                find.MatchCase = False
                find.MatchWholeWord = False
                find.Execute(Replace=2)

                print(f"🔄 Substituído {key} → {val} (NEGRITO)")

            # ==========================
            # 💾 SALVAR WORD + PDF
            # ==========================
            word_saida = self.pasta_saida_final / f"RECIBO - DN {self.dn} - MV {self.nome_navio}.doc"
            doc.SaveAs(str(word_saida))
            print(f"💾 Word do recibo salvo em: {word_saida}")

            pdf_saida = word_saida.with_suffix(".pdf")
            doc.SaveAs(str(pdf_saida), FileFormat=17)
            print(f"📑 PDF do recibo salvo em: {pdf_saida}")

            doc.Close(False)
            word.Quit()

        except Exception as e:
            if doc:
                try:
                    doc.Close(False)
                except:
                    pass
            if word:
                try:
                    word.Quit()
                except:
                    pass
            print(f"❌ Erro ao gerar recibo CARGONAVE: {e}")

    def gerar_planilha_calculo_cargonave(self):
        try:
            # ==========================
            # 🔒 TRAVA DE SEGURANÇA
            # ==========================
            ws = self.ws_front
            empresa = ws.range("C9").value

            if not empresa or str(empresa).strip().upper() != "A/C AGÊNCIA MARÍTIMA CARGONAVE LTDA.":
                print("ℹ️ Planilha de cálculo não gerada (empresa não é CARGONAVE).")
                return


            # ==========================
            # 🔤 FUNÇÃO AUXILIAR
            # ==========================
            def remover_acentos(txt: str) -> str:
                return unicodedata.normalize("NFD", txt).encode("ascii", "ignore").decode("utf-8")

            # ==========================
            # 📂 PASTA DO MODELO (BASE)
            # ==========================
            # ✅ PASTA DO MODELO (dinâmica - funciona no cliente)
            pasta_modelo = self.encontrar_pasta_modelo("CARGONAVE")


            if not pasta_modelo.exists():
                raise FileNotFoundError(f"Pasta modelo não encontrada: {pasta_modelo}")

            # ==========================
            # 📂 PASTA DO NAVIO (DESTINO)
            # ==========================
            pasta_navio = self.pasta_saida_final
            pasta_navio.mkdir(parents=True, exist_ok=True)

            # ==========================
            # 🔎 LOCALIZAR MODELO EXCEL
            # ==========================
            modelo = None

            for arq in pasta_modelo.glob("*.xlsx"):
                nome_limpo = remover_acentos(arq.name.lower())
                if "calculo" in nome_limpo:
                    modelo = arq
                    break

            if not modelo:
                raise FileNotFoundError(
                    f"Nenhum modelo de cálculo encontrado em {pasta_modelo}"
                )


            # ==========================
            # 📄 COPIAR MODELO
            # ==========================
            destino = pasta_navio / "CALCULO - YUTA.xlsx"
            shutil.copy2(modelo, destino)

            # ==========================
            # 📊 ABRIR PLANILHA
            # ==========================
            wb = openpyxl.load_workbook(destino)
            ws = wb.active  # ou wb["Cálculo"] se quiser fixar

            # ==========================
            # 📥 PEGAR ÚLTIMA LINHA DO OGMO
            # ==========================
            ws1 = self.ws1  # ✔ CONFIRMADO no teu fluxo

            ultima_linha = self.ultima_linha_com_valor(ws1, "G")



            print(f"📊 Última linha detectada no NAVIO: {ultima_linha}")

            mapa = {
                "C5": "G",
                "D5": "H",
                "E5": "I",
                "F5": "N",
                "G5": "O",
                "H5": "P",
                "I5": "Q",
                "J5": "S",
                "K5": "T",
                "L5": "V",
                "M5": "Z",
            }

            for destino_cell, origem_col in mapa.items():
                valor = ws1[f"{origem_col}{ultima_linha}"].value
                ws[destino_cell] = valor
                print(f"   🔹 {origem_col}{ultima_linha} → {destino_cell} | Valor: {valor}")

            # ==========================
            # ➕ CAMPOS ADICIONAIS
            # ==========================

            # AA (última linha OGMO) → B3
            valor_aa = ws1[f"AA{ultima_linha}"].value
            ws["B3"] = valor_aa
            print(f"   🔹 AA{ultima_linha} → B3 | Valor: {valor_aa}")

            # ==========================
            # 🚢 NOME DO NAVIO
            # ==========================
            nome_navio = self.nome_navio  # ajuste se o atributo tiver outro nome

            ws["A4"] = nome_navio

            print(f"   🔹 NAVIO → A2 e A4 | Valor: {nome_navio}")





            # ==========================
            # 💾 SALVAR
            # ==========================
            wb.save(destino)

            print("✅ Planilha CÁLCULO CARGONAVE gerada com sucesso!")

        except Exception as e:
            print(f"❌ Erro ao gerar planilha CÁLCULO CARGONAVE: {e}")
            raise




    def ultima_linha_com_valor(self, ws, coluna):
        for linha in range(ws.used_range.last_cell.row, 0, -1):
            if ws[f"{coluna}{linha}"].value not in (None, ""):
                return linha
        return None



    from pathlib import Path

    def encontrar_pasta_modelo(self, nome_cliente: str) -> Path:
        """
        Encontra ...\01. FATURAMENTOS\<nome_cliente> usando como base
        as pastas que já funcionam no PC atual.
        """
        bases = []
        if getattr(self, "pasta_saida_final", None):
            bases.append(Path(self.pasta_saida_final))
        if getattr(self, "pasta_faturamentos", None):
            bases.append(Path(self.pasta_faturamentos))

        for base in bases:
            for p in [base] + list(base.parents):
                nome_pasta = p.name.strip().upper()
                if "01. FATURAMENTOS" in nome_pasta:
                    pasta = p / nome_cliente
                    if pasta.exists():
                        return pasta

        raise FileNotFoundError(
            f"Não encontrei a pasta de modelos em ...\\01. FATURAMENTOS\\{nome_cliente} "
            f"(base testada: {[str(b) for b in bases]})"
        )



    def gerar_planilha_calculo_conesul(self):
        try:
            # ==========================
            # 🔒 TRAVA DE SEGURANÇA
            # ==========================
            ws = self.ws_front
            empresa = ws.range("C9").value

            if not empresa or str(empresa).strip().upper() != "A/C CONE SUL AGÊNCIA DE NAVEGAÇÃO LTDA.":
                print("ℹ️ Planilha de cálculo não gerada (empresa não é CONESUL).")
                return


            # ==========================
            # 🔤 FUNÇÃO AUXILIAR
            # ==========================
            def remover_acentos(txt: str) -> str:
                return unicodedata.normalize("NFD", txt).encode("ascii", "ignore").decode("utf-8")

            # ==========================
            # 📂 PASTA DO MODELO (BASE)
            # ==========================
            pasta_modelo = self.encontrar_pasta_modelo("CONESUL")


            if not pasta_modelo.exists():
                raise FileNotFoundError(f"Pasta modelo não encontrada: {pasta_modelo}")

            # ==========================
            # 📂 PASTA DO NAVIO (DESTINO)
            # ==========================
            pasta_navio = self.pasta_saida_final
            pasta_navio.mkdir(parents=True, exist_ok=True)

            # ==========================
            # 🔎 LOCALIZAR MODELO EXCEL
            # ==========================
            modelo = None

            for arq in pasta_modelo.glob("*.xlsx"):
                nome_limpo = remover_acentos(arq.name.lower())
                if "calculo" in nome_limpo:
                    modelo = arq
                    break

            if not modelo:
                raise FileNotFoundError(
                    f"Nenhum modelo de cálculo encontrado em {pasta_modelo}"
                )


            # ==========================
            # 📄 COPIAR MODELO
            # ==========================
            destino = pasta_navio / "CALCULO - YUTA.xlsx"
            shutil.copy2(modelo, destino)

            # ==========================
            # 📊 ABRIR PLANILHA
            # ==========================
            wb = openpyxl.load_workbook(destino)
            ws = wb.active  # ou wb["Cálculo"] se quiser fixar

            # ==========================
            # 📥 PEGAR ÚLTIMA LINHA DO OGMO
            # ==========================
            ws1 = self.ws1  # ✔ CONFIRMADO no teu fluxo

            ultima_linha = self.ultima_linha_com_valor(ws1, "G")



            print(f"📊 Última linha detectada no NAVIO: {ultima_linha}")

            mapa = {
                "C5": "G",
                "D5": "H",
                "E5": "I",
                "F5": "N",
                "G5": "O",
                "H5": "P",
                "I5": "Q",
                "J5": "S",
                "K5": "T",
                "L5": "V",
                "M5": "Z",
            }

            for destino_cell, origem_col in mapa.items():
                valor = ws1[f"{origem_col}{ultima_linha}"].value
                ws[destino_cell] = valor
                print(f"   🔹 {origem_col}{ultima_linha} → {destino_cell} | Valor: {valor}")

            # ==========================
            # ➕ CAMPOS ADICIONAIS
            # ==========================

            # AA (última linha OGMO) → B3
            valor_aa = ws1[f"AA{ultima_linha}"].value
            ws["B3"] = valor_aa
            print(f"   🔹 AA{ultima_linha} → B3 | Valor: {valor_aa}")

            # ==========================
            # 🚢 NOME DO NAVIO
            # ==========================
            nome_navio = self.nome_navio  # ajuste se o atributo tiver outro nome

            ws["A4"] = nome_navio

            print(f"   🔹 NAVIO → A2 e A4 | Valor: {nome_navio}")





            # ==========================
            # 💾 SALVAR
            # ==========================
            wb.save(destino)

            print("✅ Planilha CÁLCULO CONESUL gerada com sucesso!")

        except Exception as e:
            print(f"❌ Erro ao gerar planilha CÁLCULO CONESUL: {e}")
            raise




    def obter_valor_cargonave(self):
        """
        Retorna o valor do adiantamento CARGONAVE
        (lido direto do FRONT – célula H28)
        """
        valor = self.ws_front.range("H28").value
        try:
            return float(valor)
        except:
            return 0.0


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
