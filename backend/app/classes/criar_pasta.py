from pathlib import Path
import re
import os
import time
import unicodedata
from datetime import datetime

from yuta_helpers import obter_pasta_faturamentos, openpyxl
from config_manager import obter_caminho_base_faturamentos


class CriarPasta:
    # Cache de clientes para otimização (evita múltiplas varreduras na rede)
    _cache_clientes = None
    _cache_timestamp = None
    _cache_ttl = 300  # 5 minutos
    
    # Cache para próximo DN
    _cache_proximo_dn = None
    _cache_dn_timestamp = None
    
    def __init__(self, planilha_nome="CONTROLE DE FATURAMENTO"):
        self.planilha_nome = planilha_nome

    def _ano_atual_2d(self) -> str:
        return f"{datetime.now().year % 100:02d}"

    def _possiveis_desktops(self):
        home = Path.home()
        return [
            home / "Desktop",
            home / "OneDrive" / "Desktop",
            home / "Area de Trabalho",
            home / "OneDrive" / "Area de Trabalho",
        ]

    def _deve_usar_fallback_desktop(self, bases_rede: list[Path]) -> bool:
        """
        Decide se pode usar Desktop/Área de Trabalho para localizar planilha.

        Regras:
        - Se YUTA_ALLOW_DESKTOP_FALLBACK=1/true/yes/on: permite sempre
        - Caso contrário, só permite quando nenhuma base de rede está acessível
        """
        flag = os.getenv("YUTA_ALLOW_DESKTOP_FALLBACK", "").strip().lower()
        if flag in {"1", "true", "yes", "on"}:
            return True

        for base in bases_rede:
            try:
                if base.exists() and base.is_dir():
                    return False
            except Exception:
                continue

        return True

    def _possiveis_bases_clientes(self):
        """
        Retorna lista de possíveis bases. Agora usa o sistema de configuração.
        """
        try:
            # Tenta usar o caminho configurado/auto-detectado
            caminho_config = obter_caminho_base_faturamentos()
            return [caminho_config]
        except FileNotFoundError:
            # Fallback para detecção manual
            home = Path.home()
            return [
                home / "SANPORT LOGÍSTICA PORTUÁRIA LTDA" / "Central de Documentos - 01. FATURAMENTOS",
                home / "SANPORT LOGÍSTICA PORTUÁRIA LTDA" / "Central de Documentos - Documentos" / "01. FATURAMENTOS",
                home
                / "OneDrive - SANPORT LOGÍSTICA PORTUÁRIA LTDA"
                / "Central de Documentos - 01. FATURAMENTOS",
                home
                / "OneDrive - SANPORT LOGÍSTICA PORTUÁRIA LTDA"
                / "Central de Documentos - Documentos"
                / "01. FATURAMENTOS",
            ]

    def _obter_base_clientes(self) -> Path:
        """
        Obtém a pasta base onde ficam as pastas dos clientes.
        Usa o sistema de configuração centralizado.
        """
        try:
            return obter_caminho_base_faturamentos()
        except FileNotFoundError:
            # Fallback: tenta os métodos antigos
            for base in self._possiveis_bases_clientes():
                if base.exists():
                    return base
            pasta_faturamentos = obter_pasta_faturamentos()
            return pasta_faturamentos.parent

    def _encontrar_planilha(self) -> Path:
        extensoes = [".xlsx", ".xlsm", ".xls"]
        bases_rede = self._possiveis_bases_clientes()
        bases_desktop = self._possiveis_desktops() if self._deve_usar_fallback_desktop(bases_rede) else []
        bases = bases_rede + bases_desktop

        ano_atual_4d = str(datetime.now().year)
        ano_atual_2d = self._ano_atual_2d()

        # 1) Prioriza nomes explícitos do ano atual (ex.: "CONTROLE DE FATURAMENTO 2026")
        for base in bases:
            for ext in extensoes:
                nomes_prioritarios = [
                    f"{self.planilha_nome} {ano_atual_4d}{ext}",
                    f"{self.planilha_nome}-{ano_atual_4d}{ext}",
                    f"{self.planilha_nome}_{ano_atual_4d}{ext}",
                    f"{self.planilha_nome} {ano_atual_2d}{ext}",
                    f"{self.planilha_nome}-{ano_atual_2d}{ext}",
                    f"{self.planilha_nome}_{ano_atual_2d}{ext}",
                ]
                for nome in nomes_prioritarios:
                    caminho = base / nome
                    if caminho.exists():
                        return caminho

        # 2) Fallback para nome padrão sem sufixo
        for base in bases:
            for ext in extensoes:
                caminho = base / f"{self.planilha_nome}{ext}"
                if caminho.exists():
                    return caminho

        candidatos = []
        nome_ref = self._normalizar_pasta_nome(self.planilha_nome)

        for base in bases:
            if not base.exists() or not base.is_dir():
                continue

            try:
                for arquivo in base.iterdir():
                    if not arquivo.is_file():
                        continue
                    if arquivo.suffix.lower() not in extensoes:
                        continue

                    stem_norm = self._normalizar_pasta_nome(arquivo.stem)
                    if nome_ref and nome_ref in stem_norm:
                        contem_ano_atual = ano_atual_4d in stem_norm or ano_atual_2d in stem_norm
                        prioridade_ano = 0 if contem_ano_atual else 1
                        prioridade_nome_exato = 0 if stem_norm == nome_ref else 1
                        prioridade_origem = 0 if base in bases_rede else 1
                        candidatos.append((prioridade_origem, prioridade_ano, prioridade_nome_exato, -arquivo.stat().st_mtime, arquivo))
            except Exception:
                continue

        if candidatos:
            candidatos.sort(key=lambda item: (item[0], item[1], item[2], item[3]))
            return candidatos[0][4]

        locais = "\n".join(f"- {b}" for b in bases)
        raise FileNotFoundError(
            f"Planilha de controle nao encontrada.\n"
            f"Nome base procurado: '{self.planilha_nome}'\n"
            f"Locais verificados:\n{locais}"
        )

    def _ultima_linha_com_dados(self, ws, colunas):
        """
        Busca a última linha com dados. Para evitar gaps, prioriza a coluna J (DN).
        """
        ultima = 0
        
        # Prioriza coluna J (DN) que sempre tem dados
        if "J" in colunas:
            for row in range(ws.max_row, 0, -1):
                valor = ws[f"J{row}"].value
                if valor and (isinstance(valor, (int, float)) or (isinstance(valor, str) and valor.strip())):
                    return row
        
        # Fallback: busca em outras colunas
        for col in colunas:
            for row in range(ws.max_row, 0, -1):
                valor = ws[f"{col}{row}"].value
                if valor is None:
                    continue
                if isinstance(valor, str) and not valor.strip():
                    continue
                ultima = max(ultima, row)
                break
        
        if ultima == 0:
            raise RuntimeError("Nao foi possivel localizar dados nas colunas")
        return ultima

    def _normalizar_texto(self, valor):
        if valor is None:
            return ""
        return str(valor).strip()

    def _normalizar_pasta_nome(self, valor):
        texto = self._normalizar_texto(valor).upper()
        texto = unicodedata.normalize("NFKD", texto)
        texto = "".join(ch for ch in texto if not unicodedata.combining(ch))
        texto = re.sub(r"\([^)]*\)", "", texto)
        texto = re.sub(r"[^A-Z0-9]+", "", texto)
        return texto

    def listar_clientes(self, forcar_refresh=False):
        """
        Lista clientes disponíveis com cache para otimizar acesso à rede.
        
        Args:
            forcar_refresh: Se True, ignora cache e busca novamente
        """
        import time
        
        # Verifica se tem cache válido
        if not forcar_refresh and self._cache_clientes is not None and self._cache_timestamp is not None:
            tempo_decorrido = time.time() - self._cache_timestamp
            if tempo_decorrido < self._cache_ttl:
                return self._cache_clientes
        
        # Cache expirado ou refresh forçado: busca novamente
        base = self._obter_base_clientes()
        clientes = []
        
        try:
            for item in base.iterdir():
                if not item.is_dir():
                    continue
                if item.name.upper() == "FATURAMENTOS":
                    continue
                clientes.append(item.name)
        except Exception as e:
            # Se falhar (rede indisponível, etc), retorna cache antigo se existir
            if self._cache_clientes is not None:
                return self._cache_clientes
            raise
        
        resultado = sorted(clientes, key=lambda v: v.casefold())
        
        # Atualiza cache
        CriarPasta._cache_clientes = resultado
        CriarPasta._cache_timestamp = time.time()
        
        return resultado
    
    def obter_proximo_dn(self, forcar_refresh=False) -> str:
        """
        Obtém o próximo DN da sequência (último DN + 1)
        Com cache para evitar abrir planilha toda vez.
        
        Args:
            forcar_refresh: Se True, ignora cache e busca novamente
        """
        import time
        
        # Verifica cache válido
        if not forcar_refresh and self._cache_proximo_dn is not None and self._cache_dn_timestamp is not None:
            tempo_decorrido = time.time() - self._cache_dn_timestamp
            if tempo_decorrido < self._cache_ttl:
                return self._cache_proximo_dn
        
        try:
            caminho_planilha = self._encontrar_planilha()
            wb = openpyxl.load_workbook(caminho_planilha, data_only=True)
            ws = wb.active
            
            # Encontra a última linha com dados na coluna J
            ultima_linha = self._ultima_linha_com_dados(ws, ["J"])
            ano_atual = self._ano_atual_2d()
            
            if ultima_linha < 2:  # Se não há dados, começa do 1
                resultado = f"001/{ano_atual}"
            else:
                ultimo_numero = 0
                ano_sufixo = None

                for row in range(ultima_linha, 0, -1):
                    valor = ws[f"J{row}"].value
                    numero, ano_lido = self._extrair_numero_ano_dn(valor)
                    if numero > 0:
                        ultimo_numero = numero
                        ano_sufixo = ano_lido
                        break

                if ultimo_numero <= 0:
                    ultimo_numero = self._obter_maior_numero_dn_por_pastas()

                proximo = ultimo_numero + 1 if ultimo_numero > 0 else 1
                ano_final = ano_sufixo or ano_atual
                resultado = f"{str(proximo).zfill(3)}/{ano_final}"
            
            # Atualiza cache
            CriarPasta._cache_proximo_dn = resultado
            CriarPasta._cache_dn_timestamp = time.time()
            
            return resultado
            
        except FileNotFoundError:
            # Fallback: tenta inferir pela estrutura de pastas da rede
            ultimo_numero = self._obter_maior_numero_dn_por_pastas()
            proximo = ultimo_numero + 1 if ultimo_numero > 0 else 1
            return f"{str(proximo).zfill(3)}/{self._ano_atual_2d()}"
        except Exception as e:
            print(f"⚠️ Erro ao obter próximo DN: {e}")
            # Se falhar mas tem cache, usa cache antigo
            if self._cache_proximo_dn is not None:
                return self._cache_proximo_dn
            ultimo_numero = self._obter_maior_numero_dn_por_pastas()
            proximo = ultimo_numero + 1 if ultimo_numero > 0 else 1
            return f"{str(proximo).zfill(3)}/{self._ano_atual_2d()}"  # Fallback

    def _resolver_pasta_cliente(self, pasta_base: Path, cliente: str) -> Path | None:
        alvo = self._normalizar_pasta_nome(cliente)
        if not alvo:
            return None

        base_nome = re.sub(r"\s*\([^)]*\)\s*", " ", cliente).strip()
        base_norm = self._normalizar_pasta_nome(base_nome)
        hint_match = re.search(r"\(([^)]*)\)", cliente)
        hint = self._normalizar_pasta_nome(hint_match.group(1)) if hint_match else ""

        # 1) Tentativa direta
        direta = pasta_base / cliente
        if direta.exists():
            return direta

        # 2) Match por normalizacao
        candidatos = []
        for item in pasta_base.iterdir():
            if not item.is_dir():
                continue
            norm_item = self._normalizar_pasta_nome(item.name)
            if norm_item == alvo:
                candidatos.append(item)

        if len(candidatos) == 1:
            return candidatos[0]

        # 3) Match parcial pelo nome base (sem parenteses)
        if base_norm:
            candidatos = []
            for item in pasta_base.iterdir():
                if not item.is_dir():
                    continue
                norm_item = self._normalizar_pasta_nome(item.name)
                if base_norm and base_norm in norm_item:
                    candidatos.append(item)

            if len(candidatos) == 1:
                return candidatos[0]

            # 4) Desempate usando hint (ex: PSS)
            if hint and candidatos:
                filtrados = []
                for item in candidatos:
                    norm_item = self._normalizar_pasta_nome(item.name)
                    if hint in norm_item:
                        filtrados.append(item)

                if len(filtrados) == 1:
                    return filtrados[0]

                if hint == "PSS":
                    filtrados = []
                    for item in candidatos:
                        norm_item = self._normalizar_pasta_nome(item.name)
                        if "SAOSEBASTIAO" in norm_item:
                            filtrados.append(item)
                    if len(filtrados) == 1:
                        return filtrados[0]

        return None

    def _formatar_numero(self, valor):
        if valor is None:
            return ""
        if isinstance(valor, (int, float)):
            if isinstance(valor, float) and not valor.is_integer():
                return str(valor).strip()
            return str(int(valor)).zfill(3)
        texto = str(valor).strip()
        if "/" in texto:
            texto = texto.split("/", 1)[0].strip()
        if texto.isdigit() and len(texto) < 3:
            return texto.zfill(3)
        return texto

    def _padronizar_dn(self, dn: str) -> str:
        texto = str(dn).strip()
        if "/" not in texto:
            return f"{texto}/{self._ano_atual_2d()}"
        return texto

    def _extrair_numero_ano_dn(self, valor) -> tuple[int, str | None]:
        if valor is None:
            return 0, None

        texto = str(valor).strip()
        if not texto:
            return 0, None

        match = re.search(r"(\d+)\s*/\s*(\d{2,4})", texto)
        if match:
            numero = int(match.group(1))
            ano_bruto = match.group(2)
            ano_2d = ano_bruto[-2:]
            return numero, ano_2d

        numeros = re.findall(r"\d+", texto)
        if numeros:
            return int(numeros[0]), None

        return 0, None

    def _obter_maior_numero_dn_por_pastas(self) -> int:
        """
        Fallback para quando a planilha de controle não puder ser lida.
        Busca o maior número no início das pastas de navio.
        Ex: "241 - EAGLE ARROW" -> 241.
        """
        try:
            pasta_base = self._obter_base_clientes()
        except Exception:
            return 0

        maior = 0

        try:
            for pasta_cliente in pasta_base.iterdir():
                if not pasta_cliente.is_dir():
                    continue
                if pasta_cliente.name.upper() == "FATURAMENTOS":
                    continue

                try:
                    for pasta_navio in pasta_cliente.iterdir():
                        if not pasta_navio.is_dir():
                            continue

                        match = re.match(r"^\s*(\d+)", pasta_navio.name)
                        if not match:
                            continue

                        numero = int(match.group(1))
                        if numero > maior:
                            maior = numero
                except Exception:
                    continue
        except Exception:
            return 0

        return maior

    def salvar_planilha_com_retry(self, wb, caminho_planilha: Path, tentativas: int = 5, espera_inicial: float = 0.8):
        """
        Salva workbook com retry para ambiente de rede (arquivo em uso por outro PC).
        """
        erro_final = None

        for tentativa in range(1, tentativas + 1):
            try:
                wb.save(caminho_planilha)
                return
            except PermissionError as e:
                erro_final = e
            except OSError as e:
                erro_final = e

            if tentativa < tentativas:
                espera = espera_inicial * (2 ** (tentativa - 1))
                print(
                    f"⚠️ Planilha de controle em uso ou indisponível na rede "
                    f"(tentativa {tentativa}/{tentativas}). Nova tentativa em {espera:.1f}s..."
                )
                time.sleep(espera)

        raise RuntimeError(
            f"Não foi possível salvar a planilha de controle após {tentativas} tentativas: {caminho_planilha}\n"
            f"Erro final: {erro_final}"
        )

    def _gravar_planilha(self, cliente: str, navio: str, dn: str, servico: str = None, data: str = None, eta: str = None, etb: str = None, mmo: str = None, mmo_extra: str = None, adiantamento: str = None, wb_externo=None, iss: str = None, limpar_formulas_adm_cliente: bool = False, iss_formula: bool = False):
        """
        Grava informações na planilha de controle.
        
        Args:
            cliente: Nome do cliente (coluna F)
            navio: Nome do navio (coluna G)
            dn: DN (coluna J)
            servico: Tipo de serviço - "VIGIA" ou "DE ACORDO" (coluna C)
            data: Data do dia (coluna B)
            eta: Data inicial - D16 da FRONT VIGIA (coluna D)
            etb: Data final - D17 da FRONT VIGIA (coluna E)
            mmo: Valor principal para coluna K (DESPESAS/COSTS)
            mmo_extra: Valor MMO para coluna L (quando aplicável)
            adiantamento: Valor de adiantamento para coluna H (CARGONAVE)
            wb_externo: Workbook openpyxl já aberto (opcional, evita reabrir)
            iss: Valor do ISS (coluna O)
            limpar_formulas_adm_cliente: Se True, limpa fórmulas das colunas N (ADM %) e P (CLIENTE %)
            iss_formula: Se True, cria fórmula =K{linha}*5% na coluna O ao invés de valor fixo
        """
        caminho_planilha = self._encontrar_planilha()
        
        # ✅ Reutiliza workbook se fornecido
        if wb_externo is not None:
            wb = wb_externo
            deve_fechar = False
        else:
            wb = openpyxl.load_workbook(caminho_planilha)
            deve_fechar = True
        
        ws = wb.active

        dn_padronizado = self._padronizar_dn(dn)
        
        # Busca eficiente: carrega todos DNs de uma vez
        ultima_linha = self._ultima_linha_com_dados(ws, ["A", "B", "C", "D", "E", "F", "G", "J", "K", "M"])
        linha = None
        
        # Busca o DN na coluna J
        for row in range(1, ultima_linha + 1):
            dn_cell = ws[f"J{row}"].value
            if dn_cell and str(dn_cell).strip() == dn_padronizado:
                linha = row
                break
        
        # Se não encontrou, cria nova linha
        if linha is None:
            linha = ultima_linha + 1
            print(f"📝 Criando nova linha {linha} na planilha de controle")
            # Apenas na CRIAÇÃO, preenche cliente/navio/DN
            ws[f"F{linha}"].value = cliente
            ws[f"G{linha}"].value = navio
            ws[f"J{linha}"].value = dn_padronizado
            
            # Preenche coluna M (NF sequencial - próximo número disponível)
            ultimo_nf = None
            for row in range(ultima_linha, 0, -1):
                valor_m = ws[f"M{row}"].value
                if valor_m and str(valor_m).strip():
                    try:
                        ultimo_nf = int(str(valor_m).strip())
                        break
                    except:
                        continue
            
            # Se encontrou um número, incrementa; senão começa do 7986
            proximo_nf = (ultimo_nf + 1) if ultimo_nf else 7986
            ws[f"M{linha}"].value = proximo_nf
        else:
            print(f"📝 Atualizando linha {linha} existente (DN: {dn_padronizado})")

        # Atualiza APENAS os campos fornecidos (não sobrescreve vazios)
        def _to_float_moeda(valor):
            if valor in (None, ""):
                return None

            if isinstance(valor, (int, float)):
                return float(valor)

            texto = str(valor).strip().replace("R$", "").replace(" ", "")
            if not texto:
                return None

            if "," in texto:
                # Formato BR: 12.345,67
                texto = texto.replace(".", "").replace(",", ".")
            else:
                # Formato EN ou inteiro: 12345.67 / 12345
                if texto.count(".") > 1:
                    texto = texto.replace(".", "")

            return float(texto)

        if data:
            ws[f"B{linha}"].value = data
            
            # Preenche coluna A com mês abreviado (JAN, FEV, MAR...)
            try:
                from datetime import datetime
                # Tenta converter a data string para datetime
                if isinstance(data, str) and "/" in data:
                    data_obj = datetime.strptime(data, "%d/%m/%Y")
                elif hasattr(data, 'month'):
                    data_obj = data
                else:
                    data_obj = None
                
                if data_obj:
                    meses_abrev = {
                        1: "JAN", 2: "FEV", 3: "MAR", 4: "ABR",
                        5: "MAI", 6: "JUN", 7: "JUL", 8: "AGO",
                        9: "SET", 10: "OUT", 11: "NOV", 12: "DEZ"
                    }
                    mes_abrev = meses_abrev.get(data_obj.month, "")
                    if mes_abrev:
                        ws[f"A{linha}"].value = mes_abrev
            except:
                pass
        
        if servico:
            ws[f"C{linha}"].value = servico
        if eta:
            ws[f"D{linha}"].value = eta
        if etb:
            ws[f"E{linha}"].value = etb

        if adiantamento is not None:
            celula_adiantamento = ws[f"H{linha}"]
            try:
                valor_numero = _to_float_moeda(adiantamento)
                celula_adiantamento.value = valor_numero
                celula_adiantamento.number_format = '"R$ "#,##0.00'
                print(f"✓ Adiantamento gravado na coluna H, linha {linha}: {valor_numero}")
            except Exception as e:
                celula_adiantamento.value = str(adiantamento)
                print(f"⚠️ Adiantamento gravado como texto: {adiantamento} (erro: {e})")

        if mmo is not None:
            # MMO: converter para número e formatar como moeda
            celula_mmo = ws[f"K{linha}"]
            try:
                valor_numero = _to_float_moeda(mmo)
                
                # Grava como NÚMERO (não texto)
                celula_mmo.value = valor_numero
                
                # Formato de moeda brasileiro com R$
                celula_mmo.number_format = '"R$ "#,##0.00'
                print(f"✓ MMO gravado na coluna K, linha {linha}: {valor_numero}")
            except Exception as e:
                # Se falhar, grava como texto mesmo
                celula_mmo.value = str(mmo)
                print(f"⚠️ MMO gravado como texto: {mmo} (erro: {e})")

        if mmo_extra is not None:
            celula_mmo_extra = ws[f"L{linha}"]
            try:
                valor_numero = _to_float_moeda(mmo_extra)
                celula_mmo_extra.value = valor_numero
                celula_mmo_extra.number_format = '"R$ "#,##0.00'
                print(f"✓ MMO extra gravado na coluna L, linha {linha}: {valor_numero}")
            except Exception as e:
                celula_mmo_extra.value = str(mmo_extra)
                print(f"⚠️ MMO extra gravado como texto: {mmo_extra} (erro: {e})")
        
        if iss:
            # ISS: converter para número e formatar como moeda (coluna O)
            celula_iss = ws[f"O{linha}"]
            try:
                valor_numero = _to_float_moeda(iss)
                
                # Grava como NÚMERO (não texto)
                celula_iss.value = valor_numero
                
                # Formato de moeda brasileiro com R$
                celula_iss.number_format = '"R$ "#,##0.00'
            except:
                # Se falhar, grava como texto mesmo
                celula_iss.value = str(iss)
        elif iss_formula:
            # Cria fórmula =K{linha}*5% na coluna O (para DE ACORDO)
            celula_iss = ws[f"O{linha}"]
            celula_iss.value = f"=K{linha}*5%"
            celula_iss.number_format = '"R$ "#,##0.00'
            print(f"✓ Fórmula ISS criada na coluna O, linha {linha}: =K{linha}*5%")
        
        # Limpar fórmulas das colunas N (ADM %) e P (CLIENTE %) para DE ACORDO
        if limpar_formulas_adm_cliente:
            ws[f"N{linha}"].value = None  # Limpa ADM %
            ws[f"P{linha}"].value = None  # Limpa CLIENTE %
            print(f"✓ Colunas N e P limpas (linha {linha})")

        # ✅ Só salva se abriu internamente (workbook externo é responsabilidade de quem passou)
        if deve_fechar:
            self.salvar_planilha_com_retry(wb, caminho_planilha)

    def executar(
        self,
        cliente: str | None = None,
        navio: str | None = None,
        dn: str | None = None,
        return_info: bool = False,
        log_callback=None,
        servico: str = None,
        data: str = None,
        eta: str = None,
        etb: str = None,
        mmo: str = None,
        wb_externo=None,
    ):
        def log(msg, tag="info"):
            if log_callback:
                log_callback(msg + "\n", tag=tag)
            else:
                print(msg)
        
        if cliente and navio and dn:
            log(f"📋 Cliente: {cliente}")
            log(f"🚢 Navio: {navio}")
            log(f"📝 DN: {dn}")
            
            numero = self._formatar_numero(dn)
        else:
            log(f"📋 Lendo dados da planilha...")
            caminho_planilha = self._encontrar_planilha()
            wb = openpyxl.load_workbook(caminho_planilha, data_only=True)
            ws = wb.active

            ultima_linha = self._ultima_linha_com_dados(ws, ["F", "G", "J"])
            cliente = self._normalizar_texto(ws[f"F{ultima_linha}"].value)
            navio = self._normalizar_texto(ws[f"G{ultima_linha}"].value)
            numero = self._formatar_numero(ws[f"J{ultima_linha}"].value)

        if not cliente or not navio or not numero:
            raise RuntimeError(
                f"Dados incompletos: cliente={cliente}, navio={navio}, numero={numero}"
            )

        log(f"🔍 Localizando pasta do cliente '{cliente}'...")
        pasta_base = self._obter_base_clientes()
        log(f"📂 Base: {pasta_base}")
        
        pasta_cliente = self._resolver_pasta_cliente(pasta_base, cliente)

        if not pasta_cliente:
            raise FileNotFoundError(
                f"Pasta do cliente não encontrada!\n"
                f"Base: {pasta_base}\n"
                f"Cliente: '{cliente}'"
            )

        log(f"📁 Pasta do cliente: {pasta_cliente.name}")

        nome_pasta = f"{numero} - {navio}"
        destino = pasta_cliente / nome_pasta
        log(f"📝 Criando: {nome_pasta}")
        
        destino.mkdir(parents=True, exist_ok=True)
        
        if not destino.exists():
            raise RuntimeError(f"Falha ao criar pasta: {destino}")

        log(f"✅ Pasta criada com sucesso!", tag="ok")
        log(f"   📍 {destino}")
        
        if return_info:
            return {
                "destino": destino,
                "pasta_cliente": pasta_cliente,
                "base": pasta_base,
            }
        return destino
