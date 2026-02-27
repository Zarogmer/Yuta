"""
Gerador de NFS-e - GINFES Santos/SP
Manual Técnico v1.3 (LC 214/2025)
v3: Consulta CNPJ via ReceitaWS + cache SQLite
Pré-requisito: pip install PyQt6
Uso: python gerador_nfse_ginfes_v3.py
"""

import sys
import sqlite3
import json
import urllib.request
import urllib.error
import xml.etree.ElementTree as ET
from datetime import datetime
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QFormLayout, QLineEdit, QTextEdit, QComboBox, QDoubleSpinBox,
    QSpinBox, QPushButton, QLabel, QGroupBox, QFileDialog,
    QMessageBox, QFrame, QScrollArea, QDateEdit, QTabWidget,
    QStatusBar, QProgressBar
)
from PyQt6.QtCore import Qt, QDate, QThread, pyqtSignal
from PyQt6.QtGui import QFont

# ──────────────────────────────────────────────
# ESTILO DARK/PROFISSIONAL
# ──────────────────────────────────────────────
STYLE = """
QMainWindow, QWidget {
    background-color: #1e1e2e;
    color: #cdd6f4;
    font-family: 'Segoe UI', 'Inter', sans-serif;
    font-size: 13px;
}
QGroupBox {
    border: 1px solid #313244;
    border-radius: 8px;
    margin-top: 12px;
    padding: 10px;
    font-weight: bold;
    color: #89b4fa;
    font-size: 13px;
}
QGroupBox::title {
    subcontrol-origin: margin;
    subcontrol-position: top left;
    padding: 0 6px;
}
QLineEdit, QTextEdit, QComboBox, QDoubleSpinBox, QSpinBox, QDateEdit {
    background-color: #313244;
    border: 1px solid #45475a;
    border-radius: 6px;
    padding: 6px 10px;
    color: #cdd6f4;
    selection-background-color: #89b4fa;
}
QLineEdit:focus, QTextEdit:focus, QComboBox:focus,
QDoubleSpinBox:focus, QSpinBox:focus, QDateEdit:focus {
    border: 1px solid #89b4fa;
}
QLineEdit[status="ok"]    { border: 1px solid #a6e3a1; }
QLineEdit[status="error"] { border: 1px solid #f38ba8; }
QLineEdit[status="loading"]{ border: 1px solid #f9e2af; }
QComboBox::drop-down { border: none; }
QComboBox QAbstractItemView {
    background-color: #313244;
    border: 1px solid #45475a;
    selection-background-color: #89b4fa;
    color: #cdd6f4;
}
QPushButton {
    background-color: #89b4fa;
    color: #1e1e2e;
    border: none;
    border-radius: 8px;
    padding: 10px 22px;
    font-weight: bold;
    font-size: 13px;
}
QPushButton:hover   { background-color: #b4befe; }
QPushButton:pressed { background-color: #74c7ec; }
QPushButton:disabled{ background-color: #45475a; color: #6c7086; }
QPushButton#btn_secondary {
    background-color: #313244;
    color: #cdd6f4;
    border: 1px solid #45475a;
}
QPushButton#btn_secondary:hover { background-color: #45475a; }
QPushButton#btn_buscar {
    background-color: #313244;
    color: #89b4fa;
    border: 1px solid #89b4fa;
    border-radius: 6px;
    padding: 7px 14px;
    font-size: 12px;
}
QPushButton#btn_buscar:hover { background-color: #1e3a5f; }
QTabWidget::pane {
    border: 1px solid #313244;
    border-radius: 8px;
    background-color: #1e1e2e;
}
QTabBar::tab {
    background-color: #313244;
    color: #a6adc8;
    padding: 8px 16px;
    margin-right: 2px;
    border-radius: 6px 6px 0 0;
    font-size: 12px;
}
QTabBar::tab:selected { background-color: #89b4fa; color: #1e1e2e; font-weight: bold; }
QLabel#titulo   { font-size: 20px; font-weight: bold; color: #89b4fa; }
QLabel#subtitulo{ font-size: 12px; color: #a6adc8; }
QLabel#calc_label {
    background-color: #11111b; color: #a6e3a1;
    border-radius: 6px; padding: 6px 10px;
    font-family: monospace; font-size: 12px;
}
QLabel#api_status_ok    { color: #a6e3a1; font-size: 11px; }
QLabel#api_status_error { color: #f38ba8; font-size: 11px; }
QLabel#api_status_info  { color: #f9e2af; font-size: 11px; }
QStatusBar { background-color: #181825; color: #a6adc8; border-top: 1px solid #313244; }
QScrollArea { border: none; }
QFrame#sep  { background-color: #313244; max-height: 1px; }
QTextEdit#xml_preview {
    background-color: #11111b; color: #a6e3a1;
    font-family: 'Cascadia Code', 'Consolas', monospace;
    font-size: 12px; border: 1px solid #313244; border-radius: 6px;
}
QProgressBar {
    background-color: #313244; border: none; border-radius: 4px; height: 4px;
}
QProgressBar::chunk { background-color: #89b4fa; border-radius: 4px; }
"""

# ──────────────────────────────────────────────
# DADOS FIXOS DO PRESTADOR
# ──────────────────────────────────────────────
PRESTADOR = {
    "cnpj":                "24845408000122",
    "inscricao_municipal": "2686927",
    "razao_social":        "SANPORT - LOGISTICA PORTUARIA LTDA",
    "logradouro":          "Rua São José",
    "numero":              "38",
    "complemento":         "0513",
    "bairro":              "Embaré",
    "cidade":              "Santos",
    "uf":                  "SP",
    "cep":                 "11040200",
    "telefone":            "1333247708",
    "email":               "contato@teixeiracontabil.com.br",
    "codigo_municipio":    "3548500",
    "codigo_pais":         "1058",
}

# APIs em ordem de prioridade
CNPJ_APIS = [
    "https://receitaws.com.br/v1/cnpj/{cnpj}",
    "https://publica.cnpj.ws/cnpj/{cnpj}",
    "https://minhareceita.org/{cnpj}",
]

IBGE_MUNICIPIOS = {
    "santos":        "3548500",
    "sao paulo":     "3550308",
    "rio de janeiro":"3304557",
    "guaruja":       "3518701",
    "cubatao":       "3513504",
    "sao vicente":   "3551702",
    "praia grande":  "3541000",
    "bertioga":      "3506359",
}

INSCRICOES_MUNICIPAIS = {
    "08704068000163": "1767591",
    "22875162000106": "2648907",
    "60498706000904": "111737",
    "68014463000146": "1039957",
    "15193213000588": "2542003",
    "29169536000621": "3015746",
    "22253076000161": "2923787",
    "28529335000200": "3058101",
    "11256147000325": "1892039",
    "30694629000140": "2797638",
    "00728995000101": "1148369",
    "10790020000914": "812382",
    "33411794001107": "18775",
}


# ──────────────────────────────────────────────
# CACHE SQLite
# ──────────────────────────────────────────────
class CacheCNPJ:
    def __init__(self, path="tomadores_cache.db"):
        self.conn = sqlite3.connect(path)
        self.conn.execute("""
            CREATE TABLE IF NOT EXISTS tomadores (
                cnpj TEXT PRIMARY KEY,
                dados TEXT,
                atualizado TEXT
            )
        """)
        self.conn.commit()

    def get(self, cnpj: str) -> dict | None:
        cur = self.conn.execute(
            "SELECT dados FROM tomadores WHERE cnpj=?", (cnpj,))
        row = cur.fetchone()
        return json.loads(row[0]) if row else None

    def set(self, cnpj: str, dados: dict):
        self.conn.execute(
            "INSERT OR REPLACE INTO tomadores VALUES (?,?,?)",
            (cnpj, json.dumps(dados, ensure_ascii=False),
             datetime.now().isoformat()))
        self.conn.commit()

    def listar(self) -> list[tuple]:
        cur = self.conn.execute(
            "SELECT cnpj, json_extract(dados,'$.razao_social') FROM tomadores ORDER BY atualizado DESC")
        return cur.fetchall()


# ──────────────────────────────────────────────
# WORKER: consulta CNPJ em thread separada
# ──────────────────────────────────────────────
class CNPJWorker(QThread):
    resultado = pyqtSignal(dict)
    erro      = pyqtSignal(str)

    def __init__(self, cnpj: str):
        super().__init__()
        self.cnpj = ''.join(filter(str.isdigit, cnpj))

    def run(self):
        for template in CNPJ_APIS:
            url = template.format(cnpj=self.cnpj)
            try:
                req = urllib.request.Request(
                    url, headers={"User-Agent": "GeradorNFSe/3.0"})
                with urllib.request.urlopen(req, timeout=8) as r:
                    dados = json.loads(r.read().decode())
                    if dados.get("status") == "ERROR":
                        continue
                    self.resultado.emit(dados)
                    return
            except Exception:
                continue
        self.erro.emit("CNPJ não encontrado em nenhuma API disponível.")


# ──────────────────────────────────────────────
# HELPERS
# ──────────────────────────────────────────────
def limpar_num(v: str) -> str:
    return ''.join(filter(str.isdigit, v))

def formatar_cnpj(c: str) -> str:
    c = limpar_num(c)
    if len(c) == 14:
        return f"{c[:2]}.{c[2:5]}.{c[5:8]}/{c[8:12]}-{c[12:]}"
    return c

def fmt_val(v: float) -> str:
    return f"{v:.2f}"

def fmt_aliq(v: float) -> str:
    return f"{v/100:.4f}"

def municipio_para_ibge(nome: str) -> str:
    key = nome.lower().strip()
    return IBGE_MUNICIPIOS.get(key, PRESTADOR["codigo_municipio"])

def normalizar_dados_api(dados: dict) -> dict:
    """Normaliza resposta das diferentes APIs para formato único."""
    # ReceitaWS / publica.cnpj.ws / minhareceita têm campos parecidos
    telefone = (dados.get("telefone") or dados.get("ddd_telefone_1") or "").strip()
    municipio_nome = (dados.get("municipio") or dados.get("descricao_municipio") or "Santos")
    return {
        "razao_social":  dados.get("nome") or dados.get("razao_social") or "",
        "logradouro":    dados.get("logradouro") or "",
        "numero":        dados.get("numero") or "",
        "complemento":   dados.get("complemento") or "",
        "bairro":        dados.get("bairro") or "",
        "municipio":     municipio_nome,
        "uf":            dados.get("uf") or dados.get("estado") or "SP",
        "cep":           limpar_num(dados.get("cep") or ""),
        "email":         dados.get("email") or "",
        "telefone":      limpar_num(telefone),
        "cod_municipio": municipio_para_ibge(municipio_nome),
    }

def obter_inscricao_municipal(cnpj: str) -> str:
    return INSCRICOES_MUNICIPAIS.get(limpar_num(cnpj), "")


# ──────────────────────────────────────────────
# GERADOR XML (Manual v1.3)
# ──────────────────────────────────────────────
def gerar_xml_rps(d: dict) -> str:
    ns = "http://www.giss.com.br/send-nfse"
    ET.register_namespace("", ns)
    root = ET.Element(f"{{{ns}}}EnviarLoteRpsEnvio", versao="2.04")

    lote = ET.SubElement(root, "LoteRps", versao="2.04")
    ET.SubElement(lote, "NumeroLote").text = d.get("numero_lote", "1")
    pi = ET.SubElement(lote, "Prestador")
    cc = ET.SubElement(pi, "CpfCnpj")
    ET.SubElement(cc, "Cnpj").text = PRESTADOR["cnpj"]
    ET.SubElement(pi, "InscricaoMunicipal").text = PRESTADOR["inscricao_municipal"]
    ET.SubElement(lote, "QuantidadeRps").text = "1"

    lista = ET.SubElement(lote, "ListaRps")
    rps   = ET.SubElement(lista, "Rps")
    inf   = ET.SubElement(rps, "InfDeclaracaoPrestacaoServico",
                          Id=f"rps{d.get('numero_rps','1')}")

    # InfRps
    ir = ET.SubElement(inf, "Rps")
    id_r = ET.SubElement(ir, "IdentificacaoRps")
    ET.SubElement(id_r, "Numero").text  = d.get("numero_rps", "1")
    ET.SubElement(id_r, "Serie").text   = d.get("serie_rps",  "1")
    ET.SubElement(id_r, "Tipo").text    = "1"
    dt = d.get("data_emissao", datetime.now().strftime("%Y-%m-%d"))
    ET.SubElement(ir, "DataEmissao").text = dt
    ET.SubElement(ir, "Status").text      = "1"

    ET.SubElement(inf, "Competencia").text = d.get("competencia", dt[:7] + "-01")

    # Serviço
    srv  = ET.SubElement(inf, "Servico")
    vals = ET.SubElement(srv, "Valores")
    ET.SubElement(vals, "ValorServicos").text          = fmt_val(d.get("valor_servico", 0.0))
    ET.SubElement(vals, "ValorDeducoes").text          = fmt_val(d.get("valor_deducoes", 0.0))
    ET.SubElement(vals, "ValorPis").text               = "0.00"
    ET.SubElement(vals, "ValorCofins").text            = "0.00"
    ET.SubElement(vals, "ValorInss").text              = "0.00"
    ET.SubElement(vals, "ValorIr").text                = "0.00"
    ET.SubElement(vals, "ValorCsll").text              = "0.00"
    ET.SubElement(vals, "IssRetido").text              = "1" if d.get("iss_retido") else "2"
    ET.SubElement(vals, "ValorIss").text               = fmt_val(d.get("valor_iss", 0.0))
    ET.SubElement(vals, "Aliquota").text               = fmt_aliq(d.get("aliquota", 5.0))
    ET.SubElement(vals, "DescontoIncondicionado").text = "0.00"
    ET.SubElement(vals, "DescontoCondicionado").text   = "0.00"

    ET.SubElement(srv, "IssRetido").text               = "1" if d.get("iss_retido") else "2"
    ET.SubElement(srv, "ItemListaServico").text         = d.get("item_lista_servico", "20.01")
    if d.get("codigo_cnae"):
        ET.SubElement(srv, "CodigoCnae").text          = d["codigo_cnae"]
    ET.SubElement(srv, "CodigoTributacaoMunicipio").text = d.get("codigo_tributacao", "523110201")
    ET.SubElement(srv, "CodigoNbs").text               = d.get("codigo_nbs", "106500100")
    ET.SubElement(srv, "Discriminacao").text           = d.get("discriminacao", "")
    ET.SubElement(srv, "CodigoMunicipio").text         = PRESTADOR["codigo_municipio"]
    ET.SubElement(srv, "CodigoPais").text              = PRESTADOR["codigo_pais"]
    ET.SubElement(srv, "ExigibilidadeISS").text        = d.get("exigibilidade_iss", "1")
    ET.SubElement(srv, "MunicipioIncidencia").text     = PRESTADOR["codigo_municipio"]

    trib = ET.SubElement(srv, "trib")
    ET.SubElement(trib, "pTotTribFed").text = fmt_aliq(d.get("p_trib_fed", 0.0))
    ET.SubElement(trib, "pTotTribEst").text = fmt_aliq(d.get("p_trib_est", 0.0))
    ET.SubElement(trib, "pTotTribMun").text = fmt_aliq(d.get("p_trib_mun", 0.0))

    ibscbs = ET.SubElement(srv, "IBSCBS")
    ET.SubElement(ibscbs, "finNFSe").text           = "0"
    ET.SubElement(ibscbs, "cIndOp").text            = d.get("c_ind_op",           "000001")
    ET.SubElement(ibscbs, "indDest").text           = d.get("ind_dest",           "0")
    ET.SubElement(ibscbs, "CST").text               = d.get("cst_ibs_cbs",        "000")
    ET.SubElement(ibscbs, "cClassTrib").text        = d.get("c_class_trib",       "000001")
    ET.SubElement(ibscbs, "cLocalidadeIncid").text  = d.get("c_localidade_incid", PRESTADOR["codigo_municipio"])
    ET.SubElement(ibscbs, "vBC").text               = fmt_val(d.get("valor_bc_ibs", d.get("valor_servico", 0.0)))
    if d.get("v_cbs", 0.0) > 0:
        ET.SubElement(ibscbs, "vCBS").text          = fmt_val(d["v_cbs"])
    if d.get("v_ibs", 0.0) > 0:
        ET.SubElement(ibscbs, "vIBS").text          = fmt_val(d["v_ibs"])

    # Prestador
    pr = ET.SubElement(inf, "Prestador")
    cc2 = ET.SubElement(pr, "CpfCnpj")
    ET.SubElement(cc2, "Cnpj").text = PRESTADOR["cnpj"]
    ET.SubElement(pr, "InscricaoMunicipal").text = PRESTADOR["inscricao_municipal"]

    # Tomador
    tom  = ET.SubElement(inf, "TomadorServico")
    it   = ET.SubElement(tom, "IdentificacaoTomador")
    cc3  = ET.SubElement(it, "CpfCnpj")
    cnpj_t = limpar_num(d.get("tomador_cnpj", ""))
    if len(cnpj_t) == 14:
        ET.SubElement(cc3, "Cnpj").text = cnpj_t
    else:
        ET.SubElement(cc3, "Cpf").text  = cnpj_t
    if d.get("tomador_insc_municipal"):
        ET.SubElement(it, "InscricaoMunicipal").text = d["tomador_insc_municipal"]
    ET.SubElement(tom, "RazaoSocial").text = d.get("tomador_razao_social", "")

    end = ET.SubElement(tom, "Endereco")
    ET.SubElement(end, "Endereco").text        = d.get("tomador_logradouro", "")
    ET.SubElement(end, "Numero").text          = d.get("tomador_numero", "")
    if d.get("tomador_complemento"):
        ET.SubElement(end, "Complemento").text = d["tomador_complemento"]
    ET.SubElement(end, "Bairro").text          = d.get("tomador_bairro", "")
    ET.SubElement(end, "CodigoMunicipio").text = d.get("tomador_cod_municipio", PRESTADOR["codigo_municipio"])
    ET.SubElement(end, "Uf").text              = d.get("tomador_uf", "SP")
    ET.SubElement(end, "Cep").text             = limpar_num(d.get("tomador_cep", ""))

    if d.get("tomador_email") or d.get("tomador_telefone"):
        cont = ET.SubElement(tom, "Contato")
        if d.get("tomador_telefone"):
            ET.SubElement(cont, "Telefone").text = limpar_num(d["tomador_telefone"])
        if d.get("tomador_email"):
            ET.SubElement(cont, "Email").text    = d["tomador_email"]

    ET.SubElement(inf, "OptanteSimplesNacional").text = "1"
    ET.SubElement(inf, "IncentivoFiscal").text         = "2"

    ET.indent(root, space="  ")
    return '<?xml version="1.0" encoding="UTF-8"?>\n' + ET.tostring(root, encoding="unicode")


# ──────────────────────────────────────────────
# JANELA PRINCIPAL
# ──────────────────────────────────────────────
class GeradorNFSe(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Gerador NFS-e · GINFES Santos/SP · v3")
        self.setMinimumSize(1020, 800)
        self.setStyleSheet(STYLE)
        self.cache   = CacheCNPJ()
        self._worker = None
        self._build_ui()
        self._carregar_clientes_cache()
        self.status = QStatusBar()
        self.setStatusBar(self.status)
        self.status.showMessage("✓  Pronto — ReceitaWS + cache local ativo")

    # ── UI ───────────────────────────────────
    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        lay = QVBoxLayout(central)
        lay.setContentsMargins(20, 16, 20, 16)
        lay.setSpacing(10)

        # Cabeçalho
        hdr = QHBoxLayout()
        left = QVBoxLayout()
        t = QLabel("NFS-e · GINFES Santos/SP")
        t.setObjectName("titulo")
        s = QLabel("Prefeitura Municipal de Santos  ·  Secretaria de Finanças  ·  Manual v1.3  ·  LC 214/2025")
        s.setStyleSheet("font-size:12px;color:#a6adc8;")
        left.addWidget(t); left.addWidget(s)
        hdr.addLayout(left); hdr.addStretch()

        badge = QFrame()
        badge.setStyleSheet("background:#313244;border-radius:8px;")
        bl = QVBoxLayout(badge); bl.setContentsMargins(12,8,12,8); bl.setSpacing(2)
        b1 = QLabel(f"🏢  {PRESTADOR['razao_social']}")
        b1.setStyleSheet("font-weight:bold;color:#cdd6f4;font-size:12px;")
        b2 = QLabel(f"CNPJ: {formatar_cnpj(PRESTADOR['cnpj'])}  ·  IM: {PRESTADOR['inscricao_municipal']}")
        b2.setStyleSheet("font-size:11px;color:#a6adc8;")
        bl.addWidget(b1); bl.addWidget(b2)
        hdr.addWidget(badge)
        lay.addLayout(hdr)

        sep = QFrame(); sep.setObjectName("sep"); sep.setFrameShape(QFrame.Shape.HLine)
        lay.addWidget(sep)

        # Progress bar (oculta por padrão)
        self.progress = QProgressBar()
        self.progress.setRange(0, 0)
        self.progress.setFixedHeight(4)
        self.progress.hide()
        lay.addWidget(self.progress)

        # Tabs
        self.tabs = QTabWidget()
        lay.addWidget(self.tabs)
        self.tabs.addTab(self._tab_nf(),      "📄  Dados da NF")
        self.tabs.addTab(self._tab_tomador(),  "👤  Tomador")
        self.tabs.addTab(self._tab_ibscbs(),   "🏛️  IBS / CBS")
        self.tabs.addTab(self._tab_preview(),  "🔍  Preview XML")

        # Botões
        brow = QHBoxLayout(); brow.setSpacing(10)
        b_limpar = QPushButton("🔄  Reset")
        b_limpar.setObjectName("btn_secondary")
        b_limpar.clicked.connect(self._limpar)
        b_prev = QPushButton("👁  Gerar Preview")
        b_prev.setObjectName("btn_secondary")
        b_prev.clicked.connect(self._preview)
        b_salvar = QPushButton("💾  Salvar XML")
        b_salvar.clicked.connect(self._salvar)
        brow.addWidget(b_limpar); brow.addStretch()
        brow.addWidget(b_prev); brow.addWidget(b_salvar)
        lay.addLayout(brow)

    # ── TAB 1: DADOS DA NF ───────────────────
    def _tab_nf(self) -> QWidget:
        scroll = QScrollArea(); scroll.setWidgetResizable(True)
        w = QWidget(); lay = QVBoxLayout(w); lay.setSpacing(12)

        g = QGroupBox("Identificação do RPS"); f = QFormLayout(g); f.setSpacing(8)
        self.numero_rps = QSpinBox(); self.numero_rps.setRange(1,999999); self.numero_rps.setValue(1)
        f.addRow("Número RPS:", self.numero_rps)
        self.serie_rps = QLineEdit("1"); f.addRow("Série:", self.serie_rps)
        self.data_emissao = QDateEdit(QDate.currentDate())
        self.data_emissao.setDisplayFormat("dd/MM/yyyy"); self.data_emissao.setCalendarPopup(True)
        f.addRow("Data de Emissão:", self.data_emissao)
        self.competencia = QDateEdit(QDate.currentDate())
        self.competencia.setDisplayFormat("MM/yyyy"); self.competencia.setCalendarPopup(True)
        f.addRow("Competência:", self.competencia)
        lay.addWidget(g)

        g2 = QGroupBox("Serviço"); f2 = QFormLayout(g2); f2.setSpacing(8)
        self.item_lista        = QLineEdit("20.01");    f2.addRow("Item Lista Serviço (LC 116):", self.item_lista)
        self.codigo_tributacao = QLineEdit("523110201");f2.addRow("Cód. Tributação Municipal:", self.codigo_tributacao)
        self.codigo_cnae       = QLineEdit("5231102");  f2.addRow("Código CNAE:", self.codigo_cnae)
        self.codigo_nbs        = QLineEdit("106500100");f2.addRow("Código NBS (obrigatório v1.3):", self.codigo_nbs)
        self.exigibilidade = QComboBox()
        self.exigibilidade.addItems(["1 – Exigível","2 – Não incidência","3 – Isenção",
                                     "4 – Exportação","5 – Imunidade",
                                     "6 – Susp. Decisão Judicial","7 – Susp. Processo Adm."])
        f2.addRow("Exigibilidade ISSQN:", self.exigibilidade)
        self.discriminacao = QTextEdit()
        self.discriminacao.setPlaceholderText("Descrição detalhada do serviço prestado...")
        self.discriminacao.setFixedHeight(80)
        f2.addRow("Discriminação:", self.discriminacao)
        lay.addWidget(g2)

        g3 = QGroupBox("Valores"); f3 = QFormLayout(g3); f3.setSpacing(8)
        self.valor_servico = QDoubleSpinBox()
        self.valor_servico.setRange(0,9_999_999); self.valor_servico.setDecimals(2)
        self.valor_servico.setPrefix("R$ "); self.valor_servico.valueChanged.connect(self._recalcular)
        f3.addRow("Valor do Serviço:", self.valor_servico)
        self.valor_deducoes = QDoubleSpinBox()
        self.valor_deducoes.setRange(0,9_999_999); self.valor_deducoes.setDecimals(2)
        self.valor_deducoes.setPrefix("R$ "); self.valor_deducoes.valueChanged.connect(self._recalcular)
        f3.addRow("Deduções:", self.valor_deducoes)
        self.aliquota = QDoubleSpinBox()
        self.aliquota.setRange(0,10); self.aliquota.setDecimals(2); self.aliquota.setValue(5.0)
        self.aliquota.setSuffix(" %"); self.aliquota.valueChanged.connect(self._recalcular)
        f3.addRow("Alíquota ISSQN:", self.aliquota)
        self.iss_retido = QComboBox()
        self.iss_retido.addItems(["1 – Sim (ISS Retido)","2 – Não (ISS a Recolher)"])
        f3.addRow("ISS Retido?", self.iss_retido)
        self.calc_label = QLabel("—"); self.calc_label.setObjectName("calc_label")
        f3.addRow("Resumo:", self.calc_label)
        lay.addWidget(g3)

        lay.addStretch(); scroll.setWidget(w)
        return scroll

    # ── TAB 2: TOMADOR ───────────────────────
    def _tab_tomador(self) -> QWidget:
        scroll = QScrollArea(); scroll.setWidgetResizable(True)
        w = QWidget(); lay = QVBoxLayout(w); lay.setSpacing(12)

        g = QGroupBox("Dados do Tomador de Serviço")
        f = QFormLayout(g); f.setSpacing(8)

        self.tom_cliente_cache = QComboBox()
        self.tom_cliente_cache.currentIndexChanged.connect(self._on_cliente_cache_changed)
        f.addRow("Clientes salvos:", self.tom_cliente_cache)

        # CNPJ + botão buscar na mesma linha
        cnpj_row = QHBoxLayout(); cnpj_row.setSpacing(8)
        self.tom_cnpj = QLineEdit()
        self.tom_cnpj.setPlaceholderText("00.000.000/0000-00 ou 000.000.000-00")
        self.btn_buscar = QPushButton("🔍  Buscar")
        self.btn_buscar.setObjectName("btn_buscar")
        self.btn_buscar.setFixedWidth(110)
        self.btn_buscar.clicked.connect(self._buscar_cnpj)
        self.tom_cnpj.returnPressed.connect(self._buscar_cnpj)
        cnpj_row.addWidget(self.tom_cnpj); cnpj_row.addWidget(self.btn_buscar)
        f.addRow("CPF/CNPJ:", cnpj_row)

        # Status da API
        self.api_status = QLabel("Digite o CNPJ e clique em Buscar")
        self.api_status.setObjectName("api_status_info")
        f.addRow("", self.api_status)

        self.tom_razao = QLineEdit(); f.addRow("Razão Social:", self.tom_razao)
        self.tom_im    = QLineEdit(); f.addRow("Inscrição Municipal:", self.tom_im)
        self.tom_email = QLineEdit(); self.tom_email.setPlaceholderText("email@empresa.com.br")
        f.addRow("E-mail:", self.tom_email)
        self.tom_fone  = QLineEdit(); self.tom_fone.setPlaceholderText("(00) 00000-0000")
        f.addRow("Telefone:", self.tom_fone)

        sep2 = QFrame(); sep2.setFrameShape(QFrame.Shape.HLine)
        sep2.setStyleSheet("background:#45475a;max-height:1px;")
        f.addRow(sep2)

        self.tom_logradouro  = QLineEdit(); f.addRow("Logradouro:", self.tom_logradouro)
        self.tom_numero      = QLineEdit(); f.addRow("Número:", self.tom_numero)
        self.tom_complemento = QLineEdit(); f.addRow("Complemento:", self.tom_complemento)
        self.tom_bairro      = QLineEdit(); f.addRow("Bairro:", self.tom_bairro)
        self.tom_cep         = QLineEdit(); self.tom_cep.setPlaceholderText("00000-000")
        f.addRow("CEP:", self.tom_cep)
        self.tom_uf = QComboBox()
        self.tom_uf.addItems(["SP","RJ","MG","RS","PR","SC","BA","GO","DF","ES","CE",
                              "PE","AM","PA","MA","MT","MS","RN","PB","PI","AL","SE",
                              "RO","TO","AC","AP","RR"])
        f.addRow("UF:", self.tom_uf)
        self.tom_cod_municipio = QLineEdit(PRESTADOR["codigo_municipio"])
        f.addRow("Cód. Município IBGE:", self.tom_cod_municipio)
        lay.addWidget(g)

        lay.addStretch(); scroll.setWidget(w)
        return scroll

    # ── TAB 3: IBS/CBS ───────────────────────
    def _tab_ibscbs(self) -> QWidget:
        scroll = QScrollArea(); scroll.setWidgetResizable(True)
        w = QWidget(); lay = QVBoxLayout(w); lay.setSpacing(12)

        info = QLabel("⚠️  Campos obrigatórios — LC 214/2025 · Manual GINFES v1.3 (jan/2026)")
        info.setStyleSheet("background:#1e3a5f;border-radius:6px;padding:8px 12px;"
                           "color:#89b4fa;font-size:12px;")
        lay.addWidget(info)

        g = QGroupBox("IBS e CBS — Imposto / Contribuição sobre Bens e Serviços")
        f = QFormLayout(g); f.setSpacing(8)
        self.c_ind_op = QComboBox()
        self.c_ind_op.addItems(["000001 – Prestação de serviço",
                                "000002 – Locação de bem móvel",
                                "000003 – Cessão de direito"])
        f.addRow("Indicador de Operação (cIndOp):", self.c_ind_op)
        self.ind_dest = QComboBox()
        self.ind_dest.addItems(["0 – Destinatário é o tomador","1 – Destinatário é terceiro"])
        f.addRow("Destinatário (indDest):", self.ind_dest)
        self.cst_ibs_cbs = QLineEdit("000");    f.addRow("CST IBS/CBS:", self.cst_ibs_cbs)
        self.c_class_trib = QLineEdit("000001");f.addRow("Classificação Tributária (cClassTrib):", self.c_class_trib)
        self.c_localidade_incid = QLineEdit(PRESTADOR["codigo_municipio"])
        f.addRow("Localidade de Incidência (IBGE):", self.c_localidade_incid)

        s2 = QFrame(); s2.setFrameShape(QFrame.Shape.HLine)
        s2.setStyleSheet("background:#45475a;max-height:1px;")
        f.addRow(s2)

        self.v_cbs = QDoubleSpinBox()
        self.v_cbs.setRange(0,999999); self.v_cbs.setDecimals(2); self.v_cbs.setPrefix("R$ ")
        self.v_cbs.setValue(28.70); f.addRow("Valor CBS (R$):", self.v_cbs)
        self.v_ibs = QDoubleSpinBox()
        self.v_ibs.setRange(0,999999); self.v_ibs.setDecimals(2); self.v_ibs.setPrefix("R$ ")
        f.addRow("Valor IBS (R$):", self.v_ibs)
        self.valor_bc_ibs = QDoubleSpinBox()
        self.valor_bc_ibs.setRange(0,9_999_999); self.valor_bc_ibs.setDecimals(2)
        self.valor_bc_ibs.setPrefix("R$ "); f.addRow("Base de Cálculo IBS/CBS (vBC):", self.valor_bc_ibs)
        lay.addWidget(g)

        g2 = QGroupBox("Tributos Aproximados — Lei 12.741/2012")
        f2 = QFormLayout(g2); f2.setSpacing(8)
        self.p_trib_fed = QDoubleSpinBox()
        self.p_trib_fed.setRange(0,100); self.p_trib_fed.setDecimals(2); self.p_trib_fed.setSuffix(" %")
        f2.addRow("% Tributos Federais:", self.p_trib_fed)
        self.p_trib_est = QDoubleSpinBox()
        self.p_trib_est.setRange(0,100); self.p_trib_est.setDecimals(2); self.p_trib_est.setSuffix(" %")
        f2.addRow("% Tributos Estaduais:", self.p_trib_est)
        self.p_trib_mun = QDoubleSpinBox()
        self.p_trib_mun.setRange(0,100); self.p_trib_mun.setDecimals(2)
        self.p_trib_mun.setSuffix(" %"); self.p_trib_mun.setValue(5.0)
        f2.addRow("% Tributos Municipais:", self.p_trib_mun)
        lay.addWidget(g2)

        lay.addStretch(); scroll.setWidget(w)
        return scroll

    # ── TAB 4: PREVIEW ───────────────────────
    def _tab_preview(self) -> QWidget:
        w = QWidget(); lay = QVBoxLayout(w); lay.setContentsMargins(0,8,0,0)
        lbl = QLabel("XML no padrão ABRASF 2.04 + LC 214/2025 — pronto para upload no GINFES ou envio via API:")
        lbl.setStyleSheet("color:#a6adc8;font-size:12px;"); lbl.setWordWrap(True)
        lay.addWidget(lbl)
        self.xml_preview = QTextEdit(); self.xml_preview.setObjectName("xml_preview")
        self.xml_preview.setReadOnly(True)
        self.xml_preview.setPlaceholderText("Clique em 'Gerar Preview' para visualizar o XML...")
        lay.addWidget(self.xml_preview)
        return w

    # ── BUSCA CNPJ ───────────────────────────
    def _buscar_cnpj(self):
        cnpj_raw = self.tom_cnpj.text().strip()
        cnpj     = limpar_num(cnpj_raw)

        if len(cnpj) not in (11, 14):
            self._set_api_status("error", "⚠️  CNPJ inválido — digite 14 dígitos")
            return

        # Verificar cache primeiro
        cached = self.cache.get(cnpj)
        if cached:
            self._preencher_tomador(cached)
            self._set_api_status("ok", f"✓  Dados carregados do cache local")
            self.status.showMessage("✓  Tomador carregado do cache local")
            return

        # Consultar API
        self._set_api_status("info", "⏳  Consultando ReceitaWS...")
        self.progress.show()
        self.btn_buscar.setEnabled(False)
        self.btn_buscar.setText("...")

        self._worker = CNPJWorker(cnpj)
        self._worker.resultado.connect(self._on_cnpj_ok)
        self._worker.erro.connect(self._on_cnpj_erro)
        self._worker.start()

    def _on_cnpj_ok(self, dados: dict):
        self.progress.hide()
        self.btn_buscar.setEnabled(True)
        self.btn_buscar.setText("🔍  Buscar")

        norm = normalizar_dados_api(dados)
        cnpj = limpar_num(self.tom_cnpj.text())
        norm["inscricao_municipal"] = obter_inscricao_municipal(cnpj)
        self.cache.set(cnpj, norm)
        self._carregar_clientes_cache(cnpj_selecionado=cnpj)
        self._preencher_tomador(norm)
        self._set_api_status("ok", f"✓  Dados obtidos via API e salvos no cache")
        self.status.showMessage(f"✓  {norm['razao_social']} — CNPJ {formatar_cnpj(cnpj)}")

    def _on_cnpj_erro(self, msg: str):
        self.progress.hide()
        self.btn_buscar.setEnabled(True)
        self.btn_buscar.setText("🔍  Buscar")
        self._set_api_status("error", f"✗  {msg}")
        self.status.showMessage(f"Erro: {msg}")

    def _set_api_status(self, tipo: str, msg: str):
        obj = {"ok": "api_status_ok", "error": "api_status_error", "info": "api_status_info"}
        self.api_status.setObjectName(obj.get(tipo, "api_status_info"))
        self.api_status.setText(msg)
        self.api_status.setStyleSheet(self.api_status.styleSheet())
        self.api_status.style().unpolish(self.api_status)
        self.api_status.style().polish(self.api_status)

    def _preencher_tomador(self, n: dict):
        cnpj_atual = limpar_num(self.tom_cnpj.text())
        im = n.get("inscricao_municipal", "") or obter_inscricao_municipal(cnpj_atual)
        self.tom_im.setText(im)
        self.tom_razao.setText(n.get("razao_social", ""))
        self.tom_email.setText(n.get("email", ""))
        fone = n.get("telefone", "")
        if fone:
            self.tom_fone.setText(f"({fone[:2]}) {fone[2:]}" if len(fone) >= 10 else fone)
        self.tom_logradouro.setText(n.get("logradouro", ""))
        self.tom_numero.setText(n.get("numero", ""))
        self.tom_complemento.setText(n.get("complemento", ""))
        self.tom_bairro.setText(n.get("bairro", ""))
        cep = n.get("cep", "")
        self.tom_cep.setText(f"{cep[:5]}-{cep[5:]}" if len(cep) == 8 else cep)
        uf = n.get("uf", "SP")
        idx = self.tom_uf.findText(uf)
        if idx >= 0: self.tom_uf.setCurrentIndex(idx)
        self.tom_cod_municipio.setText(n.get("cod_municipio", PRESTADOR["codigo_municipio"]))

    def _carregar_clientes_cache(self, cnpj_selecionado: str | None = None):
        if not hasattr(self, "tom_cliente_cache"):
            return

        clientes = self.cache.listar()
        self.tom_cliente_cache.blockSignals(True)
        self.tom_cliente_cache.clear()
        self.tom_cliente_cache.addItem("Selecione um cliente...")

        idx_para_selecionar = 0
        for idx, (cnpj, razao) in enumerate(clientes, start=1):
            cnpj_limpo = limpar_num(cnpj or "")
            razao_txt = (razao or "(Sem razão social)").strip()
            label = f"{razao_txt} — {formatar_cnpj(cnpj_limpo)}"
            self.tom_cliente_cache.addItem(label, cnpj_limpo)
            if cnpj_selecionado and cnpj_limpo == limpar_num(cnpj_selecionado):
                idx_para_selecionar = idx

        self.tom_cliente_cache.setCurrentIndex(idx_para_selecionar)
        self.tom_cliente_cache.blockSignals(False)

    def _on_cliente_cache_changed(self, index: int):
        if index <= 0:
            return

        cnpj = limpar_num(self.tom_cliente_cache.itemData(index) or "")
        if not cnpj:
            return

        dados = self.cache.get(cnpj)
        if not dados:
            return

        if not dados.get("inscricao_municipal"):
            dados["inscricao_municipal"] = obter_inscricao_municipal(cnpj)

        self.tom_cnpj.setText(formatar_cnpj(cnpj))
        self._preencher_tomador(dados)
        self._set_api_status("ok", "✓  Cliente carregado da lista")
        self.status.showMessage(f"✓  Cliente selecionado: {dados.get('razao_social', '')}")

    # ── HELPERS ──────────────────────────────
    def _recalcular(self):
        vs  = self.valor_servico.value()
        ded = self.valor_deducoes.value()
        aliq = self.aliquota.value()
        bc  = vs - ded
        iss = round(bc * aliq / 100, 2)
        liq = round(vs - iss if self.iss_retido.currentIndex() == 0 else vs, 2)
        self.calc_label.setText(
            f"BC = R$ {bc:,.2f}  |  ISSQN = R$ {iss:,.2f}  |  Líquido = R$ {liq:,.2f}")
        self.valor_bc_ibs.setValue(bc)

    def _coletar(self) -> dict:
        return {
            "numero_rps":             str(self.numero_rps.value()),
            "serie_rps":              self.serie_rps.text().strip(),
            "data_emissao":           self.data_emissao.date().toString("yyyy-MM-dd"),
            "competencia":            self.competencia.date().toString("yyyy-MM-01"),
            "item_lista_servico":     self.item_lista.text().strip(),
            "codigo_tributacao":      self.codigo_tributacao.text().strip(),
            "codigo_cnae":            self.codigo_cnae.text().strip(),
            "codigo_nbs":             self.codigo_nbs.text().strip(),
            "exigibilidade_iss":      str(self.exigibilidade.currentIndex() + 1),
            "discriminacao":          self.discriminacao.toPlainText().strip(),
            "valor_servico":          self.valor_servico.value(),
            "valor_deducoes":         self.valor_deducoes.value(),
            "aliquota":               self.aliquota.value(),
            "valor_iss":              round((self.valor_servico.value() - self.valor_deducoes.value())
                                            * self.aliquota.value() / 100, 2),
            "iss_retido":             self.iss_retido.currentIndex() == 0,
            "tomador_cnpj":           self.tom_cnpj.text().strip(),
            "tomador_razao_social":   self.tom_razao.text().strip(),
            "tomador_insc_municipal": self.tom_im.text().strip(),
            "tomador_email":          self.tom_email.text().strip(),
            "tomador_telefone":       self.tom_fone.text().strip(),
            "tomador_logradouro":     self.tom_logradouro.text().strip(),
            "tomador_numero":         self.tom_numero.text().strip(),
            "tomador_complemento":    self.tom_complemento.text().strip(),
            "tomador_bairro":         self.tom_bairro.text().strip(),
            "tomador_cep":            self.tom_cep.text().strip(),
            "tomador_uf":             self.tom_uf.currentText(),
            "tomador_cod_municipio":  self.tom_cod_municipio.text().strip(),
            "c_ind_op":               self.c_ind_op.currentText().split(" – ")[0],
            "ind_dest":               str(self.ind_dest.currentIndex()),
            "cst_ibs_cbs":            self.cst_ibs_cbs.text().strip(),
            "c_class_trib":           self.c_class_trib.text().strip(),
            "c_localidade_incid":     self.c_localidade_incid.text().strip(),
            "v_cbs":                  self.v_cbs.value(),
            "v_ibs":                  self.v_ibs.value(),
            "valor_bc_ibs":           self.valor_bc_ibs.value(),
            "p_trib_fed":             self.p_trib_fed.value(),
            "p_trib_est":             self.p_trib_est.value(),
            "p_trib_mun":             self.p_trib_mun.value(),
            "numero_lote":            "1",
        }

    def _validar(self, d: dict) -> str | None:
        if not d["discriminacao"]:        return "Preencha a discriminação do serviço."
        if not d["tomador_cnpj"]:         return "Informe o CPF/CNPJ do tomador."
        if not d["tomador_razao_social"]: return "Informe a Razão Social do tomador."
        if d["valor_servico"] <= 0:       return "O valor do serviço deve ser maior que zero."
        if not d["codigo_nbs"]:           return "Código NBS é obrigatório (manual v1.3)."
        return None

    def _preview(self):
        d = self._coletar(); e = self._validar(d)
        if e: QMessageBox.warning(self, "Atenção", e); return
        xml = gerar_xml_rps(d)
        self.xml_preview.setPlainText(xml)
        self.tabs.setCurrentIndex(3)
        self.status.showMessage(f"✓  XML gerado — {len(xml):,} caracteres")

    def _salvar(self):
        d = self._coletar(); e = self._validar(d)
        if e: QMessageBox.warning(self, "Atenção", e); return
        xml = gerar_xml_rps(d)
        sug = f"RPS_{d['numero_rps']}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xml"
        path, _ = QFileDialog.getSaveFileName(self, "Salvar XML", sug, "XML (*.xml)")
        if path:
            with open(path, "w", encoding="utf-8") as fp:
                fp.write(xml)
            self.status.showMessage(f"✓  Arquivo salvo: {path}")
            QMessageBox.information(self, "Sucesso!",
                f"XML salvo!\n\n{path}\n\n"
                "Para emitir: acesse santos.ginfes.com.br → Importar RPS.")

    def _limpar(self):
        if QMessageBox.question(self, "Reset", "Deseja resetar todos os campos?") \
                != QMessageBox.StandardButton.Yes:
            return

        # Aba NF
        self.numero_rps.setValue(1); self.serie_rps.setText("1")
        self.data_emissao.setDate(QDate.currentDate())
        self.competencia.setDate(QDate.currentDate())
        self.item_lista.setText("20.01")
        self.codigo_tributacao.setText("523110201")
        self.codigo_cnae.setText("5231102")
        self.codigo_nbs.setText("106500100")
        self.exigibilidade.setCurrentIndex(0)
        self.discriminacao.clear()
        self.valor_servico.setValue(0); self.valor_deducoes.setValue(0)
        self.aliquota.setValue(5.0)
        self.iss_retido.setCurrentIndex(0)
        self.calc_label.setText("—")

        # Aba Tomador
        for w in [self.tom_cnpj, self.tom_razao, self.tom_im, self.tom_email,
                  self.tom_fone, self.tom_logradouro, self.tom_numero,
                  self.tom_complemento, self.tom_bairro, self.tom_cep]:
            w.clear()
        self.tom_uf.setCurrentText("SP")
        self.tom_cod_municipio.setText(PRESTADOR["codigo_municipio"])
        if hasattr(self, "tom_cliente_cache"):
            self.tom_cliente_cache.setCurrentIndex(0)

        # Aba IBS/CBS
        self.c_ind_op.setCurrentIndex(0)
        self.ind_dest.setCurrentIndex(0)
        self.cst_ibs_cbs.setText("000")
        self.c_class_trib.setText("000001")
        self.c_localidade_incid.setText(PRESTADOR["codigo_municipio"])
        self.v_cbs.setValue(28.70)
        self.v_ibs.setValue(0.0)
        self.valor_bc_ibs.setValue(0.0)
        self.p_trib_fed.setValue(0.0)
        self.p_trib_est.setValue(0.0)
        self.p_trib_mun.setValue(5.0)

        # UI geral
        self._set_api_status("info", "Digite o CNPJ e clique em Buscar")
        self.xml_preview.clear()
        self.tabs.setCurrentIndex(0)
        self.status.showMessage("Campos resetados.")

# ──────────────────────────────────────────────
if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    w = GeradorNFSe()
    w.show()
    sys.exit(app.exec())