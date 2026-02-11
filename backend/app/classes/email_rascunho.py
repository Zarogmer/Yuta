import pythoncom
import win32com.client
from pathlib import Path
from datetime import datetime, date
import unicodedata


DEFAULT_ASSUNTO = "FATURAMENTO SANTOS {dn}/{ano} - M/V {navio}"
ASSUNTO_SAO_SEBASTIAO = "FATURAMENTO {dn}/{ano2} - M/V {navio} - PORTO DE SÃO SEBASTIÃO"
DEFAULT_CORPO = """Prezados, {saudacao}!

Seguem anexos faturamento e folhas OGMO do navio {navio} em referência.

Obs.: Gentileza notar alteração nos dados bancários.

Dados para depósito:

Banco Itaú

Agência: 0447

Conta Corrente: 99807-1

Pix: 24.845.408/0001-22

Atenciosamente,
"""

DEFAULT_CORPO_HTML = """<div style="font-family: Arial, sans-serif; font-size: 12pt; color: #000;">
    <p>Prezados, {saudacao}!</p>
    <p>Seguem anexos faturamento e folhas OGMO do navio {navio} em referência.</p>
    <p><strong>Obs.:</strong> Gentileza notar alteração nos dados bancários.</p>
    <table style="border: 1px solid #000; border-collapse: collapse; margin: 6px 0 10px;" cellpadding="6" cellspacing="0">
        <tr>
            <td style="padding: 6px 10px;">
                <div><strong>Dados para depósito:</strong></div>
                <div>Banco Itaú</div>
                <div>Agência: 0447</div>
                <div>Conta Corrente: 99807-1</div>
                <div>Pix: 24.845.408/0001-22</div>
            </td>
        </tr>
    </table>
    <p>Atenciosamente,</p>
</div>
"""

CORPO_SAO_SEBASTIAO = """Prezados, {saudacao}!

Segue anexo faturamento do navio {navio} em referência.

Atenciosamente,
"""

CORPO_SAO_SEBASTIAO_HTML = """<div style="font-family: Arial, sans-serif; font-size: 12pt; color: #000;">
    <p>Prezados, {saudacao}!</p>
    <p>Segue anexo faturamento do navio {navio} em referência.</p>
    <p>Atenciosamente,</p>
</div>
"""

CORPO_CARGONAVE = """Prezados, {saudacao}!

Gentileza nos enviar dados para emissão de nota fiscal do navio {navio} em referência.

Solicitamos remessa na conta corrente abaixo no valor de {adiantamento_fmt} conforme acordo.

Dados para depósito:

Banco Itaú

Agência: 0447

Conta Corrente: 99807-1

Pix: 24.845.408/0001-22

Desde já muito obrigado.

Atenciosamente,
"""

CORPO_CARGONAVE_HTML = """<div style="font-family: Arial, sans-serif; font-size: 12pt; color: #000;">
    <p>Prezados, {saudacao}!</p>
    <p>Gentileza nos enviar dados para emissão de nota fiscal do navio {navio} em referência.</p>
    <p>Solicitamos remessa na conta corrente abaixo no valor de <strong>{adiantamento_fmt}</strong> conforme acordo.</p>
    <table style="border: 1px solid #000; border-collapse: collapse; margin: 6px 0 10px;" cellpadding="6" cellspacing="0">
        <tr>
            <td style="padding: 6px 10px;">
                <div><strong>Dados para depósito:</strong></div>
                <div>Banco Itaú</div>
                <div>Agência: 0447</div>
                <div>Conta Corrente: 99807-1</div>
                <div>Pix: 24.845.408/0001-22</div>
            </td>
        </tr>
    </table>
    <p>Desde já muito obrigado.</p>
    <p>Atenciosamente,</p>
</div>
"""

CORPO_ROCHAMAR = """Prezados, {saudacao}!

Solicitamos o número da OC do navio em referência e seguem valores da fatura abaixo:

M/V {navio}

Atracação: {atracacao_ini} a {atracacao_fim}

Despesas OGMO: {costs_fmt}

Taxa Administrativa: {adm_fmt}

Atenciosamente,
"""

CORPO_ROCHAMAR_HTML = """<div style="font-family: Arial, sans-serif; font-size: 12pt; color: #000;">
    <p>Prezados, {saudacao}!</p>
    <p>Solicitamos o número da OC do navio em referência e seguem valores da fatura abaixo:</p>
    <p><strong>M/V {navio}</strong></p>
    <p>Atracação: {atracacao_ini} a {atracacao_fim}</p>
    <p>Despesas OGMO: {costs_fmt}</p>
    <p>Taxa Administrativa: {adm_fmt}</p>
    <p>Atenciosamente,</p>
</div>
"""

# Edite aqui para personalizar por cliente.
CLIENTES_EMAIL = {
    "CARGONAVE": {
        "para": ["guigui12306@gmail.com"],
        "cc": [],
        "assunto": DEFAULT_ASSUNTO,
        "corpo": CORPO_CARGONAVE,
        "corpo_html": CORPO_CARGONAVE_HTML,
    },
    "WILSON": {
        "para": ["guigui12306@gmail.com"],
        "cc": [],
        "assunto": DEFAULT_ASSUNTO,
        "corpo": DEFAULT_CORPO,
    },
    "ROCHAMAR": {
        "para": ["guigui12306@gmail.com"],
        "cc": [],
        "assunto": DEFAULT_ASSUNTO,
        "corpo": CORPO_ROCHAMAR,
        "corpo_html": CORPO_ROCHAMAR_HTML,
    },
    "UNIMAR": {
        "para": ["guigui12306@gmail.com"],
        "cc": [],
        "assunto": DEFAULT_ASSUNTO,
        "corpo": DEFAULT_CORPO,
        "corpo_html": DEFAULT_CORPO_HTML,
    },
    "WILSON SONS - SAO SEBASTIAO": {
        "para": ["guigui12306@gmail.com"],
        "cc": [],
        "assunto": ASSUNTO_SAO_SEBASTIAO,
        "corpo": CORPO_SAO_SEBASTIAO,
        "corpo_html": CORPO_SAO_SEBASTIAO_HTML,
    },
    "AQUARIUS - PSS": {
        "para": ["guigui12306@gmail.com"],
        "cc": [],
        "assunto": ASSUNTO_SAO_SEBASTIAO,
        "corpo": CORPO_SAO_SEBASTIAO,
        "corpo_html": CORPO_SAO_SEBASTIAO_HTML,
    },
    "SEA SIDE - PSS": {
        "para": ["guigui12306@gmail.com"],
        "cc": [],
        "assunto": ASSUNTO_SAO_SEBASTIAO,
        "corpo": CORPO_SAO_SEBASTIAO,
        "corpo_html": CORPO_SAO_SEBASTIAO_HTML,
    },
    "ARENNA": {
        "para": ["guigui12306@gmail.com"],
        "cc": [],
        "assunto": DEFAULT_ASSUNTO,
        "corpo": DEFAULT_CORPO,
    },
    "CARGILL": {
        "para": ["guigui12306@gmail.com"],
        "cc": [],
        "assunto": DEFAULT_ASSUNTO,
        "corpo": DEFAULT_CORPO,
    },
    "CONESUL": {
        "para": ["guigui12306@gmail.com"],
        "cc": [],
        "assunto": DEFAULT_ASSUNTO,
        "corpo": DEFAULT_CORPO,
    },
    "DELTA": {
        "para": ["guigui12306@gmail.com"],
        "cc": [],
        "assunto": DEFAULT_ASSUNTO,
        "corpo": DEFAULT_CORPO,
    },
    "GEM": {
        "para": ["guigui12306@gmail.com"],
        "cc": [],
        "assunto": DEFAULT_ASSUNTO,
        "corpo": DEFAULT_CORPO,
    },
    "HMS": {
        "para": ["guigui12306@gmail.com"],
        "cc": [],
        "assunto": DEFAULT_ASSUNTO,
        "corpo": DEFAULT_CORPO,
    },
    "JBG": {
        "para": ["guigui12306@gmail.com"],
        "cc": [],
        "assunto": DEFAULT_ASSUNTO,
        "corpo": DEFAULT_CORPO,
    },
    "LMA": {
        "para": ["guigui12306@gmail.com"],
        "cc": [],
        "assunto": DEFAULT_ASSUNTO,
        "corpo": DEFAULT_CORPO,
    },
    "NAABSA": {
        "para": ["guigui12306@gmail.com"],
        "cc": [],
        "assunto": DEFAULT_ASSUNTO,
        "corpo": DEFAULT_CORPO,
    },
    "NORTH STAR": {
        "para": ["guigui12306@gmail.com"],
        "cc": [],
        "assunto": DEFAULT_ASSUNTO,
        "corpo": DEFAULT_CORPO,
    },
    "PROJECT CARGO": {
        "para": ["guigui12306@gmail.com"],
        "cc": [],
        "assunto": DEFAULT_ASSUNTO,
        "corpo": DEFAULT_CORPO,
    },
    "TRANSATLANTICA": {
        "para": ["guigui12306@gmail.com"],
        "cc": [],
        "assunto": DEFAULT_ASSUNTO,
        "corpo": DEFAULT_CORPO,
    },
    "WILLIAMS": {
        "para": ["guigui12306@gmail.com"],
        "cc": [],
        "assunto": DEFAULT_ASSUNTO,
        "corpo": DEFAULT_CORPO,
    },
    "WILSON SONS": {
        "para": ["guigui12306@gmail.com"],
        "cc": [],
        "assunto": DEFAULT_ASSUNTO,
        "corpo": DEFAULT_CORPO,
    },
}


def normalizar_nome_cliente(nome: str) -> str:
    if not nome:
        return ""
    nome_norm = unicodedata.normalize("NFKD", str(nome))
    nome_norm = nome_norm.encode("ASCII", "ignore").decode("ASCII")
    return " ".join(nome_norm.upper().split())


def formatar_brl(valor) -> str:
    if valor in (None, ""):
        return "R$ 0,00"
    try:
        num = float(valor)
    except Exception:
        return str(valor)
    texto = f"{num:,.2f}"
    texto = texto.replace(",", "X").replace(".", ",").replace("X", ".")
    return f"R$ {texto}"


def formatar_data(valor) -> str:
    if valor is None or valor == "":
        return ""
    if isinstance(valor, datetime):
        return valor.strftime("%d/%m")
    if isinstance(valor, date):
        return datetime(valor.year, valor.month, valor.day).strftime("%d/%m")
    return str(valor)


def obter_saudacao(agora: datetime | None = None) -> str:
    hora = (agora or datetime.now()).hour
    if hora < 12:
        return "bom dia"
    if hora < 18:
        return "boa tarde"
    return "boa noite"


def criar_rascunho_email_cliente(
    nome_cliente: str,
    anexos=None,
    assunto: str | None = None,
    corpo: str | None = None,
    corpo_html: str | None = None,
    abrir_rascunho: bool = False,
    dn: str | None = None,
    navio: str | None = None,
    ano: int | None = None,
    adiantamento: float | None = None,
    atracacao_ini=None,
    atracacao_fim=None,
    costs=None,
    adm=None,
):
    """
    Cria um rascunho no Outlook com base no nome do cliente.
    Permite sobrescrever assunto e corpo para casos especiais.
    """
    nome_cliente_norm = normalizar_nome_cliente(nome_cliente)
    config = CLIENTES_EMAIL.get(nome_cliente_norm)
    if not config:
        raise ValueError(f"Cliente nao encontrado: {nome_cliente_norm}")

    para = config.get("para", [])
    cc = config.get("cc", [])
    contexto = {
        "cliente": nome_cliente_norm,
        "dn": dn or "",
        "navio": navio or "",
        "ano": ano or datetime.now().year,
        "ano2": f"{(ano or datetime.now().year) % 100:02d}",
        "saudacao": obter_saudacao(),
        "adiantamento": adiantamento,
        "adiantamento_fmt": formatar_brl(adiantamento),
        "atracacao_ini": formatar_data(atracacao_ini),
        "atracacao_fim": formatar_data(atracacao_fim),
        "costs": costs,
        "adm": adm,
        "costs_fmt": formatar_brl(costs),
        "adm_fmt": formatar_brl(adm),
    }
    assunto_final = (assunto or config.get("assunto") or DEFAULT_ASSUNTO).format(
        **contexto
    )
    corpo_final = (corpo or config.get("corpo") or DEFAULT_CORPO).format(
        **contexto
    )
    corpo_html_final = (
        corpo_html or config.get("corpo_html") or DEFAULT_CORPO_HTML
    ).format(**contexto)

    try:
        pythoncom.CoInitialize()
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)

        mail.Subject = assunto_final
        if corpo_html_final:
            mail.HTMLBody = corpo_html_final
        else:
            mail.Body = corpo_final
        mail.To = "; ".join(para)
        mail.CC = "; ".join(cc)

        if anexos:
            for anexo in anexos:
                anexo_path = Path(anexo)
                if anexo_path.exists():
                    mail.Attachments.Add(str(anexo_path))
                else:
                    print(f"Arquivo nao encontrado: {anexo_path}")

        mail.Save()
        if abrir_rascunho:
            mail.Display(True)

        return True
    finally:
        pythoncom.CoUninitialize()
