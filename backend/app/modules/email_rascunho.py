import pythoncom
import win32com.client
from pathlib import Path
from datetime import datetime, date
import unicodedata
import re

from backend.app.config_manager import obter_caminho_assinatura_usuario


DEFAULT_ASSUNTO = "FATURAMENTO SANTOS {dn}/{ano} - M/V {navio}"
ASSUNTO_SAO_SEBASTIAO = "FATURAMENTO {dn}/{ano2} - M/V {navio} - PORTO DE SÃƒO SEBASTIÃƒO"
ASSUNTO_CARGONAVE = "ADIANTAMENTO / DADOS - M/V {navio}"
ASSUNTO_ROCHAMAR = "SOLICITAR OC - M/V {navio}"
CC_FIXO = ["financeiro@sanportlogistica.com.br"]
DEFAULT_CORPO = """Prezados, {saudacao}!

Seguem anexos faturamento e folhas OGMO do navio {navio} em referÃªncia.

Obs.: Gentileza notar alteraÃ§Ã£o nos dados bancÃ¡rios.

Dados para depÃ³sito:

Banco ItaÃº

AgÃªncia: 0447

Conta Corrente: 99807-1

Pix: 24.845.408/0001-22

Atenciosamente,
"""

DEFAULT_CORPO_HTML = """<div style="font-family: Arial, sans-serif; font-size: 12pt; color: #000;">
    <p>Prezados, {saudacao}!</p>
    <p>Seguem anexos faturamento e folhas OGMO do navio {navio} em referÃªncia.</p>
    <p><strong>Obs.:</strong> Gentileza notar alteraÃ§Ã£o nos dados bancÃ¡rios.</p>
    <table style="border: 1px solid #000; border-collapse: collapse; margin: 6px 0 10px;" cellpadding="6" cellspacing="0">
        <tr>
            <td style="padding: 6px 10px;">
                <div><strong>Dados para depÃ³sito:</strong></div>
                <div>Banco ItaÃº</div>
                <div>AgÃªncia: 0447</div>
                <div>Conta Corrente: 99807-1</div>
                <div>Pix: 24.845.408/0001-22</div>
            </td>
        </tr>
    </table>
    <p>Atenciosamente,</p>
</div>
"""

CORPO_SAO_SEBASTIAO = """Prezados, {saudacao}!

Segue anexo faturamento do navio {navio} em referÃªncia.

Atenciosamente,
"""

CORPO_SAO_SEBASTIAO_HTML = """<div style="font-family: Arial, sans-serif; font-size: 12pt; color: #000;">
    <p>Prezados, {saudacao}!</p>
    <p>Segue anexo faturamento do navio {navio} em referÃªncia.</p>
    <p>Atenciosamente,</p>
</div>
"""

CORPO_CARGONAVE = """Prezados, {saudacao}!

Gentileza nos enviar dados para emissÃ£o de nota fiscal do navio {navio} em referÃªncia.

Solicitamos remessa na conta corrente abaixo no valor de {adiantamento_fmt} conforme acordo.

Dados para depÃ³sito:

Banco ItaÃº

AgÃªncia: 0447

Conta Corrente: 99807-1

Pix: 24.845.408/0001-22

Desde jÃ¡ muito obrigado.

Atenciosamente,
"""

CORPO_CARGONAVE_HTML = """<div style="font-family: Arial, sans-serif; font-size: 12pt; color: #000;">
    <p>Prezados, {saudacao}!</p>
    <p>Gentileza nos enviar dados para emissÃ£o de nota fiscal do navio {navio} em referÃªncia.</p>
    <p>Solicitamos remessa na conta corrente abaixo no valor de <strong>{adiantamento_fmt}</strong> conforme acordo.</p>
    <table style="border: 1px solid #000; border-collapse: collapse; margin: 6px 0 10px;" cellpadding="6" cellspacing="0">
        <tr>
            <td style="padding: 6px 10px;">
                <div><strong>Dados para depÃ³sito:</strong></div>
                <div>Banco ItaÃº</div>
                <div>AgÃªncia: 0447</div>
                <div>Conta Corrente: 99807-1</div>
                <div>Pix: 24.845.408/0001-22</div>
            </td>
        </tr>
    </table>
    <p>Desde jÃ¡ muito obrigado.</p>
    <p>Atenciosamente,</p>
</div>
"""

CORPO_ROCHAMAR = """Prezados, {saudacao}!

Solicitamos o nÃºmero da OC do navio em referÃªncia e seguem valores da fatura abaixo:

M/V {navio}

AtracaÃ§Ã£o: {atracacao_ini} a {atracacao_fim}

Despesas OGMO: {costs_fmt}

Taxa Administrativa: {adm_fmt}

Atenciosamente,
"""

CORPO_ROCHAMAR_HTML = """<div style="font-family: Arial, sans-serif; font-size: 12pt; color: #000;">
    <p>Prezados, {saudacao}!</p>
    <p>Solicitamos o nÃºmero da OC do navio em referÃªncia e seguem valores da fatura abaixo:</p>
    <p><strong>M/V {navio}</strong></p>
    <p>AtracaÃ§Ã£o: {atracacao_ini} a {atracacao_fim}</p>
    <p>Despesas OGMO: {costs_fmt}</p>
    <p>Taxa Administrativa: {adm_fmt}</p>
    <p>Atenciosamente,</p>
</div>
"""

# Edite aqui para personalizar por cliente.
CLIENTES_EMAIL = {
    "CARGONAVE": {
        "para": ["fiscal@cgnvsantos.com.br", "contabil@cgnvsantos.com.br", "solange.leandro@cgnvsantos.com.br", "financeiro@cgnvsantos.com.br"],
        "cc": ["sanport@sanportlogistica.com.br"],
        "assunto": ASSUNTO_CARGONAVE,
        "corpo": CORPO_CARGONAVE,
        "corpo_html": CORPO_CARGONAVE_HTML,
    },
    "WILSON": {
        "para": ["contasapagar.ssz@unishipping.com.br", "ineto.ssz@unishipping.com.br"],
        "cc": ["sanport@sanportlogistica.com.br"],
        "assunto": DEFAULT_ASSUNTO,
        "corpo": DEFAULT_CORPO,
    },
    "ROCHAMAR": {
        "para": ["faturas@rochamar.com", "cpagar@rochamar.com", "oprsts@rochamar.com","solicitaroc@rochamar.com"],
        "cc": ["sanport@sanportlogistica.com.br"],
        "assunto": ASSUNTO_ROCHAMAR,
        "corpo": CORPO_ROCHAMAR,
        "corpo_html": CORPO_ROCHAMAR_HTML,
    },
    "UNIMAR": {
        "para": ["contasapagar.ssz@unishipping.com.br", "ineto.ssz@unishipping.com.br"],
        "cc": ["jsilva.ssz@unishipping.com.br"],
        "assunto": DEFAULT_ASSUNTO,
        "corpo": DEFAULT_CORPO,
        "corpo_html": DEFAULT_CORPO_HTML,
    },
    "WILSON SONS - SAO SEBASTIAO": {
        "para": ["pagamento.csc@wilsonsons.com.br", "pagamento.ws@wilsonsons.com.br"],
        "cc": ["sanport@sanportlogistica.com.br"],
        "assunto": ASSUNTO_SAO_SEBASTIAO,
        "corpo": CORPO_SAO_SEBASTIAO,
        "corpo_html": CORPO_SAO_SEBASTIAO_HTML,
    },
    "AQUARIUS - PSS": {
        "para": ["fabio.cruz@aquariusoffshore.com.br", "finance@aquariusoffshore.com.br"],
        "cc": ["sanport@sanportlogistica.com.br"],
        "assunto": ASSUNTO_SAO_SEBASTIAO,
        "corpo": CORPO_SAO_SEBASTIAO,
        "corpo_html": CORPO_SAO_SEBASTIAO_HTML,
    },
    "SEA SIDE - PSS": {
        "para": ["seaside@seasidebrazil.com.br", "accounts@seasidebrazil.com.br"],
        "cc": ["sanport@sanportlogistica.com.br"],
        "assunto": ASSUNTO_SAO_SEBASTIAO,
        "corpo": CORPO_SAO_SEBASTIAO,
        "corpo_html": CORPO_SAO_SEBASTIAO_HTML,
    },
    "ARENNA": {
        "para": ["financeiro@arennalogistica.com.br"],
        "cc": ["sanport@sanportlogistica.com.br"],
        "assunto": DEFAULT_ASSUNTO,
        "corpo": DEFAULT_CORPO,
    },
    "CARGILL": {
        "para": ["Santosagencydacct@cargill.com", "Regina_Silva@cargill.com"],
        "cc": ["sanport@sanportlogistica.com.br"],
        "assunto": DEFAULT_ASSUNTO,
        "corpo": DEFAULT_CORPO,
    },
    "CONESUL": {
        "para": ["financial@conesulagencia.com.br"],
        "cc": ["sanport@sanportlogistica.com.br"],
        "assunto": DEFAULT_ASSUNTO,
        "corpo": DEFAULT_CORPO,
    },
    "DELTA": {
        "para": ["operations@deltashipping.com.br"],
        "cc": ["sanport@sanportlogistica.com.br"],
        "assunto": DEFAULT_ASSUNTO,
        "corpo": DEFAULT_CORPO,
    },
    "GEM": {
        "para": ["agency@transatlanticamaritima.com.br", 
                 "financial@transatlanticamaritima.com.br",
                 "brunoserrano@transatlanticamaritima.com.br", 
                 "fernandovalle@transatlanticamaritima.com.br"],

        "cc": ["sanport@sanportlogistica.com.br"],
        "assunto": DEFAULT_ASSUNTO,
        "corpo": DEFAULT_CORPO,
    },
    "HMS": {
        "para": ["marcos@hmsbrasil.com.br"],
        "cc": ["sanport@sanportlogistica.com.br"],
        "assunto": DEFAULT_ASSUNTO,
        "corpo": DEFAULT_CORPO,
    },
    "JBG": {
        "para": ["bruno@jbgshipping.com.br", "operations@jbgshipping.com.br"],
        "cc": ["sanport@sanportlogistica.com.br"],
        "assunto": DEFAULT_ASSUNTO,
        "corpo": DEFAULT_CORPO,
    },
    "LMA": {
        "para": ["gabriela.silva@lmashipping.com.br", "account@lmashipping.com.br" , "leandro.alves@lmashipping.com.br"],
        "cc": ["sanport@sanportlogistica.com.br"],
        "assunto": DEFAULT_ASSUNTO,
        "corpo": DEFAULT_CORPO,
    },
    "NAABSA": {
        "para": ["lais@naabsa.com", "alex@naabsa.com", "daccount@naabsa.com"],
        "cc": ["sanport@sanportlogistica.com.br"],
        "assunto": DEFAULT_ASSUNTO,
        "corpo": DEFAULT_CORPO,
    },
    "NORTH STAR": {
        "para": ["FDA@nsshipping.com.br", "Faturamento.nss@nsshipping.com.br"],
        "cc": ["sanport@sanportlogistica.com.br"],
        "assunto": DEFAULT_ASSUNTO,
        "corpo": DEFAULT_CORPO,
    },
    "PROJECT CARGO": {
        "para": ["wagner@pcargo.com.br", "financeiro@pcargo.com.br"],
        "cc": ["sanport@sanportlogistica.com.br"],
        "assunto": DEFAULT_ASSUNTO,
        "corpo": DEFAULT_CORPO,
    },
    "TRANSATLANTICA": {
        "para": ["operacional@sanportlogistica.com.br", "agency@transatlanticamaritima.com.br", 
                 "brunoserrano@transatlanticamaritima.com.br", "fernandovalle@transatlanticamaritima.com.br"],
        "cc": ["sanport@sanportlogistica.com.br"],
        "assunto": DEFAULT_ASSUNTO,
        "corpo": DEFAULT_CORPO,
    },
    "WILLIAMS": {
        "para": ["financeiro.santos@williams.com.br"],
        "cc": ["sanport@sanportlogistica.com.br"],
        "assunto": DEFAULT_ASSUNTO,
        "corpo": DEFAULT_CORPO,
    },
    "WILSON SONS": {
        "para": ["pagamento.csc@wilsonsons.com.br", "pagamento.ws@wilsonsons.com.br"],
        "cc": ["sanport@sanportlogistica.com.br"],
        "assunto": DEFAULT_ASSUNTO,
        "corpo": DEFAULT_CORPO,
    },
    "ZPORT": {
        "para": ["everton.pereira@zport.com.br", "agency@zport.com.br"],
        "cc": ["sanport@sanportlogistica.com.br"],
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


def _normalizar_nome_usuario(nome: str | None) -> str:
    if not nome:
        return ""
    nome_norm = unicodedata.normalize("NFKD", str(nome))
    nome_norm = nome_norm.encode("ASCII", "ignore").decode("ASCII")
    return " ".join(nome_norm.upper().split())


def _cid_assinatura(usuario_nome: str | None) -> str:
    base = _normalizar_nome_usuario(usuario_nome) or "ASSINATURA"
    base = re.sub(r"[^A-Z0-9]+", "_", base).strip("_")
    return f"assinatura_{base.lower() or 'usuario'}"


def _inserir_assinatura_no_final(corpo_html: str, cid: str) -> str:
    bloco = (
        f"<p style=\"margin:4px 0 0 0;\">"
        f"<img src=\"cid:{cid}\" width=\"420\" style=\"width:420px;max-width:420px;height:auto;display:block;\">"
        "</p>"
    )

    for padrao in (r"</body\s*>", r"</html\s*>", r"</div\s*>"):
        matches = list(re.finditer(padrao, corpo_html, flags=re.IGNORECASE))
        if matches:
            idx = matches[-1].start()
            return corpo_html[:idx] + bloco + corpo_html[idx:]

    return corpo_html + bloco


def _normalizar_lista_emails(lista_emails) -> list[str]:
    resultado = []
    if not lista_emails:
        return resultado

    for item in lista_emails:
        partes = re.split(r"[;,]", str(item or ""))
        for parte in partes:
            email = parte.strip()
            if email:
                resultado.append(email)

    return resultado


def _mesclar_cc(*listas_cc) -> list[str]:
    resultado = []
    vistos = set()
    for lista in listas_cc:
        for email in _normalizar_lista_emails(lista):
            email_limpo = str(email).strip()
            if not email_limpo:
                continue
            chave = email_limpo.lower()
            if chave in vistos:
                continue
            vistos.add(chave)
            resultado.append(email_limpo)
    return resultado


def _corrigir_mojibake_texto(texto: str | None) -> str:
    if texto is None:
        return ""
    s = str(texto)

    # Tentativa generica de recuperar texto UTF-8 que foi lido como latin-1.
    if "Ã" in s or "Â" in s or "â" in s:
        try:
            rec = s.encode("latin-1").decode("utf-8")
            if rec:
                s = rec
        except Exception:
            pass

    trocas = {
        "NÃ£o": "Não",
        "nÃ£o": "não",
        "NÃºmero": "Número",
        "nÃºmero": "número",
        "referÃªncia": "referência",
        "alteraÃ§Ã£o": "alteração",
        "depÃ³sito": "depósito",
        "AgÃªncia": "Agência",
        "emissÃ£o": "emissão",
        "jÃ¡": "já",
        "AtracaÃ§Ã£o": "Atracação",
        "SÃƒO": "SÃO",
    }
    for antigo, novo in trocas.items():
        s = s.replace(antigo, novo)

    return s


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
    usuario_nome: str | None = None,
):
    """
    Cria um rascunho no Outlook com base no nome do cliente.
    Permite sobrescrever assunto e corpo para casos especiais.
    """
    nome_cliente_norm = normalizar_nome_cliente(nome_cliente)
    config = CLIENTES_EMAIL.get(nome_cliente_norm)
    if not config:
        raise ValueError(f"Cliente nao encontrado: {nome_cliente_norm}")

    para = _normalizar_lista_emails(config.get("para", []))
    cc = _mesclar_cc(config.get("cc", []), CC_FIXO)
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

    assunto_final = _corrigir_mojibake_texto(assunto_final)
    corpo_final = _corrigir_mojibake_texto(corpo_final)
    corpo_html_final = _corrigir_mojibake_texto(corpo_html_final)

    try:
        pythoncom.CoInitialize()
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)

        mail.Subject = assunto_final
        caminho_assinatura = obter_caminho_assinatura_usuario(usuario_nome or "")

        if corpo_html_final:
            if caminho_assinatura:
                cid = _cid_assinatura(usuario_nome)
                corpo_html_com_assinatura = _inserir_assinatura_no_final(
                    corpo_html_final,
                    cid,
                )
                mail.HTMLBody = corpo_html_com_assinatura
                try:
                    anexo_ass = mail.Attachments.Add(str(caminho_assinatura))
                    anexo_ass.PropertyAccessor.SetProperty(
                        "http://schemas.microsoft.com/mapi/proptag/0x3712001F",
                        cid,
                    )
                except Exception as e:
                    print(f"âš ï¸ Falha ao embutir assinatura de e-mail: {e}")
                    mail.HTMLBody = corpo_html_final
            else:
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

