import pythoncom
import win32com.client
from pathlib import Path
from datetime import datetime, date
import unicodedata
import re
from html import escape

from backend.app.config_manager import obter_caminho_assinatura_usuario


EMAIL_FONT_FAMILY = "Aptos, Calibri, Arial, sans-serif"
EMAIL_HTML_STYLE = f"font-family: {EMAIL_FONT_FAMILY}; font-size: 11pt; color: #000;"

DEFAULT_ASSUNTO = "FATURAMENTO SANTOS {dn}/{ano} - M/V {navio}"
ASSUNTO_SAO_SEBASTIAO = "FATURAMENTO {dn}/{ano2} - M/V {navio} - PORTO DE SÃO SEBASTIÃO"
ASSUNTO_CARGONAVE = "ADIANTAMENTO / DADOS - M/V {navio}"
ASSUNTO_ROCHAMAR = "SOLICITAR OC - M/V {navio}"
CC_FIXO = ["financeiro@sanportlogistica.com.br"]
DEFAULT_CORPO = """Prezados, bom dia!

Segue anexo faturamento do navio M/V {navio} em referência.

Atenciosamente,
"""

DEFAULT_CORPO_HTML = """<div style="{style}">
    <p>Prezados, bom dia!</p>
    <p>Segue anexo faturamento do navio <strong>M/V {{navio}}</strong> em referência.</p>
    <p>Atenciosamente,</p>
</div>
""".format(style=EMAIL_HTML_STYLE)

CORPO_SAO_SEBASTIAO = """Prezados, bom dia!

Segue anexo faturamento do navio M/V {navio} em referência.

Atenciosamente,
"""

CORPO_SAO_SEBASTIAO_HTML = """<div style="{style}">
    <p>Prezados, bom dia!</p>
    <p>Segue anexo faturamento do navio <strong>M/V {{navio}}</strong> em referência.</p>
    <p>Atenciosamente,</p>
</div>
""".format(style=EMAIL_HTML_STYLE)

CORPO_DE_ACORDO = (
    "Prezados, {saudacao}!\n\n"
    "Confirmo que o de acordo foi devidamente concedido.\n\n"
    "Conforme procedimento padrão, haverá cobrança no valor de R$ 500,00.\n\n"
    "Gentileza nos avisar assim que o navio deixar o Porto de Santos, para que possamos informar a Autoridade Portuaria sobre a inexistencia de operacao de carga.\n"
    "Ressaltamos que o sistema da Supervia possui uma nova funcionalidade que obriga o Operador Portuário a informar, via sistema, quando o navio não realiza operações durante sua estadia no porto.\n\n"
    "A ausência dessa informação gera automaticamente um Auto de Inspeção, bem como as respectivas penalidades ao Operador Portuário.\n\n"
    "Certo de sua colaboração e entendimento, agradecemos a parceria.\n\n"
    "Atenciosamente,\n"
)

CORPO_DE_ACORDO_HTML = """<div style="{style}">
    <p>Prezados, {{saudacao}}!</p>
    <p>Confirmo que o de acordo foi devidamente concedido.</p>
    <p>Conforme procedimento padrão, haverá cobrança no valor de <strong>R$ 500,00</strong>.</p>
    <p>Gentileza nos avisar assim que o navio deixar o Porto de Santos, para que possamos informar a Autoridade Portuária sobre a inexistência de operação de carga.</p>
    <p>Ressaltamos que o sistema da Supervia possui uma nova funcionalidade que obriga o Operador Portuário a informar, via sistema, quando o navio não realiza operações durante sua estadia no porto.</p>
    <p>A ausência dessa informação gera automaticamente um Auto de Inspeção, bem como as respectivas penalidades ao Operador Portuário.</p>
    <p>Certo de sua colaboração e entendimento, agradecemos a parceria.</p>
    <p>Atenciosamente,</p>
</div>
""".format(style=EMAIL_HTML_STYLE)

TIPOS_EMAIL = {
    "DE_ACORDO": {
        "corpo": CORPO_DE_ACORDO,
        "corpo_html": CORPO_DE_ACORDO_HTML,
    },
}

CORPO_CARGONAVE = """Prezados, {saudacao}!

Gentileza nos enviar dados para emissão de nota fiscal do navio {navio} em referência.

Solicitamos remessa na conta corrente abaixo no valor de {adiantamento_fmt} conforme acordo.

Dados para depósito:

Banco ItaÃº

Agência: 0447

Conta Corrente: 99807-1

Pix: 24.845.408/0001-22

Desde já muito obrigado.

Atenciosamente,
"""

CORPO_CARGONAVE_HTML = """<div style="{style}">
    <p>Prezados, {{saudacao}}!</p>
    <p>Gentileza nos enviar dados para emissão de nota fiscal do navio {{navio}} em referência.</p>
    <p>Solicitamos remessa na conta corrente abaixo no valor de <strong>{{adiantamento_fmt}}</strong> conforme acordo.</p>
    <table style="border: 1px solid #000; border-collapse: collapse; margin: 6px 0 10px;" cellpadding="6" cellspacing="0">
        <tr>
            <td style="padding: 6px 10px;">
                <div><strong>Dados para depósito:</strong></div>
                <div>Banco ItaÃº</div>
                <div>Agência: 0447</div>
                <div>Conta Corrente: 99807-1</div>
                <div>Pix: 24.845.408/0001-22</div>
            </td>
        </tr>
    </table>
    <p>Desde já muito obrigado.</p>
    <p>Atenciosamente,</p>
</div>
""".format(style=EMAIL_HTML_STYLE)

CORPO_ROCHAMAR = """Prezados, {saudacao}!

Solicitamos o número da OC do navio em referência e seguem valores da fatura abaixo:

M/V {navio}

Atracação: {atracacao_ini} a {atracacao_fim}

Despesas OGMO: {costs_fmt}

Taxa Administrativa: {adm_fmt}

Atenciosamente,
"""

CORPO_ROCHAMAR_HTML = """<div style="{style}">
    <p>Prezados, {{saudacao}}!</p>
    <p>Solicitamos o número da OC do navio em referência e seguem valores da fatura abaixo:</p>
    <p><strong>M/V {{navio}}</strong></p>
    <p>Atracação: {{atracacao_ini}} a {{atracacao_fim}}</p>
    <p>Despesas OGMO: {{costs_fmt}}</p>
    <p>Taxa Administrativa: {{adm_fmt}}</p>
    <p>Atenciosamente,</p>
</div>
""".format(style=EMAIL_HTML_STYLE)

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
        f"<p style=\"margin:12px 0 0 0;\">"
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

    # Tentativa generica de recuperar texto UTF-8 que foi lido como cp1252/latin-1.
    if "Ã" in s or "Â" in s or "â" in s:
        for origem in ("cp1252", "latin-1"):
            try:
                rec = s.encode(origem).decode("utf-8")
            except Exception:
                continue

            if rec:
                marcador_atual = s.count("Ã") + s.count("Â") + s.count("â")
                marcador_rec = rec.count("Ã") + rec.count("Â") + rec.count("â")
                if marcador_rec <= marcador_atual:
                    s = rec
                    break

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
        "ItaÃº": "Itaú",
        "SÃƒO": "SÃO",
        "SEBASTIÃƒO": "SEBASTIÃO",
        "Portuaria": "Portuária",
        "Portuario": "Portuário",
        "operacao": "operação",
        "operacoes": "operações",
        "inexistencia": "inexistência",
        "ausencia": "ausência",
        "informacao": "informação",
        "Inspecao": "Inspeção",
        "colaboracao": "colaboração",
        "padrao": "padrão",
        "havera": "haverá",
        "cobranca": "cobrança",
    }
    for antigo, novo in trocas.items():
        s = s.replace(antigo, novo)

    return s


def _converter_texto_para_html(texto: str) -> str:
    paragrafos = [p.strip() for p in str(texto or "").split("\n\n") if p.strip()]
    if not paragrafos:
        return f"<div style=\"{EMAIL_HTML_STYLE}\"></div>"

    html_paragrafos = []
    for paragrafo in paragrafos:
        conteudo = escape(paragrafo).replace("\n", "<br>")
        html_paragrafos.append(f"<p>{conteudo}</p>")

    return f"<div style=\"{EMAIL_HTML_STYLE}\">{''.join(html_paragrafos)}</div>"


def criar_rascunho_email_cliente(
    nome_cliente: str,
    anexos=None,
    assunto: str | None = None,
    corpo: str | None = None,
    corpo_html: str | None = None,
    tipo_email: str | None = None,
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
    tipo_email_norm = " ".join(str(tipo_email or "").upper().split()).replace(" ", "_")
    template_tipo = TIPOS_EMAIL.get(tipo_email_norm, {})

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
    corpo_template = corpo or template_tipo.get("corpo") or config.get("corpo") or DEFAULT_CORPO
    corpo_html_template = corpo_html
    if corpo_html_template is None:
        corpo_html_template = template_tipo.get("corpo_html")
    if corpo_html_template is None:
        corpo_html_template = config.get("corpo_html")

    corpo_final = corpo_template.format(**contexto)
    if corpo_html_template is not None:
        corpo_html_final = corpo_html_template.format(**contexto)
    else:
        corpo_html_final = _converter_texto_para_html(corpo_final)

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

