import pythoncom
import win32com.client
from pathlib import Path
import time
from datetime import datetime
import json


def enviar_email_outlook(dn: str, nome_navio: str, anexos=None):
    """
    Cria um email no Outlook usando sempre o modelo padrão 'modelo_email.json'.
    O DN e o nome do navio entram apenas nos placeholders do assunto/corpo.
    """
    try:
        pythoncom.CoInitialize()

        # 🔧 Sempre usa modelo padrão
        modelo_path = Path("email") / "modelo_email.json"
        if not modelo_path.exists():
            raise FileNotFoundError(f"❌ Arquivo de email não encontrado: {modelo_path}")

        with open(modelo_path, "r", encoding="utf-8") as f:
            modelo = json.load(f)

        # Contexto dinâmico
        contexto = {
            "dn": dn,
            "ano": datetime.now().year,
            "navio": nome_navio
        }

        # 📨 Criar email
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)

        mail.Subject = modelo["assunto"].format(**contexto)
        mail.HTMLBody = modelo["corpo_html"].format(**contexto)

        mail.To = "; ".join(modelo["para"])
        mail.CC = "; ".join(modelo.get("cc", []))

        # 📎 Anexos
        if anexos:
            for anexo in anexos:
                anexo_path = Path(anexo)
                if anexo_path.exists():
                    mail.Attachments.Add(str(anexo_path))
                else:
                    print(f"⚠️ Arquivo não encontrado: {anexo_path}")

        # 💾 Salvar em Rascunhos
        mail.Save()
        time.sleep(0.3)
        print("✅ Email criado e salvo em RASCUNHOS com sucesso!")

    except Exception as e:
        print("❌ Erro ao enviar email:", e)
    finally:
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    dn = "123"  # exemplo de DN
    nome_navio = "NAVIOOO"  # exemplo de nome do navio

    # 🔹 Pasta onde estão os arquivos
    pasta_navio = Path(r"C:\Users\Guilherme\Desktop\UNIMAR\123 - naviooo")

    # 🔹 Nome padrão para Excel e PDF
    nome_arquivo_excel = f"FATURAMENTO - DN {dn} - MV {nome_navio}.xlsx"
    nome_arquivo_pdf = f"FATURAMENTO - DN {dn} - MV {nome_navio}.pdf"

    # 🔹 Montar caminhos completos
    caminho_excel = pasta_navio / nome_arquivo_excel
    caminho_pdf = pasta_navio / nome_arquivo_pdf

    anexos = [caminho_excel, caminho_pdf]

    enviar_email_outlook(dn, nome_navio, anexos)
