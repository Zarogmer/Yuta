import pythoncom
import win32com.client
from pathlib import Path
import json
import time

def enviar_email_outlook(modelo, contexto, anexos=None, mostrar=False):
    try:
        pythoncom.CoInitialize()  # Inicializa COM na thread atual

        BASE_DIR = Path(__file__).resolve().parent
        EMAIL_DIR = BASE_DIR / "email"

        config_path = EMAIL_DIR / f"{modelo}.json"
        if not config_path.exists():
            raise FileNotFoundError(f"Arquivo de email não encontrado: {config_path}")

        with open(config_path, encoding="utf-8") as f:
            cfg = json.load(f)

        outlook = win32com.client.DispatchEx("Outlook.Application")
        mail = outlook.CreateItem(0)

        mail.To = ";".join(cfg["para"])
        mail.CC = ";".join(cfg.get("cc", []))
        mail.Subject = cfg["assunto"].format(**contexto)
        mail.Body = cfg["corpo"].format(**contexto)

        for anexo in anexos or []:
            mail.Attachments.Add(str(anexo))

        if mostrar:
            mail.Display()  # abre o rascunho para ver
        else:
            mail.Save()     # salva diretamente em rascunhos
            # Às vezes um pequeno delay ajuda
            time.sleep(0.3)

        print("📧 Email criado no Outlook e salvo em RASCUNHOS com sucesso!")

    except Exception as e:
        print("❌ Erro ao enviar email:", e)
    finally:
        pythoncom.CoUninitialize()  # Libera COM
