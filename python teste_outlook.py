from outlook_mailer import enviar_email_outlook

enviar_email_outlook(
    modelo="sanport",
    contexto={
        "numero": 999,
        "ano": 2025,
        "navio": "NAVIO TESTE"
    },
    anexos=[],
    mostrar=True
)
