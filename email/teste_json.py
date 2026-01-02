from pathlib import Path

modelo_path = Path("email") / "modelo_email.json"
with open(modelo_path, "r", encoding="utf-8") as f:
    conteudo = f.read()

print("📄 Conteúdo bruto do arquivo:")
print(repr(conteudo))
