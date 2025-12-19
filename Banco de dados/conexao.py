import sqlite3
from pathlib import Path

root = Path(__file__).parent
con = sqlite3.connect(root / "banco_de_dados.sqlite")
cur = con.cursor()

clientes = [
    ("Cargonave", "cargonave@gmail.com"),
    ("Wilson", "wilson@gmail.com"),
    ("Rochamar", "rochamar@gmail.com")
]

for nome, email in clientes:
    try:
        cur.execute(
            "INSERT INTO clientes (nome, email) VALUES (?, ?)",
            (nome, email)
        )
    except sqlite3.IntegrityError:
        pass  # email já existe, ignora

con.commit()
con.close()
