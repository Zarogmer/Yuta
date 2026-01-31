from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware

app = FastAPI(title="Yuta API")

app.add_middleware(
    CORSMiddleware,
    allow_origins=[
    "http://127.0.0.1:5500",
    "http://localhost:5500",
    ],

    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

OPCOES_MENU = [
    "FATURAMENTO",
    "FATURAMENTO SÃO SEBASTIÃO",
    "DE ACORDO",
    "FAZER PONTO",
    "DESFAZER PONTO - X",
    "RELATÓRIO - X",
    "SAIR DO PROGRAMA"
]

@app.get("/menu")
def menu():
    return {"opcoes": OPCOES_MENU}

@app.post("/menu/acao/{indice}")
def executar_acao(indice: int):
    mapa = {
        0: "Executaria: FaturamentoCompleto().executar()",
        1: "Executaria: FaturamentoSaoSebastiao().executar()",
        2: "Executaria: FaturamentoDeAcordo().executar()",
        3: "Executaria: ProgramaCopiarPeriodo().executar()",
        4: "Executaria: ProgramaRemoverPeriodo().executar()",
        5: "Executaria: GerarRelatorio().executar()",
        6: "Sair (na web apenas responde).",
    }
    if indice < 0 or indice >= len(OPCOES_MENU):
        return {"ok": False, "erro": "Índice inválido", "indice": indice}

    return {
        "ok": True,
        "indice": indice,
        "opcao": OPCOES_MENU[indice],
        "msg": mapa.get(indice, "Ação ainda não implementada"),
    }

@app.get("/")
def root():
    return {"msg": "Yuta API online (front separado)"}
