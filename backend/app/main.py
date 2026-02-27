from __future__ import annotations

import sys
from pathlib import Path
from threading import Lock
from typing import Callable

from fastapi import FastAPI
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles

ROOT_DIR = Path(__file__).resolve().parents[2]
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

from backend.app import (
    FaturamentoAtipico,
    FaturamentoCompleto,
    FaturamentoDeAcordo,
    FaturamentoSaoSebastiao,
    GerarRelatorio,
    FazerPonto,
    ProgramaRemoverPeriodo,
    CriarPasta,
    validar_licenca,
)

app = FastAPI(title="Yuta API")

_LOCK = Lock()


def _resource_path(relative_path: str) -> Path:
    base_dir = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parents[2]))
    return base_dir / relative_path


STATIC_DIR = _resource_path("frontend/static")
INDEX_FILE = STATIC_DIR / "index.html"

if STATIC_DIR.exists():
    app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")


OPCOES_MENU = [
    "FATURAMENTO (NORMAL)",
    "FATURAMENTO (ATIPICO)",
    "FATURAMENTO SAO SEBASTIAO",
    "DE ACORDO",
    "FAZER PONTO",
    "DESFAZER PONTO - X",
    "RELATORIO - X",
    "CRIAR PASTA",
    "SAIR DO PROGRAMA",
]


def _executar_acao(indice: int) -> dict:
    if indice == 0:
        FaturamentoCompleto().executar()
        return {"msg": "Faturamento (Normal) finalizado"}
    if indice == 1:
        FaturamentoAtipico().executar()
        return {"msg": "Faturamento (Atipico) finalizado"}
    if indice == 2:
        FaturamentoSaoSebastiao().executar()
        return {"msg": "Faturamento Sao Sebastiao finalizado"}
    if indice == 3:
        FaturamentoDeAcordo().executar()
        return {"msg": "De Acordo finalizado"}
    if indice == 4:
        FazerPonto(debug=True).executar()
        return {"msg": "Fazer Ponto finalizado"}
    if indice == 5:
        ProgramaRemoverPeriodo(debug=True).executar()
        return {"msg": "Desfazer Ponto finalizado"}
    if indice == 6:
        if hasattr(GerarRelatorio, "executar"):
            GerarRelatorio().executar()
            return {"msg": "Relatorio finalizado"}
        return {"msg": "Relatorio nao implementado"}
    if indice == 7:
        CriarPasta().executar()
        return {"msg": "Criar Pasta finalizado"}
    if indice == 8:
        return {"msg": "Saindo (acao ignorada na web)"}

    raise ValueError("Indice invalido")


@app.get("/api/menu")
def menu():
    return {"opcoes": OPCOES_MENU}


@app.post("/api/menu/acao/{indice}")
def executar_acao(indice: int):
    if indice < 0 or indice >= len(OPCOES_MENU):
        return JSONResponse(
            status_code=400,
            content={"ok": False, "erro": "Indice invalido", "indice": indice},
        )

    if not _LOCK.acquire(blocking=False):
        return JSONResponse(
            status_code=409,
            content={"ok": False, "erro": "Outra acao ja esta em execucao"},
        )

    try:
        validar_licenca()
        resultado = _executar_acao(indice)
        return {"ok": True, "indice": indice, "opcao": OPCOES_MENU[indice], **resultado}
    except Exception as exc:
        return JSONResponse(
            status_code=500,
            content={"ok": False, "erro": str(exc), "indice": indice},
        )
    finally:
        _LOCK.release()


@app.get("/api/health")
def health():
    return {"ok": True}


@app.get("/")
def root():
    if INDEX_FILE.exists():
        return FileResponse(str(INDEX_FILE))
    return {"msg": "Frontend nao encontrado"}


def _abrir_browser():
    try:
        import webbrowser

        webbrowser.open("http://127.0.0.1:8000/")
    except Exception:
        pass


if __name__ == "__main__":
    _abrir_browser()
    import uvicorn

    uvicorn.run(app, host="127.0.0.1", port=8000)
