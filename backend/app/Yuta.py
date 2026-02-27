from yuta_helpers import *
from classes import (
    CriarPasta,
    FaturamentoAtipico,
    FaturamentoCompleto,
    FaturamentoDeAcordo,
    GerarRelatorio,
    FaturamentoSaoSebastiao,
    FazerPonto,
    ProgramaRemoverPeriodo,
)
from desktop_app import run_desktop

__all__ = [
    "FaturamentoCompleto",
    "FaturamentoAtipico",
    "FaturamentoDeAcordo",
    "FazerPonto",
    "ProgramaRemoverPeriodo",
    "FaturamentoSaoSebastiao",
    "GerarRelatorio",
    "CriarPasta",
]


if __name__ == "__main__":
    validar_licenca()
    run_desktop()