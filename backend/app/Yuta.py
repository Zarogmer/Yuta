from yuta_helpers import *
from classes import (
    FaturamentoCompleto,
    FaturamentoAtipico,
    FaturamentoDeAcordo,
    ProgramaCopiarPeriodo,
    ProgramaRemoverPeriodo,
    FaturamentoSaoSebastiao,
    GerarRelatorio,
    CentralSanport,
)
from desktop_app import run_desktop

__all__ = [
    "FaturamentoCompleto",
    "FaturamentoAtipico",
    "FaturamentoDeAcordo",
    "ProgramaCopiarPeriodo",
    "ProgramaRemoverPeriodo",
    "FaturamentoSaoSebastiao",
    "GerarRelatorio",
    "CentralSanport",
]


if __name__ == "__main__":
    validar_licenca()
    run_desktop()
