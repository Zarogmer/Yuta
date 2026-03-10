try:
    from .utils.yuta_helpers import *
except ModuleNotFoundError as exc:
    # Fallback apenas para execucao legada fora do pacote backend.app.
    if exc.name in {
        "backend",
        "backend.app",
        "backend.app.utils",
        "backend.app.utils.yuta_helpers",
    }:
        from utils.yuta_helpers import *
    else:
        raise
