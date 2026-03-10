import sys
from pathlib import Path

# Garante imports do pacote backend.app ao rodar este arquivo diretamente.
ROOT_DIR = Path(__file__).resolve().parents[2]
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

from backend.app.ui.desktop_app import run_desktop
from backend.app.yuta_helpers import validar_licenca


def main() -> None:
    validar_licenca()
    run_desktop()


if __name__ == "__main__":
    main()
