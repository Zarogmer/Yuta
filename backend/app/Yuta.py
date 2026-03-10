from pathlib import Path
import sys


# Suporta execucao direta: `python backend/app/Yuta.py`
PROJ_ROOT = Path(__file__).resolve().parents[2]
if str(PROJ_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJ_ROOT))

from backend.app.main import main


if __name__ == "__main__":
    main()

    
