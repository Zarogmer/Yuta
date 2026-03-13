import os
import sys
from pathlib import Path


def _resource_base_path() -> Path:
    if getattr(sys, "frozen", False):
        return Path(getattr(sys, "_MEIPASS", Path(sys.executable).resolve().parent))
    return Path(__file__).resolve().parents[3]


def app_base_path() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parents[3]


def resource_path(relative_path: str) -> str:
    return str(_resource_base_path() / relative_path)


def project_root_path() -> Path:
    if getattr(sys, "frozen", False):
        return app_base_path()
    return Path(__file__).resolve().parents[3]


def _valid_poppler_bin(path: Path) -> bool:
    return path.exists() and (path / "pdfinfo.exe").exists() and (path / "pdftoppm.exe").exists()


def poppler_paths_candidatos() -> list[Path]:
    candidatos = []

    env_poppler = os.environ.get("POPPLER_PATH")
    if env_poppler:
        env_path = Path(env_poppler)
        candidatos.append(env_path)
        candidatos.append(env_path / "Library" / "bin")
        candidatos.append(env_path / "bin")

    base_dirs = [_resource_base_path(), app_base_path()]
    for base in base_dirs:
        candidatos.extend(
            [
                base / "poppler" / "Library" / "bin",
                base / "poppler" / "bin",
            ]
        )

    path_env = os.environ.get("PATH", "")
    for parte in path_env.split(os.pathsep):
        if parte and "poppler" in parte.lower():
            candidatos.append(Path(parte))

    candidatos.extend(
        [
            Path(r"C:\poppler-25.12.0\Library\bin"),
            Path(r"C:\poppler\Library\bin"),
            Path(r"C:\Program Files\poppler\Library\bin"),
            Path(r"C:\Program Files (x86)\poppler\Library\bin"),
        ]
    )

    vistos = set()
    validos = []
    for pasta in candidatos:
        chave = str(pasta).lower().strip()
        if not chave or chave in vistos:
            continue
        vistos.add(chave)
        if _valid_poppler_bin(pasta):
            validos.append(pasta)
    return validos


def configurar_tesseract_runtime() -> Path | None:
    import pytesseract

    candidatos = []

    env_tesseract = os.environ.get("TESSERACT_EXE")
    if env_tesseract:
        candidatos.append(Path(env_tesseract))

    for base in (_resource_base_path(), app_base_path()):
        candidatos.append(base / "tesseract" / "tesseract.exe")

    candidatos.extend(
        [
            Path(r"C:\Program Files\Tesseract-OCR\tesseract.exe"),
            Path(r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe"),
        ]
    )

    vistos = set()
    for exe in candidatos:
        chave = str(exe).lower().strip()
        if not chave or chave in vistos:
            continue
        vistos.add(chave)

        if exe.exists():
            pytesseract.pytesseract.tesseract_cmd = str(exe)
            tessdata_dir = exe.parent / "tessdata"
            if tessdata_dir.exists():
                os.environ["TESSDATA_PREFIX"] = str(tessdata_dir)
            return exe

    return None
