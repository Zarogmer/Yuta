"""
Gerenciador de configuração do Yuta
Permite configurar caminhos específicos por computador
"""

import json
from pathlib import Path
from typing import Dict, Any
import os
import unicodedata


def _obter_caminho_config() -> Path:
    """Retorna o caminho do arquivo de configuração"""
    # Tenta encontrar a raiz do projeto (onde está config.json)
    current = Path(__file__).resolve()
    
    # Sobe até encontrar config.json ou chegar na raiz
    for parent in [current.parent] + list(current.parents):
        config_path = parent / "config.json"
        if config_path.exists():
            return config_path
    
    # Se não encontrar, cria na pasta do projeto (2 níveis acima de app/)
    return current.parent.parent / "config.json"


def _carregar_config() -> Dict[str, Any]:
    """Carrega a configuração do arquivo JSON"""
    caminho = _obter_caminho_config()
    
    if not caminho.exists():
        # Retorna configuração padrão
        return {
            "caminho_base_faturamentos": "",
            "auto_detectar": True
        }
    
    try:
        with open(caminho, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception as e:
        print(f"⚠️ Erro ao ler config.json: {e}")
        return {
            "caminho_base_faturamentos": "",
            "auto_detectar": True
        }


def _salvar_config(config: Dict[str, Any]) -> None:
    """Salva a configuração no arquivo JSON"""
    caminho = _obter_caminho_config()
    
    try:
        with open(caminho, "w", encoding="utf-8") as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
        print(f"✅ Configuração salva em: {caminho}")
    except Exception as e:
        print(f"❌ Erro ao salvar config.json: {e}")


def _auto_detectar_base_faturamentos() -> Path | None:
    """Tenta detectar automaticamente a pasta base de faturamentos em diferentes padrões."""

    def _normalizar_texto(valor: str) -> str:
        texto = unicodedata.normalize("NFKD", str(valor or ""))
        texto = "".join(ch for ch in texto if not unicodedata.combining(ch))
        return texto.upper().strip()

    def _eh_base_faturamentos(path: Path) -> bool:
        nome = _normalizar_texto(path.name)
        pai_nome = _normalizar_texto(path.parent.name) if path.parent else ""

        if "FATURAMENTOS" not in nome:
            return False

        if "CENTRAL DE DOCUMENTOS" in nome:
            return True

        if nome.startswith("01"):
            return True

        if "CENTRAL DE DOCUMENTOS" in pai_nome:
            return True

        return False

    home = Path.home()

    raiz_usuario = [home]

    userprofile = os.environ.get("USERPROFILE")
    if userprofile:
        raiz_usuario.append(Path(userprofile))

    onedrive = os.environ.get("OneDrive")
    if onedrive:
        raiz_usuario.append(Path(onedrive))

    onedrive_consumer = os.environ.get("OneDriveConsumer")
    if onedrive_consumer:
        raiz_usuario.append(Path(onedrive_consumer))

    onedrive_commercial = os.environ.get("OneDriveCommercial")
    if onedrive_commercial:
        raiz_usuario.append(Path(onedrive_commercial))

    candidatos_raiz = []
    vistos = set()
    for raiz in raiz_usuario:
        if not raiz:
            continue
        try:
            raiz_abs = raiz.resolve()
        except Exception:
            raiz_abs = raiz
        chave = str(raiz_abs).upper()
        if chave in vistos:
            continue
        vistos.add(chave)
        if raiz_abs.exists():
            candidatos_raiz.append(raiz_abs)

    for raiz in candidatos_raiz:
        # 1) Caminho dentro da pasta SANPORT (suporta variações de layout)
        try:
            for item in raiz.iterdir():
                if not item.is_dir():
                    continue

                item_nome = _normalizar_texto(item.name)
                if "SANPORT" in item_nome:
                    # 1.1) Direto dentro da pasta SANPORT
                    for sub in item.iterdir():
                        if sub.is_dir() and _eh_base_faturamentos(sub):
                            return sub

                    # 1.2) Padrão: SANPORT\Central de Documentos - Documentos\01. FATURAMENTOS
                    central_docs = item / "Central de Documentos - Documentos"
                    if central_docs.exists() and central_docs.is_dir():
                        for sub in central_docs.iterdir():
                            if sub.is_dir() and _eh_base_faturamentos(sub):
                                return sub
        except Exception:
            pass

        # 2) Caminhos alternativos diretos
        candidatos_diretos = [
            raiz / "Central de Documentos - 01. FATURAMENTOS",
            raiz / "Central de Documentos - Documentos" / "01. FATURAMENTOS",
        ]
        for direto in candidatos_diretos:
            if direto.exists() and direto.is_dir() and _eh_base_faturamentos(direto):
                return direto

    # 3) Busca recursiva limitada no perfil do usuário
    for raiz in candidatos_raiz:
        try:
            for candidato in raiz.rglob("*"):
                if candidato.is_dir() and _eh_base_faturamentos(candidato):
                    return candidato
        except Exception:
            continue

    return None


def obter_caminho_base_faturamentos() -> Path:
    """
    Obtém o caminho base da pasta de faturamentos.
    
    Ordem de prioridade:
    1. Caminho configurado em config.json
    2. Auto-detecção (se habilitada)
    3. Erro se nada funcionar
    """
    config = _carregar_config()
    
    # 1. Tenta usar caminho configurado
    caminho_config = config.get("caminho_base_faturamentos", "").strip()
    if caminho_config:
        caminho = Path(caminho_config)
        if caminho.exists():
            return caminho
        else:
            print(f"⚠️ Caminho configurado não existe: {caminho}")
    
    # 2. Tenta auto-detecção
    if config.get("auto_detectar", True):
        caminho_auto = _auto_detectar_base_faturamentos()
        if caminho_auto:
            print(f"✅ Caminho auto-detectado: {caminho_auto}")
            return caminho_auto
    
    # 3. Erro
    raise FileNotFoundError(
        "❌ Pasta de faturamentos não encontrada!\n"
        "\n"
        "Soluções:\n"
        "1. Configure o caminho manualmente no menu 'Configurações' do sistema\n"
        "2. Ou edite o arquivo config.json na raiz do projeto\n"
        "\n"
        f"O caminho deve ser algo como:\n"
        f"C:\\Users\\SeuNome\\SANPORT LOGÍSTICA PORTUÁRIA LTDA\\Central de Documentos - 01. FATURAMENTOS\n"
        f"ou\n"
        f"C:\\Users\\SeuNome\\SANPORT LOGÍSTICA PORTUÁRIA LTDA\\Central de Documentos - Documentos\\01. FATURAMENTOS"
    )


def configurar_caminho_base(caminho: str) -> bool:
    """
    Permite configurar manualmente o caminho base.
    
    Args:
        caminho: Caminho completo para a pasta base de faturamentos
    
    Returns:
        True se a configuração foi salva com sucesso
    """
    caminho_path = Path(caminho)
    
    if not caminho_path.exists():
        print(f"❌ O caminho não existe: {caminho}")
        return False
    
    if not caminho_path.is_dir():
        print(f"❌ O caminho não é uma pasta: {caminho}")
        return False
    
    config = _carregar_config()
    config["caminho_base_faturamentos"] = str(caminho_path)
    _salvar_config(config)
    
    return True


def obter_caminho_configurado() -> str:
    """Retorna o caminho atualmente configurado (vazio se não configurado)"""
    config = _carregar_config()
    return config.get("caminho_base_faturamentos", "")


def listar_caminhos_detectados() -> list[Path]:
    """Lista caminhos detectados no sistema."""
    caminho = _auto_detectar_base_faturamentos()
    return [caminho] if caminho else []


def obter_caminho_assinatura_usuario(usuario_nome: str) -> Path | None:
    """
    Retorna o caminho da imagem de assinatura para o usuário, se configurado.

    Config esperado no config.json:
    {
      "assinaturas_usuarios": {
        "CAROL CARMO": "C:/caminho/assinatura_carol.png",
        "DIOGO BARROS": "C:/caminho/assinatura_diogo.png"
      }
    }
    """

    def _normalizar_usuario(valor: str) -> str:
        texto = unicodedata.normalize("NFKD", str(valor or ""))
        texto = "".join(ch for ch in texto if not unicodedata.combining(ch))
        return " ".join(texto.upper().split())

    usuario_norm = _normalizar_usuario(usuario_nome)
    if not usuario_norm:
        return None

    config = _carregar_config()
    mapa = config.get("assinaturas_usuarios", {})

    if isinstance(mapa, dict):
        caminho = str(mapa.get(usuario_norm, "")).strip()
        if caminho:
            path = Path(caminho)
            if path.exists() and path.is_file():
                return path

    return None
