"""
Gerenciador de configuração do Yuta
Permite configurar caminhos específicos por computador
"""

import json
from pathlib import Path
from typing import Dict, Any


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
    """Tenta detectar automaticamente a pasta base de faturamentos"""
    home = Path.home()
    
    possiveis_bases = [
        # OneDrive empresarial
        home / "OneDrive - SANPORT LOGÍSTICA PORTUÁRIA LTDA" / "Central de Documentos - 01. FATURAMENTOS",
        # Pasta sincronizada diretamente
        home / "SANPORT LOGÍSTICA PORTUÁRIA LTDA" / "Central de Documentos - 01. FATURAMENTOS",
    ]
    
    for base in possiveis_bases:
        if base.exists():
            return base
    
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
        f"C:\\Users\\SeuNome\\SANPORT LOGÍSTICA PORTUÁRIA LTDA\\Central de Documentos - 01. FATURAMENTOS"
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
    """Lista todos os caminhos possíveis detectados no sistema"""
    home = Path.home()
    
    possiveis = [
        home / "OneDrive - SANPORT LOGÍSTICA PORTUÁRIA LTDA" / "Central de Documentos - 01. FATURAMENTOS",
        home / "SANPORT LOGÍSTICA PORTUÁRIA LTDA" / "Central de Documentos - 01. FATURAMENTOS",
    ]
    
    return [p for p in possiveis if p.exists()]
