try:
    from .services.config_manager import *
except ModuleNotFoundError as exc:
    if exc.name in {
        "backend",
        "backend.app",
        "backend.app.services",
        "backend.app.services.config_manager",
    }:
        from services.config_manager import *
    else:
        raise
