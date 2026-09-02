"""
config_store.py
Punto unico de lectura de los ajustes de IA. La clave y el modelo se
definen en ai_config.py (editalo directamente en el codigo para activar
la IA); opcionalmente tambien se puede definir la variable de entorno
ANTHROPIC_API_KEY, que tiene prioridad sobre ai_config.py (util para un
despliegue centralizado sin tocar el codigo).
"""
import os

import ai_config

DEFAULT_MODEL = "claude-sonnet-5"


def get_api_key():
    env_key = os.environ.get("ANTHROPIC_API_KEY")
    if env_key and env_key.strip():
        return env_key.strip()
    return (getattr(ai_config, "ANTHROPIC_API_KEY", "") or "").strip()


def get_model():
    return (getattr(ai_config, "AI_MODEL", "") or "").strip() or DEFAULT_MODEL


def is_ai_enabled():
    use_ai = bool(getattr(ai_config, "USE_AI", True))
    return use_ai and bool(get_api_key())
