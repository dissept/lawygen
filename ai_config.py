"""
ai_config.py
========================================================================
CONFIGURACIÓN DE LA IA (Claude / Anthropic) — EDITA SOLO ESTE ARCHIVO
========================================================================
Para activar el análisis asistido por IA (mejora partes/representantes,
materia, objeto, plazos, cuantía/costas, argumentos y el resumen
ejecutivo), pega tu clave de la API de Anthropic entre las comillas de
ANTHROPIC_API_KEY más abajo.

Si la dejas vacía (""), LawyGen funciona igualmente al 100%, solo que
usando exclusivamente las reglas de texto de legal_parser.py (gratis y
offline, sin conexión a internet).

Consigue una clave en: https://console.anthropic.com/
========================================================================
"""

# Pega tu clave aquí, entre las comillas. Ejemplo: "sk-ant-api03-xxxxxxxx"
ANTHROPIC_API_KEY = ""

# Modelo a usar para el análisis. Configurado en Sonnet 5 para la máxima
# calidad posible (el objetivo es no tener que revisar nada a mano). Si en
# algún momento quieres priorizar el coste sobre la precisión para tandas
# muy grandes, cambia esto a "claude-haiku-4-5-20251001" (aprox. la mitad
# de precio, algo menos fino en textos narrativos/ambiguos).
AI_MODEL = "claude-sonnet-5"

# Pon esto en False para desactivar la IA aunque haya una clave puesta
# arriba (por ejemplo, para volver temporalmente al modo 100% gratuito
# sin borrar la clave).
USE_AI = True
