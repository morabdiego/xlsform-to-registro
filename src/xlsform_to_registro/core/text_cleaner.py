import re
from typing import Dict, Optional


# Reemplazos comunes de variables de texto en preguntas tipo ODK
DEFAULT_VAR_REPLACEMENTS: Dict[str, str] = {
    "nombre": "la persona entrevistada",
    "nombre_integrante": "el/la integrante",
    "integrante": "el/la integrante",
    "miembro": "el/la integrante",
    "persona": "la persona",
    "jefe": "el/la jefe/a de hogar",
}


def clean_html(text: str) -> str:
    """Elimina etiquetas HTML comunes."""
    if not text:
        return ""
    # Reemplazar <br> o <br/> por espacio
    text = re.sub(r"<br\s*/?>", " ", text, flags=re.IGNORECASE)
    # Remover otras etiquetas HTML como <b>, </i>, <span>, etc.
    text = re.sub(r"<[^>]+>", "", text)
    return text


def clean_text_placeholders(text: str, custom_replacements: Optional[Dict[str, str]] = None) -> str:
    """
    Reemplaza placeholders tipo ${nombre} por texto legible.
    """
    if not text:
        return ""
    
    replacements = DEFAULT_VAR_REPLACEMENTS.copy()
    if custom_replacements:
        replacements.update(custom_replacements)

    def repl(match: re.Match) -> str:
        var_name = match.group(1).strip().lower()
        if var_name in replacements:
            return replacements[var_name]
        return f"[{var_name}]"

    # Buscar ${var_name}
    cleaned = re.sub(r"\$\{([a-zA-Z0-9_]+)\}", repl, text)
    return cleaned


def clean_label(text: Optional[str], custom_replacements: Optional[Dict[str, str]] = None) -> str:
    """Aplica la secuencia completa de limpieza a una etiqueta."""
    if text is None:
        return ""
    s = str(text)
    s = clean_html(s)
    s = clean_text_placeholders(s, custom_replacements)
    # Normalizar espacios múltiples
    s = re.sub(r"[ \t]+", " ", s)
    return s.strip()
