import re
from typing import Optional


def parse_relevant_to_universe(
    relevant_expr: Optional[str],
    parent_relevant: Optional[str] = None,
    default_universe: str = "Población total"
) -> str:
    """
    Traduce una expresión relevant de ODK / XLSForm a una descripción
    de universo de referencia comprensible en el diseño de registro.
    """
    combined = []
    if parent_relevant and parent_relevant.strip():
        combined.append(f"({parent_relevant.strip()})")
    if relevant_expr and relevant_expr.strip():
        combined.append(f"({relevant_expr.strip()})")

    if not combined:
        return default_universe

    expr = " and ".join(combined) if len(combined) > 1 else combined[0]

    # Quitar paréntesis redundantes envolventes
    if expr.startswith("(") and expr.endswith(")"):
        # Solo si los paréntesis envuelven todo
        # Un check rápido:
        depth = 0
        all_enclosed = True
        for i, char in enumerate(expr):
            if char == "(":
                depth += 1
            elif char == ")":
                depth -= 1
                if depth == 0 and i < len(expr) - 1:
                    all_enclosed = False
                    break
        if all_enclosed:
            expr = expr[1:-1].strip()

    # Reemplazar ${variable} por variable
    expr = re.sub(r"\$\{([a-zA-Z0-9_]+)\}", r"\1", expr)

    # Reemplazar selected(var, 'val') por var = 'val' (o contiene)
    expr = re.sub(
        r"selected\s*\(\s*([a-zA-Z0-9_]+)\s*,\s*['\"]([^'\"]+)['\"]\s*\)",
        r"\1 incluye '\2'",
        expr
    )

    # Reemplazar operadores lógicos
    expr = re.sub(r"\band\b", "y", expr, flags=re.IGNORECASE)
    expr = re.sub(r"\bor\b", "o", expr, flags=re.IGNORECASE)
    expr = re.sub(r"!=", "≠", expr)
    expr = re.sub(r"\s*=\s*", " = ", expr)

    # Limpiar espacios repetidos
    expr = re.sub(r"\s+", " ", expr).strip()

    return expr
