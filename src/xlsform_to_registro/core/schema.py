from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional


@dataclass
class Choice:
    """Representa una opción de respuesta proveniente de la hoja choices."""
    name: Any
    label: str
    extra_cols: Dict[str, Any] = field(default_factory=dict)


@dataclass
class Variable:
    """Representa una variable (pregunta) del formulario."""
    name: str
    qtype: str
    label: str
    relevant: Optional[str] = None
    universe: str = ""
    choice_list: Optional[str] = None
    choices: List[Choice] = field(default_factory=list)
    has_other: bool = False
    is_external: bool = False
    calculation: Optional[str] = None
    hint: Optional[str] = None


@dataclass
class Section:
    """Representa una sección temática derivada de begin_group."""
    name: str
    title: str
    variables: List[Variable] = field(default_factory=list)


@dataclass
class EntitySheet:
    """Representa una tabla/entidad (hoja principal o repeat)."""
    name: str
    title: str
    is_repeat: bool = False
    sections: List[Section] = field(default_factory=list)


@dataclass
class CatalogueSheet:
    """Representa una hoja complementaria para listas de opciones extensas."""
    sheet_name: str
    list_name: str
    title: str
    headers: List[str]
    rows: List[List[Any]]
