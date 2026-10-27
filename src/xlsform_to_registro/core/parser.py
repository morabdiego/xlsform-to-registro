import re
from typing import Any, Dict, List, Optional, Tuple
import pandas as pd

from .relevant_parser import parse_relevant_to_universe
from .schema import CatalogueSheet, Choice, EntitySheet, Section, Variable
from .text_cleaner import clean_label


class XLSFormParser:
    """Parsea un XLSForm completo y genera la estructura lógica del diseño de registro."""

    def __init__(
        self,
        filepath: str,
        lang: Optional[str] = None,
        include_calculate: bool = False,
        include_notes: bool = False,
        catalogue_threshold: int = 30,
        main_sheet_name: str = "base_principal",
        custom_replacements: Optional[Dict[str, str]] = None,
    ):
        self.filepath = filepath
        self.lang = lang
        self.include_calculate = include_calculate
        self.include_notes = include_notes
        self.catalogue_threshold = catalogue_threshold
        self.main_sheet_name = main_sheet_name
        self.custom_replacements = custom_replacements or {}

        self.survey_df: pd.DataFrame = pd.DataFrame()
        self.choices_df: pd.DataFrame = pd.DataFrame()
        self.settings_df: pd.DataFrame = pd.DataFrame()

        self.choices_by_list: Dict[str, List[Choice]] = {}
        self.label_col_survey: str = "label"
        self.label_col_choices: str = "label"

    def load(self) -> None:
        """Carga las hojas del archivo Excel."""
        xl = pd.ExcelFile(self.filepath)
        sheet_names = xl.sheet_names

        if "survey" not in sheet_names:
            raise ValueError(f"El archivo {self.filepath} no contiene la hoja obligatoria 'survey'.")

        self.survey_df = pd.read_excel(xl, sheet_name="survey")
        if "choices" in sheet_names:
            self.choices_df = pd.read_excel(xl, sheet_name="choices")
        if "settings" in sheet_names:
            self.settings_df = pd.read_excel(xl, sheet_name="settings")

        self._detect_languages()
        self._load_choices()

    def _detect_languages(self) -> None:
        """Detecta la columna de etiqueta adecuada según el idioma especificado o disponible."""
        # Survey label
        survey_cols = list(self.survey_df.columns)
        choices_cols = list(self.choices_df.columns) if not self.choices_df.empty else []

        label_survey_candidates = [c for c in survey_cols if c.startswith("label")]
        if not label_survey_candidates:
            self.label_col_survey = "label"
        elif self.lang:
            matched = [c for c in label_survey_candidates if self.lang.lower() in c.lower()]
            self.label_col_survey = matched[0] if matched else label_survey_candidates[0]
        else:
            if "label" in survey_cols:
                self.label_col_survey = "label"
            else:
                self.label_col_survey = label_survey_candidates[0]

        # Choices label
        label_choices_candidates = [c for c in choices_cols if c.startswith("label")]
        if not label_choices_candidates:
            self.label_col_choices = "label"
        elif self.lang:
            matched = [c for c in label_choices_candidates if self.lang.lower() in c.lower()]
            self.label_col_choices = matched[0] if matched else label_choices_candidates[0]
        else:
            if "label" in choices_cols:
                self.label_col_choices = "label"
            else:
                self.label_col_choices = label_choices_candidates[0]

    def _load_choices(self) -> None:
        """Agrupa las opciones de respuesta por list_name."""
        if self.choices_df.empty:
            return

        for list_name, group in self.choices_df.groupby("list_name", sort=False):
            if pd.isna(list_name):
                continue
            lname = str(list_name).strip()
            self.choices_by_list[lname] = []

            for _, row in group.iterrows():
                val_raw = row.get("name")
                if pd.isna(val_raw):
                    continue
                # Normalizar valor
                val = val_raw
                if isinstance(val_raw, float) and val_raw.is_integer():
                    val = int(val_raw)

                lbl_raw = row.get(self.label_col_choices, row.get("label", ""))
                lbl = str(lbl_raw).strip() if pd.notna(lbl_raw) else ""

                extra = {k: v for k, v in row.items() if k not in ["list_name", "name", self.label_col_choices, "label"] and pd.notna(v)}
                self.choices_by_list[lname].append(Choice(name=val, label=lbl, extra_cols=extra))

    def parse(self) -> Tuple[List[EntitySheet], List[CatalogueSheet]]:
        """Procesa el flujo de preguntas y genera las hojas de entidades y catálogos."""
        self.load()

        # Determinar nombre de la entidad principal
        main_name = self.main_sheet_name
        if not self.settings_df.empty:
            if "form_id" in self.settings_df.columns and pd.notna(self.settings_df["form_id"].iloc[0]):
                form_id = str(self.settings_df["form_id"].iloc[0]).strip()
                if form_id:
                    main_name = form_id
            elif "form_title" in self.settings_df.columns and pd.notna(self.settings_df["form_title"].iloc[0]):
                form_title = str(self.settings_df["form_title"].iloc[0]).strip()
                if form_title:
                    main_name = form_title

        main_sheet = EntitySheet(name=main_name, title="Encuesta Principal", is_repeat=False)
        entity_sheets: List[EntitySheet] = [main_sheet]
        catalogue_sheets: List[CatalogueSheet] = []

        # Pilas de contexto
        # Entity stack: lista de EntitySheet
        entity_stack: List[EntitySheet] = [main_sheet]
        # Section stack: (group_name, group_title, group_relevant)
        group_stack: List[Tuple[str, str, Optional[str]]] = []

        # Asegurar sección inicial por defecto
        current_section = Section(name="general", title="Identificación y Datos Generales")
        main_sheet.sections.append(current_section)

        # Registro de listas usadas para catálogos
        used_choice_lists: Dict[str, str] = {}

        for _, row in self.survey_df.iterrows():
            raw_type = str(row.get("type", "")).strip()
            if not raw_type or pd.isna(row.get("type")):
                continue

            qname = str(row.get("name", "")).strip() if pd.notna(row.get("name")) else ""
            qlbl_raw = row.get(self.label_col_survey, row.get("label", ""))
            qlbl = clean_label(qlbl_raw, self.custom_replacements) if pd.notna(qlbl_raw) else ""
            qrel = str(row.get("relevant")).strip() if pd.notna(row.get("relevant")) else None
            qcalc = str(row.get("calculation")).strip() if pd.notna(row.get("calculation")) else None
            qhint = str(row.get("hint")).strip() if pd.notna(row.get("hint")) else None

            # 1. Manejo de repeats
            if raw_type.startswith("begin_repeat") or raw_type == "begin repeat":
                repeat_name = qname or f"repeat_{len(entity_sheets)}"
                repeat_title = qlbl or repeat_name.replace("_", " ").title()
                new_sheet = EntitySheet(name=repeat_name, title=repeat_title, is_repeat=True)
                # Crear sección inicial del repeat
                rep_sec = Section(name="datos_repeat", title=repeat_title)
                new_sheet.sections.append(rep_sec)
                entity_sheets.append(new_sheet)
                entity_stack.append(new_sheet)
                continue

            elif raw_type.startswith("end_repeat") or raw_type == "end repeat":
                if len(entity_stack) > 1:
                    entity_stack.pop()
                continue

            # 2. Manejo de groups
            if raw_type.startswith("begin_group") or raw_type == "begin group":
                gname = qname or f"group_{len(group_stack)}"
                gtitle = qlbl or gname.replace("_", " ").title()
                group_stack.append((gname, gtitle, qrel))

                # Crear nueva sección visible si tiene título
                if qlbl:
                    new_sec = Section(name=gname, title=gtitle)
                    entity_stack[-1].sections.append(new_sec)
                continue

            elif raw_type.startswith("end_group") or raw_type == "end group":
                if group_stack:
                    group_stack.pop()
                continue

            # 3. Filtrar preguntas que no son variables de datos
            if not qname:
                continue

            if raw_type == "note" and not self.include_notes:
                continue
            if raw_type == "calculate" and not self.include_calculate:
                continue

            # Parsear tipo de pregunta
            type_parts = raw_type.split()
            base_type = type_parts[0]
            choice_list = type_parts[1] if len(type_parts) > 1 else None
            has_other = "or_other" in type_parts

            # Obtener relevant acumulado de los grupos padres
            parent_rel_parts = [g[2] for g in group_stack if g[2]]
            parent_relevant = " and ".join(parent_rel_parts) if parent_rel_parts else None

            # Determinar universo
            default_univ = "Población total" if entity_stack[-1].is_repeat else "Total de registros"
            universe = parse_relevant_to_universe(qrel, parent_relevant, default_univ)

            # Obtener choices si corresponde
            choices_list: List[Choice] = []
            is_ext = False
            if choice_list:
                if base_type in ["select_one_from_file", "select_multiple_from_file"]:
                    is_ext = True
                elif choice_list in self.choices_by_list:
                    choices_list = self.choices_by_list[choice_list].copy()
                    used_choice_lists[choice_list] = qlbl or choice_list
                else:
                    is_ext = True

            var = Variable(
                name=qname,
                qtype=raw_type,
                label=qlbl if qlbl else f"[{qname}]",
                relevant=qrel,
                universe=universe,
                choice_list=choice_list,
                choices=choices_list,
                has_other=has_other,
                is_external=is_ext,
                calculation=qcalc,
                hint=qhint,
            )

            # Agregar variable a la sección activa de la entidad actual
            curr_entity = entity_stack[-1]
            if not curr_entity.sections:
                curr_entity.sections.append(Section(name="general", title="Datos Generales"))
            curr_entity.sections[-1].variables.append(var)

        # 4. Limpiar entidades vacías y secciones vacías
        final_entities = []
        for es in entity_sheets:
            non_empty_sections = [sec for sec in es.sections if sec.variables]
            if non_empty_sections:
                es.sections = non_empty_sections
                final_entities.append(es)

        # 5. Generar hojas de catálogo complementarias para listas con más de `catalogue_threshold` ítems
        for lname, ltitle in used_choice_lists.items():
            ch_items = self.choices_by_list.get(lname, [])
            if len(ch_items) >= self.catalogue_threshold:
                # Detectar encabezados extras
                extra_keys = []
                for ch in ch_items:
                    for k in ch.extra_cols.keys():
                        if k not in extra_keys:
                            extra_keys.append(k)

                cat_headers = ["Código", "Descripción"] + [k.replace("_", " ").title() for k in extra_keys]
                cat_rows = []
                for ch in ch_items:
                    row_vals = [ch.name, ch.label] + [ch.extra_cols.get(k, "") for k in extra_keys]
                    cat_rows.append(row_vals)

                sheet_name = f"catalogo_{lname}"[:31]  # Limite Excel 31 chars
                catalogue_sheets.append(
                    CatalogueSheet(
                        sheet_name=sheet_name,
                        list_name=lname,
                        title=f"Catálogo de {lname} ({ltitle})",
                        headers=cat_headers,
                        rows=cat_rows,
                    )
                )

        return final_entities, catalogue_sheets
