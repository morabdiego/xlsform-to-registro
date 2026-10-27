from typing import Any, Dict, List, Optional
import openpyxl

from .core.parser import XLSFormParser
from .core.schema import CatalogueSheet, EntitySheet
from .exporters.excel_styler import ExcelStyler


class XLSFormToRegistro:
    """
    Clase principal para transformar un formulario XLSForm en un
    Diseño de Registro institucional en Excel.
    """

    def __init__(
        self,
        xlsform_path: str,
        lang: Optional[str] = None,
        header_color: str = "333399",
        include_calculate: bool = False,
        include_notes: bool = False,
        catalogue_threshold: int = 30,
        main_sheet_name: str = "base_hogar",
        add_not_applicable_code: bool = True,
        custom_replacements: Optional[Dict[str, str]] = None,
    ):
        self.xlsform_path = xlsform_path
        self.lang = lang
        self.header_color = header_color
        self.include_calculate = include_calculate
        self.include_notes = include_notes
        self.catalogue_threshold = catalogue_threshold
        self.main_sheet_name = main_sheet_name
        self.add_not_applicable_code = add_not_applicable_code
        self.custom_replacements = custom_replacements or {}

        self.parser = XLSFormParser(
            filepath=self.xlsform_path,
            lang=self.lang,
            include_calculate=self.include_calculate,
            include_notes=self.include_notes,
            catalogue_threshold=self.catalogue_threshold,
            main_sheet_name=self.main_sheet_name,
            custom_replacements=self.custom_replacements,
        )

        self.styler = ExcelStyler(
            header_color=self.header_color,
            add_not_applicable_code=self.add_not_applicable_code,
        )

        self.entity_sheets: List[EntitySheet] = []
        self.catalogue_sheets: List[CatalogueSheet] = []
        self._parsed = False

    def parse(self) -> None:
        """Parsea el XLSForm y almacena la estructura lógica."""
        self.entity_sheets, self.catalogue_sheets = self.parser.parse()
        self._parsed = True

    def build_workbook(self) -> openpyxl.Workbook:
        """Construye el Workbook de openpyxl."""
        if not self._parsed:
            self.parse()
        return self.styler.create_workbook(self.entity_sheets, self.catalogue_sheets)

    def to_excel(self, output_path: str) -> str:
        """Genera y guarda el archivo Excel en la ruta especificada."""
        import os
        dir_name = os.path.dirname(os.path.abspath(output_path))
        if dir_name:
            os.makedirs(dir_name, exist_ok=True)
        wb = self.build_workbook()
        wb.save(output_path)
        return output_path
