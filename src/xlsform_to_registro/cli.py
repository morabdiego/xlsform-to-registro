import argparse
import os
import sys

from .generator import XLSFormToRegistro


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Generador automático de Diseño de Registro (Data Dictionary) en Excel a partir de un XLSForm."
    )
    parser.add_argument(
        "xlsform",
        help="Ruta al archivo XLSForm (.xlsx o .xls)."
    )
    parser.add_argument(
        "-o", "--output",
        help="Ruta al archivo Excel de salida. Por defecto: diseño_registro_<nombre>.xlsx",
        default=None
    )
    parser.add_argument(
        "--lang",
        help="Idioma o sufijo de etiqueta a utilizar (ej. 'es', 'Español', 'en').",
        default=None
    )
    parser.add_argument(
        "--header-color",
        help="Color hexadecimal de los encabezados (sin '#'). Por defecto: 64b692",
        default="64b692"
    )
    parser.add_argument(
        "--include-calculate",
        help="Incluir variables de tipo 'calculate' en el diseño de registro.",
        action="store_true"
    )
    parser.add_argument(
        "--include-notes",
        help="Incluir variables de tipo 'note' en el diseño de registro.",
        action="store_true"
    )
    parser.add_argument(
        "--catalogue-threshold",
        type=int,
        help="Umbral de opciones para separar una lista en hoja de catálogo independiente. Por defecto: 30",
        default=30
    )
    parser.add_argument(
        "--main-sheet-name",
        help="Nombre de la hoja principal (fuera de repeats). Por defecto: base_hogar",
        default="base_hogar"
    )
    parser.add_argument(
        "--no-na-code",
        help="No agregar automáticamente la opción '0: No corresponde / No aplica' en variables con salto de pregunta.",
        action="store_true"
    )

    args = parser.parse_args()

    if not os.path.exists(args.xlsform):
        print(f"Error: No se encontró el archivo de entrada '{args.xlsform}'.", file=sys.stderr)
        sys.exit(1)

    output_path = args.output
    if not output_path:
        base_name = os.path.splitext(os.path.basename(args.xlsform))[0]
        output_path = f"diseño_registro_{base_name}.xlsx"

    print(f"-> Procesando XLSForm: {args.xlsform}")
    try:
        converter = XLSFormToRegistro(
            xlsform_path=args.xlsform,
            lang=args.lang,
            header_color=args.header_color,
            include_calculate=args.include_calculate,
            include_notes=args.include_notes,
            catalogue_threshold=args.catalogue_threshold,
            main_sheet_name=args.main_sheet_name,
            add_not_applicable_code=not args.no_na_code,
        )
        converter.to_excel(output_path)
        print(f"-> Diseño de registro generado con éxito en: {output_path}")

        print("\nHojas generadas:")
        for es in converter.entity_sheets:
            total_vars = sum(len(sec.variables) for sec in es.sections)
            print(f"  • [Entidad] {es.name}: {len(es.sections)} secciones, {total_vars} variables")
        for cat in converter.catalogue_sheets:
            print(f"  • [Catálogo] {cat.sheet_name}: {len(cat.rows)} opciones")

    except Exception as e:
        print(f"Error durante la generación: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
