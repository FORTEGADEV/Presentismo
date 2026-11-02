"""Calcula quién cobra presentismo a partir del Excel de control horario.

Estructura del Excel que entiende (una hoja por local, un bloque por persona):

* La fila 4 de cada hoja tiene los encabezados; cada bloque empieza en una
  columna "Nombre" y termina antes de la siguiente.
* La fila 3 del bloque trae el nombre de la persona.
* La columna "PRESENTISMO" del bloque lleva, día por día, los minutos de
  llegada tarde (se carga a mano).
* Las faltas aparecen como texto libre ("FALTA LUNES 20", "falta con aviso",
  etc.) en cualquier celda del bloque.
* Al pie del bloque puede haber un resumen manual (PRESENTISMO / FALTAS /
  LICENCIA con SI/NO al lado). Se lee solo para compararlo con el cálculo.

Uso:
    python presentismo.py Control_Horario_JULIO_2026.xlsx
    python presentismo.py archivo.xlsx --max-late-minutes 15 --excusar 'aviso|certific' --csv salida.csv
"""

from __future__ import annotations

import argparse
import csv
import re
import sys
import unicodedata
from dataclasses import dataclass, field
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet

HEADER_ROW = 4
NAME_ROW = 3
FIRST_DATA_ROW = 5
IGNORED_SHEETS = {"informe"}

FALTA_RE = re.compile(r"\bfalta")


def _norm(value) -> str:
    """Minúsculas, sin tildes y sin espacios sobrantes."""
    text = unicodedata.normalize("NFKD", str(value or ""))
    text = "".join(c for c in text if not unicodedata.combining(c))
    return text.strip().lower()


@dataclass
class Persona:
    local: str
    nombre: str
    minutos_tarde: float = 0
    dias_tarde: int = 0
    marca_manual: str | None = None  # "SI" / "NO" / None si no está cargado
    notas_faltas: list[str] = field(default_factory=list)  # "AL26: Falta / Muela"

    @property
    def faltas(self) -> int:
        return len(self.notas_faltas)

    def faltas_que_cuentan(self, excusar: re.Pattern | None = None) -> int:
        """Faltas cuyo texto NO coincide con el patrón de faltas excusadas."""
        if excusar is None:
            return self.faltas
        return sum(1 for nota in self.notas_faltas if not excusar.search(_norm(nota)))

    def cobra_presentismo(
        self,
        max_late_minutes: float = 0,
        max_faltas: int = 0,
        excusar: re.Pattern | None = None,
    ) -> bool:
        return (
            self.minutos_tarde <= max_late_minutes
            and self.faltas_que_cuentan(excusar) <= max_faltas
        )


def _find_blocks(ws: Worksheet) -> list[tuple[int, int]]:
    """Devuelve (col_inicio, col_fin) de cada bloque, según la fila de encabezados."""
    starts = [
        c
        for c in range(1, ws.max_column + 1)
        if _norm(ws.cell(HEADER_ROW, c).value) == "nombre"
    ]
    return [
        (start, (starts[i + 1] - 1) if i + 1 < len(starts) else ws.max_column)
        for i, start in enumerate(starts)
    ]


def _header_columns(ws: Worksheet, start: int, end: int) -> dict[str, int]:
    return {
        _norm(ws.cell(HEADER_ROW, c).value): c
        for c in range(start, end + 1)
        if ws.cell(HEADER_ROW, c).value
    }


def _summary_row(ws: Worksheet, start: int) -> int | None:
    """Fila donde empieza el resumen manual del bloque (label 'PRESENTISMO')."""
    for r in range(FIRST_DATA_ROW, ws.max_row + 1):
        if _norm(ws.cell(r, start).value) == "presentismo":
            return r
    return None


def _manual_mark(ws: Worksheet, start: int, summary_row: int | None) -> str | None:
    if summary_row is None:
        return None
    value = _norm(ws.cell(summary_row, start + 1).value)
    return value.upper() if value in ("si", "no") else None


def _person_name(ws: Worksheet, start: int, end: int, last_row: int) -> str | None:
    name = ws.cell(NAME_ROW, start).value
    if name and str(name).strip():
        return str(name).strip()
    for r in range(FIRST_DATA_ROW, last_row + 1):
        value = ws.cell(r, start).value
        if value and str(value).strip():
            return str(value).strip()
    return None


def parse_sheet(ws: Worksheet) -> list[Persona]:
    personas: list[Persona] = []
    for start, end in _find_blocks(ws):
        columns = _header_columns(ws, start, end)
        late_col = columns.get("presentismo")
        summary = _summary_row(ws, start)
        last_row = (summary - 1) if summary else ws.max_row
        name = _person_name(ws, start, end, last_row)
        if not name:
            continue  # bloque vacío

        persona = Persona(
            local=ws.title.strip(),
            nombre=name,
            marca_manual=_manual_mark(ws, start, summary),
        )

        for r in range(FIRST_DATA_ROW, last_row + 1):
            if late_col:
                value = ws.cell(r, late_col).value
                if isinstance(value, (int, float)) and not isinstance(value, bool):
                    persona.minutos_tarde += value
                    if value > 0:
                        persona.dias_tarde += 1
            for c in range(start, end + 1):
                value = ws.cell(r, c).value
                if not isinstance(value, str) or value.startswith("="):
                    continue
                text = _norm(value)
                if FALTA_RE.search(text):
                    persona.notas_faltas.append(
                        f"{ws.cell(r, c).coordinate}: {value.strip()}"
                    )
        personas.append(persona)
    return personas


def merge_personas(personas: list[Persona]) -> list[Persona]:
    """Une a quien aparece en más de una hoja (ej.: cubre otro local)."""
    merged: dict[str, Persona] = {}
    for p in personas:
        key = " ".join(_norm(p.nombre).split())
        if key not in merged:
            merged[key] = p
            continue
        base = merged[key]
        if p.local not in base.local.split("/"):
            base.local += "/" + p.local
        base.minutos_tarde += p.minutos_tarde
        base.dias_tarde += p.dias_tarde
        base.notas_faltas += p.notas_faltas
        base.marca_manual = base.marca_manual or p.marca_manual
    return list(merged.values())


def parse_workbook(path: str | Path) -> list[Persona]:
    wb = load_workbook(path, data_only=False)
    personas: list[Persona] = []
    for ws in wb.worksheets:
        if _norm(ws.title) in IGNORED_SHEETS:
            continue
        personas.extend(parse_sheet(ws))
    return merge_personas(personas)


def _print_report(personas, rule) -> None:
    cobran = [p for p in personas if p.cobra_presentismo(*rule)]
    no_cobran = [p for p in personas if not p.cobra_presentismo(*rule)]

    def line(p: Persona) -> str:
        manual = f"  [marca manual: {p.marca_manual}]" if p.marca_manual else ""
        return (
            f"  {p.local:<18} {p.nombre:<24} "
            f"tarde: {p.minutos_tarde:g} min ({p.dias_tarde} d)  "
            f"faltas: {p.faltas_que_cuentan(rule[2])}/{p.faltas}{manual}"
        )

    print(f"CON PRESENTISMO ({len(cobran)})")
    for p in cobran:
        print(line(p))
    print(f"\nSIN PRESENTISMO ({len(no_cobran)})")
    for p in no_cobran:
        print(line(p))

    discrepancias = [
        p
        for p in personas
        if p.marca_manual and (p.marca_manual == "SI") != p.cobra_presentismo(*rule)
    ]
    if discrepancias:
        print("\nREVISAR: el cálculo no coincide con la marca manual del Excel")
        for p in discrepancias:
            print(line(p))


def _write_csv(path, personas, rule) -> None:
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow(
            ["local", "nombre", "minutos_tarde", "dias_tarde", "faltas",
             "faltas_que_cuentan", "marca_manual", "presentismo"]
        )
        for p in personas:
            writer.writerow(
                [p.local, p.nombre, p.minutos_tarde, p.dias_tarde, p.faltas,
                 p.faltas_que_cuentan(rule[2]), p.marca_manual or "",
                 "SI" if p.cobra_presentismo(*rule) else "NO"]
            )


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    parser.add_argument("archivo", help="Excel de control horario (.xlsx)")
    parser.add_argument("--max-late-minutes", type=float, default=0,
                        help="minutos tarde tolerados en el mes (default: 0)")
    parser.add_argument("--max-faltas", type=int, default=0,
                        help="faltas toleradas en el mes (default: 0)")
    parser.add_argument("--excusar", metavar="REGEX",
                        help="no contar las faltas cuyo texto coincida "
                             "(ej.: 'aviso|certific|madre')")
    parser.add_argument("--csv", help="guardar el resultado en un CSV")
    args = parser.parse_args(argv)

    personas = parse_workbook(args.archivo)
    if not personas:
        print("No se encontraron personas en el archivo.", file=sys.stderr)
        return 1

    excusar = re.compile(_norm(args.excusar)) if args.excusar else None
    rule = (args.max_late_minutes, args.max_faltas, excusar)
    _print_report(personas, rule)
    if args.csv:
        _write_csv(args.csv, personas, rule)
        print(f"\nCSV guardado en {args.csv}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
