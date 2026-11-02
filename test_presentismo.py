import re

import pytest
from openpyxl import Workbook

from presentismo import main, parse_workbook

HEADERS = ["Nombre", "Fecha Hora", "Estado", "PRESENTISMO", "Observaciones"]


def _block(ws, col, name, rows, summary_row=None, mark=None):
    ws.cell(3, col, name)
    for i, h in enumerate(HEADERS):
        ws.cell(4, col + i, h)
    for r, (late, nota) in enumerate(rows, start=5):
        ws.cell(r, col, name)
        if late is not None:
            ws.cell(r, col + 3, late)
        if nota:
            ws.cell(r, col + 4, nota)
    if summary_row:
        ws.cell(summary_row, col, "PRESENTISMO")
        ws.cell(summary_row, col + 1, mark)


@pytest.fixture
def excel(tmp_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Centro"
    _block(ws, 1, "Ana Perez", [(None, None), (None, None)], 20, "SI")
    _block(ws, 8, "Beto Gomez", [(7, None), (10, None), (None, None)])
    _block(ws, 15, "Cami Diaz", [(None, "falta con aviso"), (None, "Falta")])
    ws2 = wb.create_sheet("Oficina")
    _block(ws2, 1, "Cami Diaz", [(5, None)])  # misma persona, otra hoja
    wb.create_sheet("Informe")["B5"] = "ANA PEREZ"  # se ignora
    path = tmp_path / "control.xlsx"
    wb.save(path)
    return path


def test_minutos_tarde_se_suman_por_persona(excel):
    personas = {p.nombre: p for p in parse_workbook(excel)}
    assert personas["Beto Gomez"].minutos_tarde == 17
    assert personas["Beto Gomez"].dias_tarde == 2
    assert not personas["Beto Gomez"].cobra_presentismo()
    assert personas["Beto Gomez"].cobra_presentismo(max_late_minutes=20)


def test_marca_manual_y_persona_sin_incidentes(excel):
    ana = {p.nombre: p for p in parse_workbook(excel)}["Ana Perez"]
    assert ana.marca_manual == "SI"
    assert ana.cobra_presentismo()


def test_faltas_y_excusadas(excel):
    cami = {p.nombre: p for p in parse_workbook(excel)}["Cami Diaz"]
    assert cami.faltas == 2
    assert cami.faltas_que_cuentan(re.compile("aviso")) == 1
    assert not cami.cobra_presentismo(max_late_minutes=99, max_faltas=0)


def test_persona_en_dos_hojas_se_une(excel):
    cami = [p for p in parse_workbook(excel) if p.nombre == "Cami Diaz"]
    assert len(cami) == 1
    assert cami[0].local == "Centro/Oficina"
    assert cami[0].minutos_tarde == 5


def test_cli_escribe_csv(excel, tmp_path, capsys):
    out = tmp_path / "out.csv"
    assert main([str(excel), "--max-late-minutes", "20", "--csv", str(out)]) == 0
    assert "SIN PRESENTISMO" in capsys.readouterr().out
    assert "Ana Perez" in out.read_text(encoding="utf-8-sig")
