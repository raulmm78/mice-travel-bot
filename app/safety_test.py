from __future__ import annotations

import os
import tempfile
from pathlib import Path
from datetime import date

from openpyxl import Workbook, load_workbook

from process_emails import (
    TravelRequest,
    is_travel_request_email,
    write_basic_excel,
    write_nn_excel,
)


def assert_true(condition: bool, message: str) -> None:
    if not condition:
        raise AssertionError(message)


def request(nombre: str, dni: str, email_id: str = "mail") -> TravelRequest:
    return TravelRequest(
        email_id=email_id,
        source_message_id=f"<{email_id}@example.com>",
        evento="Congreso IMS",
        nombre=nombre,
        apellidos="TEST",
        dni=dni,
        email=f"{nombre.lower()}@example.com",
        origen="Madrid",
        destino="Rio",
        fecha_viaje="2026-09-28",
        estado="ok",
    )


def test_filter() -> None:
    assert_true(
        is_travel_request_email(
            "RV: NUEVA SOLICITUD: Congreso IMS - TEST",
            "Texto previo\nDATOS PARA PETICIONES DE INVITADOS\nNOMBRE: TEST",
        ),
        "No reconoce reenvios RV/FW validos con formulario",
    )
    assert_true(
        not is_travel_request_email("NUEVA SOLICITUD: otra cosa", "Correo sin formulario"),
        "No esta filtrando correos sin formulario",
    )


def test_basic_append() -> None:
    with tempfile.TemporaryDirectory() as temp:
        path = Path(temp) / "basic.xlsx"
        old = request("ANTIGUO", "11111111H", "old")
        new = request("NUEVO", "22222222J", "new")
        write_basic_excel([old], path, "TEST")
        write_basic_excel([old, new], path, "TEST")
        write_basic_excel([old, new], path, "TEST")
        ws = load_workbook(path)["Totales"]
        dni_values = [ws.cell(row=row, column=5).value for row in range(3, ws.max_row + 1)]
        assert_true(dni_values.count("11111111H") == 1, "La fila antigua se ha duplicado o perdido")
        assert_true(dni_values.count("22222222J") == 1, "La fila nueva no se ha anadido una sola vez")


def test_nn_append() -> None:
    with tempfile.TemporaryDirectory() as temp:
        path = Path(temp) / "existing.xlsx"
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Totales"
        for col in range(1, 46):
            sheet.cell(row=3, column=col).value = f"H{col}"
        sheet["D4"] = "CONGRESO IMS"
        sheet["M4"] = "ANTIGUO"
        sheet["N4"] = "TEST"
        sheet["O4"] = "11111111H"
        sheet["U4"] = "antiguo@example.com"
        sheet["X4"] = date(2026, 9, 28)
        workbook.save(path)

        os.environ["NN_TEMPLATE_PATH"] = str(path)
        old = request("ANTIGUO", "11111111H", "old")
        new = request("NUEVO", "22222222J", "new")
        write_nn_excel([old, new], path, "TEST")
        write_nn_excel([old, new], path, "TEST")

        sheet = load_workbook(path)["Totales"]
        assert_true(sheet["M4"].value == "ANTIGUO", "La fila antigua ha cambiado")
        assert_true(sheet["O4"].value == "11111111H", "El DNI antiguo ha cambiado")
        assert_true(sheet["M5"].value == "NUEVO", "No se ha anadido la fila nueva")
        assert_true(sheet["O5"].value == "22222222J", "El DNI nuevo no esta en la fila esperada")
        assert_true(sheet["O6"].value in (None, ""), "Se ha duplicado la fila nueva")


def main() -> None:
    test_filter()
    test_basic_append()
    test_nn_append()
    print("OK - Seguridad Excel verificada")
    print("- Filtra correos que no son formulario")
    print("- Acepta reenvios RV/FW si contienen el formulario")
    print("- No borra filas antiguas")
    print("- No duplica la misma solicitud al repetir proceso")


if __name__ == "__main__":
    main()
