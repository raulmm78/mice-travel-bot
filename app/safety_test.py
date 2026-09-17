from __future__ import annotations

import os
import json
import shutil
import subprocess
import tempfile
import threading
import time
import urllib.error
import urllib.request
from http.server import ThreadingHTTPServer
from html.parser import HTMLParser
from pathlib import Path
from datetime import date
from unittest.mock import patch

from openpyxl import Workbook, load_workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Border, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation

from process_emails import (
    TravelRequest,
    is_travel_request_email,
    write_basic_excel,
    write_nn_excel,
    workbook_digest,
    save_workbook_safely,
)
import process_emails


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


def test_imap_gate_before_processing() -> None:
    with patch.object(process_emails, "test_imap_connection", side_effect=RuntimeError("IMAP no disponible")):
        with patch.object(process_emails, "run_bot_once") as process:
            try:
                process_emails.start_bot()
            except RuntimeError:
                pass
            else:
                raise AssertionError("El bot se activo sin comprobar IMAP")
            process.assert_not_called()


def test_dashboard_checks_imap_before_on() -> None:
    with patch.object(process_emails, "rows_as_dicts", return_value=[]):
        with patch.object(process_emails, "excel_is_ready", return_value=True):
            with patch.object(process_emails, "recent_logs", return_value=[]):
                with patch.object(process_emails, "pending_event_routes", return_value=[]):
                    html = process_emails.dashboard_html()
    assert_true('id="powerButton"' in html and "disabled" in html, "ON no espera a la comprobacion IMAP")
    assert_true("verifyImap();" in html and f"MICE TRAVEL BOT {process_emails.APP_VERSION}" in html,
                "Falta la verificacion IMAP del panel")


def test_ignored_event_stays_ignored_without_touching_excels() -> None:
    with tempfile.TemporaryDirectory() as temp:
        root = Path(temp)
        global_path = root / "global.xlsx"
        routes_path = root / "event_routes.json"
        old = request("EXISTENTE", "11111111H", "old")
        old.evento = "Congreso ADA"
        write_basic_excel([old], global_path, "GLOBAL")
        original = global_path.read_bytes()
        with patch.dict(os.environ, {"EVENT_ROUTES_PATH": str(routes_path)}):
            process_emails.ignore_event("Congreso ÁDA")
            process_emails.remember_pending_event("CONGRESO ADA", [old])
            route = process_emails.load_event_routes()["CONGRESO ADA"]
            assert_true(route["status"] == "ignored", "El evento se reactivo al procesar nuevos correos")
            assert_true(not process_emails.pending_event_routes([{"evento": "Congreso ADA"}]),
                        "El evento ignorado sigue pendiente en el panel")
            with patch.object(process_emails, "extract_request_auto", return_value=old):
                with patch.object(process_emails, "write_outputs") as write:
                    with patch.object(process_emails, "mark_processed_message_ids"):
                        rows = process_emails.process_paths([root / "mail.txt"], use_openai=False)
                        assert_true(rows == [], "El evento ignorado llego al volcado")
                        write.assert_called_once_with([])
        assert_true(global_path.read_bytes() == original, "Ignorar ha modificado el Excel existente")


def test_event_buttons_do_not_reprocess_all_mail() -> None:
    with tempfile.TemporaryDirectory() as temp:
        root = Path(temp)
        new_path = root / "event.xlsx"
        sample = request("NUEVO", "22222222J")
        with patch.object(process_emails, "rows_as_dicts", return_value=[sample.__dict__]):
            with patch.object(process_emails, "process_all", side_effect=AssertionError("Reproceso inesperado")):
                with patch.object(process_emails, "suggested_event_workbook_path", return_value=new_path):
                    with patch.object(process_emails, "write_event_excel_rows") as write:
                        with patch.object(process_emails, "assign_event_excel") as assign:
                            assert_true(process_emails.create_event_excel("Congreso IMS") == new_path,
                                        "Crear Excel no devolvio la ruta del evento")
                            write.assert_called_once()
                            assign.assert_called_once()
                new_path.touch()
                with patch.object(process_emails, "write_event_excel_rows") as write:
                    with patch.object(process_emails, "assign_event_excel") as assign:
                        assert_true(process_emails.choose_event_excel("Congreso IMS", new_path) == new_path,
                                    "Elegir Excel no devolvio la ruta del evento")
                        write.assert_called_once()
                        assign.assert_called_once()


def test_event_buttons_over_http() -> None:
    with tempfile.TemporaryDirectory() as temp:
        root = Path(temp)
        template = root / "template.xlsx"
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Listado"
        for column, label in {
            "M": "NOMBRE", "N": "APELLIDOS", "O": "DNI",
            "X": "FECHA INICIO SERVICIO", "AN": "ALERGIAS", "AQ": "OBSERVACIONES",
        }.items():
            sheet[f"{column}3"] = label
        sheet["AQ10"].fill = PatternFill(fill_type="solid", fgColor="FFFFFF")
        workbook.save(template)
        destination = root / "destino"
        destination.mkdir()
        global_path = root / "global.xlsx"
        global_path.write_bytes(b"GLOBAL SIN CAMBIOS")
        first = request("ANA", "11111111H", "first")
        first.evento = "Congreso ADA"
        second = request("LUIS", "22222222J", "second")
        second.evento = "Congreso EASD"
        environment = patch.dict(os.environ, {
            "EVENT_ROUTES_PATH": str(root / "routes.json"),
            "EVENT_TEMPLATE_PATH": str(template),
        })
        with environment, patch.object(process_emails, "rows_as_dicts", return_value=[first.__dict__, second.__dict__]), \
                patch.object(process_emails, "event_output_root", return_value=root), \
                patch.object(process_emails, "xlsx_path", return_value=global_path):
            server = ThreadingHTTPServer(("127.0.0.1", 0), process_emails.DashboardHandler)
            thread = threading.Thread(target=server.serve_forever, daemon=True)
            thread.start()
            base = f"http://127.0.0.1:{server.server_port}"
            def send(path: str, data: dict[str, str]) -> tuple[int, dict]:
                req = urllib.request.Request(
                    base + path, data=json.dumps(data).encode(),
                    headers={"Content-Type": "application/json"}, method="POST",
                )
                try:
                    response = urllib.request.urlopen(req, timeout=10)
                except urllib.error.HTTPError as exc:
                    response = exc
                with response:
                    return response.status, json.load(response)
            try:
                with urllib.request.urlopen(base + "/api/create-event-options?event_name=Congreso%20ADA", timeout=10) as response:
                    proposal = json.load(response)
                assert_true(proposal["filename"].endswith(".xlsx"), "Falta nombre propuesto")
                status, created = send("/api/create-event-excel", {
                    "event_name": "Congreso ADA", "folder": str(destination), "filename": proposal["filename"],
                })
                assert_true(status == 200, f"Crear Excel fallo: {created}")
                created_path = Path(created["created"])
                assert_true(created_path.parent == destination and created_path.exists(),
                            "Crear Excel no respeto la carpeta elegida")
                status, duplicate = send("/api/create-event-excel", {
                    "event_name": "Congreso ADA", "folder": str(destination), "filename": proposal["filename"],
                })
                assert_true(status == 400 and "no se ha sobrescrito" in duplicate["error"],
                            "Crear Excel permite sobrescribir un archivo existente")
                with urllib.request.urlopen(base + "/api/event-excels", timeout=10) as response:
                    files = json.load(response)["files"]
                assert_true(str(created_path) in files, "Elegir Excel no muestra el Excel creado")
                status, assigned = send("/api/assign-event-excel", {
                    "event_name": "Congreso EASD", "excel_path": str(created_path),
                })
                assert_true(status == 200, f"Elegir Excel fallo: {assigned}")
                sheet = load_workbook(created_path)["Listado"]
                assert_true(sheet["O4"].value == "11111111H" and sheet["O5"].value == "22222222J",
                            "Elegir Excel no anadio la fila debajo de la existente")
                status, error = send("/api/assign-event-excel", {
                    "event_name": "Congreso EASD", "excel_path": str(root / "no_existe.xlsx"),
                })
                assert_true(status == 400 and "error" in error, "No comunica un Excel inexistente")
                assert_true(global_path.read_bytes() == b"GLOBAL SIN CAMBIOS", "Los botones modificaron el global")
            finally:
                server.shutdown()
                server.server_close()
                thread.join(timeout=2)


def test_event_javascript_syntax() -> None:
    if not shutil.which("node"):
        return
    with patch.object(process_emails, "rows_as_dicts", return_value=[]):
        page = process_emails.dashboard_html()
    script = page.split("<script>", 1)[1].split("</script>", 1)[0]
    result = subprocess.run(["node", "--check"], input=script, text=True, capture_output=True)
    assert_true(result.returncode == 0, f"Error JavaScript del panel: {result.stderr}")


def test_event_button_handlers_are_valid_javascript() -> None:
    class Buttons(HTMLParser):
        def __init__(self) -> None:
            super().__init__()
            self.handlers: list[str] = []

        def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
            if tag == "button":
                handler = dict(attrs).get("onclick")
                if handler and "Event" in handler:
                    self.handlers.append(handler)

    sample = request("TEST", "12345678Z")
    sample.evento = "Reunion d'Alvarez"
    with patch.object(process_emails, "rows_as_dicts", return_value=[sample.__dict__]):
        with patch.object(process_emails, "recent_logs", return_value=[]):
            page = process_emails.dashboard_html()
    parser = Buttons()
    parser.feed(page)
    assert_true(len(parser.handlers) >= 3, "No aparecen los tres botones del evento")
    if shutil.which("node"):
        for handler in parser.handlers:
            result = subprocess.run(["node", "--check"], input=handler, text=True, capture_output=True)
            assert_true(result.returncode == 0, f"Boton con JavaScript roto: {handler}")


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
        template_path = Path(temp) / "template.xlsx"
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
        sheet["AQ5"] = "NOTA MANUAL"
        workbook.save(path)
        workbook.save(template_path)

        os.environ["NN_TEMPLATE_PATH"] = str(template_path)
        old = request("ANTIGUO", "11111111H", "old")
        new = request("NUEVO", "22222222J", "new")
        write_nn_excel([old, new], path, "TEST")
        write_nn_excel([old, new], path, "TEST")

        sheet = load_workbook(path)["Totales"]
        assert_true(sheet["M4"].value == "ANTIGUO", "La fila antigua ha cambiado")
        assert_true(sheet["O4"].value == "11111111H", "El DNI antiguo ha cambiado")
        assert_true(sheet["AQ5"].value == "NOTA MANUAL", "Se ha borrado una celda manual")
        assert_true(sheet["M6"].value == "NUEVO", "No se ha anadido la fila nueva debajo")
        assert_true(sheet["O6"].value == "22222222J", "El DNI nuevo no esta en la fila esperada")
        assert_true(sheet["O7"].value in (None, ""), "Se ha duplicado la fila nueva")
        backups = list((path.parent / "_bot_backups").glob("*.xlsx"))
        assert_true(len(backups) == 1, "Falta la copia del Excel anterior")
        assert_true(load_workbook(backups[0])["Totales"]["AQ5"].value == "NOTA MANUAL",
                    "La copia de seguridad no conserva la nota manual")


def test_sparse_formatted_rows_append_position() -> None:
    sheet = Workbook().active
    sheet.title = "Totales"
    sheet["O4"] = "11111111H"
    sheet["AR5"] = "NOTA MANUAL"
    sheet["A65519"].fill = PatternFill(fill_type="solid", fgColor="FFFFFF")
    before = len(sheet._cells)
    start = time.monotonic()
    row = process_emails.nn_next_append_row(sheet)
    assert_true(row == 6, "La nueva solicitud no iria debajo de los datos existentes")
    assert_true(len(sheet._cells) == before, "Buscar la siguiente fila creo celdas vacias")
    assert_true(time.monotonic() - start < 3, "Buscar la siguiente fila tarda demasiado")
    sheet["AR42"] = "TOTAL"
    assert_true(process_emails.nn_next_append_row(sheet) == 6,
                "No se eligio el primer hueco libre antes de los totales")


def test_nn_uses_each_free_row_without_overwriting() -> None:
    with tempfile.TemporaryDirectory() as temp:
        path = Path(temp) / "global.xlsx"
        template = Path(temp) / "template.xlsx"
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Totales"
        sheet["O4"] = "11111111H"
        sheet["AQ5"] = "NOTA MANUAL"
        sheet["U7"] = "mailto:conservar@example.com"
        sheet["AH42"] = "=SUM(AH4:AH40)"
        workbook.save(path)
        workbook.save(template)
        first = request("PRIMERO", "22222222J", "first-free")
        second = request("SEGUNDO", "33333333P", "second-free")
        write_nn_excel([first, second], path, "GLOBAL", template)
        result = load_workbook(path)["Totales"]
        assert_true(result["O6"].value == "22222222J", "No uso el primer hueco libre")
        assert_true(result["O8"].value == "33333333P", "La segunda fila piso datos existentes")
        assert_true(result["AQ5"].value == "NOTA MANUAL", "Se borro una nota")
        assert_true(result["U7"].value == "mailto:conservar@example.com", "Se borro un correo")
        assert_true(result["AH42"].value == "=SUM(AH4:AH40)", "Se borro una formula")


def test_conflict_stops_before_replace() -> None:
    with tempfile.TemporaryDirectory() as temp:
        path = Path(temp) / "original.xlsx"
        old_book = Workbook()
        old_book.active["A1"] = "ORIGINAL"
        old_book.save(path)
        original_digest = workbook_digest(path)
        new_book = Workbook()
        new_book.active["A1"] = "CAMBIO"
        try:
            save_workbook_safely(new_book, path, "digest-desactualizado")
        except RuntimeError:
            pass
        else:
            raise AssertionError("No se detuvo el guardado ante un cambio concurrente")
        assert_true(workbook_digest(path) == original_digest, "El conflicto modifico el Excel")
        try:
            save_workbook_safely(new_book, path)
        except FileExistsError:
            pass
        else:
            raise AssertionError("Crear Excel sustituyo un archivo ya existente")
        assert_true(workbook_digest(path) == original_digest, "Crear Excel modifico un archivo existente")


def test_global_and_event_flow() -> None:
    with tempfile.TemporaryDirectory() as temp:
        root = Path(temp)
        template = root / "plantilla.xlsx"
        event_template = root / "plantilla_eventos.xlsx"
        global_path = root / "global.xlsx"
        events = root / "eventos"
        events.mkdir()
        workbook = Workbook()
        workbook.active.title = "Totales"
        workbook.active["M4"] = "PERSONA DE PLANTILLA"
        workbook.active["O4"] = "99999999R"
        workbook.save(template)
        template_digest = workbook_digest(template)
        event_book = Workbook()
        event_sheet = event_book.active
        event_sheet.title = "Listado"
        event_book.create_sheet("Condiciones")
        for column, label in {
            "M": "NOMBRE", "N": "APELLIDOS", "O": "DNI", "X": "FECHA INICIO SERVICIO",
            "AN": "ALERGIAS", "AQ": "OBSERVACIONES",
        }.items():
            event_sheet[f"{column}3"] = label
        event_sheet["A4"] = 1
        for row_num in range(4, 8):
            event_sheet[f"BE{row_num}"] = f"=SUM(AR{row_num}:BD{row_num})"
        event_sheet["BI5"] = 0
        event_sheet["DM4"] = "AIR ALGERIE"
        event_sheet["M4"].border = Border(left=Side(style="thin"))
        event_sheet["M4"].fill = PatternFill("solid", fgColor="FFFFFF")
        validation = DataValidation(type="list", formula1='"A,B"')
        event_sheet.add_data_validation(validation)
        validation.add("M4:M7")
        event_sheet.conditional_formatting.add("M4:M7", CellIsRule(operator="equal", formula=["0"]))
        event_book.save(event_template)
        old = request("ANTIGUO", "11111111H", "old")
        new = request("NUEVO", "22222222J", "new")
        other = request("OTRO", "33333333P", "other")
        other.evento = "Congreso EULAR"
        other.conex_ida = "DIRECTO"
        other.hotel_in = "2026-09-28"
        other.restricciones_alimentarias = "SIN LACTOSA"
        existing_event = events / "NO ENVIAR -----LISTADO CONGRESO IMS 28 SEP.xlsx"
        env = {
            "NN_TEMPLATE_PATH": str(template),
            "EVENT_TEMPLATE_PATH": str(event_template),
            "OUTPUT_DIR": str(root),
            "XLSX_PATH": str(global_path),
            "EVENT_OUTPUT_DIR": str(events),
            "EVENT_ROUTES_PATH": str(root / "routes.json"),
        }
        with patch.dict(os.environ, env):
            try:
                write_nn_excel([request("ERROR", "88888888Q")], template, "NO DEBE ESCRIBIR")
            except ValueError:
                pass
            else:
                raise AssertionError("Se permitio escribir sobre la plantilla")
            write_nn_excel([old], global_path, "GLOBAL")
            write_nn_excel([old], existing_event, "Congreso IMS")
            process_emails.assign_event_excel("Congreso IMS", existing_event)
            process_emails.write_outputs([old, new, other])
            assert_true(workbook_digest(template) == template_digest, "La plantilla ha cambiado")
            assert_true(existing_event.exists(), "Se ha eliminado el Excel de evento")
            suggested = process_emails.suggested_event_workbook_path("Congreso EULAR", [other])
            assert_true(not suggested.exists(), "Se ha creado un Excel antes de confirmar el evento")
            with patch.object(process_emails, "process_all", return_value=[old, new, other]):
                created = process_emails.create_event_excel("Congreso EULAR")
            assert_true(created == suggested and created.exists(), "No se ha creado el Excel sugerido")
            process_emails.write_outputs([old, new, other])
            next_attendee = request("CUARTO", "44444444L", "fourth")
            next_attendee.evento = "Congreso EULAR"
            process_emails.write_nn_excel([other, next_attendee], suggested, "Congreso EULAR",
                                          process_emails.event_template_path())

        global_sheet = load_workbook(global_path)["Totales"]
        event_sheet = load_workbook(existing_event)["Totales"]
        new_book = load_workbook(suggested)
        new_sheet = new_book["Listado"]
        source_sheet = load_workbook(event_template)["Listado"]
        assert_true([global_sheet[f"O{row}"].value for row in (4, 5, 6)] ==
                    ["11111111H", "22222222J", "33333333P"], "El global perdio o duplico filas")
        assert_true([event_sheet[f"O{row}"].value for row in (4, 5)] ==
                    ["11111111H", "22222222J"], "El Excel de evento perdio o duplico filas")
        assert_true(new_sheet["O4"].value == "33333333P" and new_sheet["O5"].value == "44444444L"
                    and new_sheet["O6"].value is None, "El evento nuevo no anade debajo sin duplicar")
        assert_true(new_sheet["Y4"].value == "DIRECTO", "La conexion de ida fue a otra columna")
        assert_true(new_sheet["AD4"].value is not None, "La fecha de entrada del hotel no esta en IN")
        assert_true(new_sheet["AN4"].value == "SIN LACTOSA", "La alergia no esta en su columna")
        assert_true(new_sheet["BE4"].value == source_sheet["BE4"].value,
                    "Se ha perdido una formula de la plantilla")
        assert_true(new_sheet["DM4"].value == source_sheet["DM4"].value,
                    "Se ha perdido la lista auxiliar de la plantilla")
        assert_true(new_sheet["M4"].fill.patternType == source_sheet["M4"].fill.patternType
                    and new_sheet["M4"].border.left.style == source_sheet["M4"].border.left.style,
                    "Se ha cambiado el formato base de la plantilla")
        assert_true("Condiciones" in new_book.sheetnames, "Falta una hoja de la plantilla")
        assert_true(len(new_sheet.data_validations.dataValidation) == len(source_sheet.data_validations.dataValidation),
                    "Se han perdido validaciones de la plantilla")
        assert_true(len(new_sheet.conditional_formatting) == len(source_sheet.conditional_formatting),
                    "Se han perdido formatos condicionales de la plantilla")


def test_supplied_template_when_available() -> None:
    template = process_emails.BUNDLED_NN_TEMPLATE_PATH
    if not template.exists():
        return
    with tempfile.TemporaryDirectory() as temp:
        path = Path(temp) / "nuevo_evento.xlsx"
        attendee = request("CLIENTE", "55555555M", "supplied")
        attendee.conex_ida = "DIRECTO"
        attendee.restricciones_alimentarias = "SIN GLUTEN"
        write_nn_excel([attendee], path, "Congreso IMS", template)
        source = load_workbook(template)["Listado"]
        output = load_workbook(path)
        sheet = output["Listado"]
        assert_true(sheet["O4"].value == "55555555M", "La plantilla real no usa la primera fila")
        assert_true(sheet["Y4"].value == "DIRECTO" and sheet["AN4"].value == "SIN GLUTEN",
                    "La plantilla real tiene columnas desalineadas")
        assert_true(sheet["BE4"].value == source["BE4"].value, "La plantilla real perdio formulas")
        assert_true(sheet["DM4"].value == source["DM4"].value, "La plantilla real perdio listas")
        assert_true(len(sheet.conditional_formatting) == len(source.conditional_formatting),
                    "La plantilla real perdio formato condicional")
        assert_true("Condiciones" in output.sheetnames, "La plantilla real perdio hojas")


def main() -> None:
    test_filter()
    test_imap_gate_before_processing()
    test_dashboard_checks_imap_before_on()
    test_ignored_event_stays_ignored_without_touching_excels()
    test_event_buttons_do_not_reprocess_all_mail()
    test_event_buttons_over_http()
    test_event_javascript_syntax()
    test_event_button_handlers_are_valid_javascript()
    test_basic_append()
    test_nn_append()
    test_sparse_formatted_rows_append_position()
    test_nn_uses_each_free_row_without_overwriting()
    test_conflict_stops_before_replace()
    test_global_and_event_flow()
    test_supplied_template_when_available()
    print("OK - Seguridad Excel verificada")
    print("- Filtra correos que no son formulario")
    print("- Acepta reenvios RV/FW si contienen el formulario")
    print("- No borra filas antiguas")
    print("- No duplica la misma solicitud al repetir proceso")


if __name__ == "__main__":
    main()
