import json
import os
import shutil
import socket
import subprocess
import sys
import tempfile
import threading
import time
from contextlib import contextmanager
from datetime import date, datetime

import uno
from com.sun.star.beans import PropertyValue
from com.sun.star.connection import NoConnectException


@contextmanager
def show_progress(message):
    print(message + ".", end="", flush=True)
    finished = threading.Event()

    def print_dots():
        while not finished.wait(1):
            print(".", end="", flush=True)

    progress_thread = threading.Thread(target=print_dots, daemon=True)
    progress_thread.start()
    try:
        yield
    finally:
        finished.set()
        progress_thread.join()
        print(flush=True)


def property_value(name, value):
    prop = PropertyValue()
    prop.Name = name
    prop.Value = value
    return prop


def connect_to_libreoffice(port):
    local_context = uno.getComponentContext()
    resolver = local_context.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver",
        local_context,
    )
    connection_url = (
        "uno:socket,host=127.0.0.1,port={};urp;StarOffice.ComponentContext".format(port)
    )
    last_error = None

    for _ in range(100):
        try:
            return resolver.resolve(connection_url)
        except NoConnectException as error:
            last_error = error
            time.sleep(0.1)

    raise RuntimeError("Le serveur UNO de LibreOffice n'a pas démarré.") from last_error


def wait_for_server_shutdown(port):
    deadline = time.monotonic() + 10
    while time.monotonic() < deadline:
        try:
            with socket.create_connection(("127.0.0.1", port), timeout=0.25):
                pass
        except OSError:
            return True
        time.sleep(0.25)
    return False


def date_serial(value, null_date):
    parsed_date = datetime.strptime(value, "%d/%m/%Y").date()
    origin = date(null_date.Year, null_date.Month, null_date.Day)
    return (parsed_date - origin).days


def remove_profile_directory(profile_directory):
    for attempt in range(20):
        try:
            shutil.rmtree(profile_directory)
            return
        except FileNotFoundError:
            return
        except PermissionError as error:
            if attempt == 19:
                print(
                    "Impossible de supprimer le profil temporaire LibreOffice "
                    "'{}' : {}".format(profile_directory, error),
                    file=sys.stderr,
                )
                return
            time.sleep(0.5)


def clean_up_libreoffice(profile_directory, port, process, desktop):
    try:
        if desktop is not None:
            with show_progress("Arrêt du service LibreOffice"):
                desktop.terminate()
    finally:
        if process is not None:
            if process.poll() is None:
                with show_progress("Attente de l'arrêt de LibreOffice"):
                    try:
                        process.wait(timeout=10)
                    except subprocess.TimeoutExpired:
                        process.terminate()
                        try:
                            process.wait(timeout=5)
                        except subprocess.TimeoutExpired:
                            process.kill()
                            process.wait()
            else:
                process.wait()

        if wait_for_server_shutdown(port):
            print("Service LibreOffice arrêté.", flush=True)
            print("Nettoyage du profil temporaire...", flush=True)
            remove_profile_directory(profile_directory)
        else:
            print(
                "LibreOffice utilise encore le profil temporaire '{}'; "
                "ce profil n'a pas été supprimé.".format(profile_directory),
                file=sys.stderr,
            )


def set_cell_values(sheet, row_index, transaction, null_date):
    sheet.getCellByPosition(0, row_index).setValue(
        date_serial(str(transaction["Date"]), null_date)
    )
    sheet.getCellByPosition(1, row_index).setString(str(transaction["Type"]))

    operation_date = str(transaction["Date opération"])
    operation_cell = sheet.getCellByPosition(2, row_index)
    if operation_date:
        operation_cell.setValue(date_serial(operation_date, null_date))
    else:
        operation_cell.setString("")

    sheet.getCellByPosition(3, row_index).setString(
        str(transaction["Débiteur/Destinataire"])
    )
    amount = str(transaction["Montant(EUROS)"]).strip()
    amount = amount.replace(" ", "").replace(",", ".")
    sheet.getCellByPosition(4, row_index).setValue(float(amount))
    # sheet.getCellByPosition(5, row_index).setString("Non identifié")
    # sheet.getCellByPosition(6, row_index).setString("")


def update_document(document_path, transactions_path, new_document, account_number):
    with open(transactions_path, "r", encoding="utf-8-sig") as file:
        transactions = json.load(file)

    program_directory = os.path.dirname(os.path.abspath(sys.executable))
    soffice_path = os.path.join(program_directory, "soffice.exe")
    if not os.path.isfile(soffice_path):
        raise FileNotFoundError(
            "soffice.exe est introuvable dans le dossier du Python LibreOffice."
        )

    with socket.socket() as listener:
        listener.bind(("127.0.0.1", 0))
        port = listener.getsockname()[1]

    profile_directory = tempfile.mkdtemp(prefix="gestion-releves-lo-")
    profile_url = uno.systemPathToFileUrl(profile_directory)
    process = None
    desktop = None
    document = None

    try:
        process = subprocess.Popen(
            [
                soffice_path,
                "--headless",
                "--nologo",
                "--nodefault",
                "--nofirststartwizard",
                "-env:UserInstallation=" + profile_url,
                "--accept=socket,host=127.0.0.1,port={};urp;StarOffice.ServiceManager".format(
                    port
                ),
            ],
            cwd=program_directory,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )

        with show_progress("Connexion au service UNO de LibreOffice"):
            context = connect_to_libreoffice(port)
        desktop = context.ServiceManager.createInstanceWithContext(
            "com.sun.star.frame.Desktop",
            context,
        )
        document_url = uno.systemPathToFileUrl(os.path.abspath(document_path))
        with show_progress("Ouverture du classeur ODS dans Calc"):
            document = desktop.loadComponentFromURL(
                document_url,
                "_blank",
                0,
                (
                    property_value("Hidden", True),
                    property_value("ReadOnly", False),
                ),
            )
        if document is None:
            raise RuntimeError("LibreOffice n'a pas pu ouvrir le classeur ODS.")
        print("Classeur ODS ouvert.", flush=True)

        sheet_name = "données brutes"
        if not document.Sheets.hasByName(sheet_name):
            raise RuntimeError(
                "La feuille '{}' est introuvable dans le classeur ODS.".format(
                    sheet_name
                )
            )
        sheet = document.Sheets.getByName(sheet_name)
        model_row = 7
        first_inserted_row = model_row + 1

        if new_document:
            used_area = sheet.createCursor()
            used_area.gotoEndOfUsedArea(True)
            last_used_row = used_area.RangeAddress.EndRow
            if last_used_row >= first_inserted_row:
                sheet.getCellRangeByPosition(
                    0,
                    first_inserted_row,
                    used_area.RangeAddress.EndColumn,
                    last_used_row,
                ).clearContents(1 | 2 | 4 | 8 | 16 | 128 | 256)
            sheet.getCellByPosition(1, 0).setString(account_number)

        if transactions:
            used_area = sheet.createCursor()
            used_area.gotoEndOfUsedArea(True)
            last_column = max(6, used_area.RangeAddress.EndColumn)
            model_height = sheet.Rows.getByIndex(model_row).Height
            sheet.Rows.insertByIndex(model_row, len(transactions))
            model_row = model_row + len(transactions)
            for transaction in reversed(transactions):
                destination_row = model_row - 1
                model_range = sheet.getCellRangeByPosition(0, model_row, last_column, model_row)
                destination = sheet.getCellByPosition(0, destination_row).getCellAddress()
                sheet.copyRange(destination, model_range.RangeAddress)
                sheet.Rows.getByIndex(destination_row).Height = model_height
                model_row, destination_row = destination_row, model_row
                set_cell_values(sheet, destination_row, transaction, document.NumberFormatSettings.NullDate, )
        print("Enregistrement du classeur ODS...", flush=True)
        document.store()
        print("Enregistrement terminé.", flush=True)
        print("Transactions ajoutées sur la plage {} de la feuille '{}'.".format(
            sheet.getCellRangeByPosition(0, model_row + 1, last_column, model_row + len(transactions)).RangeAddress, sheet_name
        ), flush=True)
    finally:
        try:
            if document is not None:
                with show_progress("Fermeture du classeur ODS"):
                    document.close(True)
                print("Classeur ODS fermé.", flush=True)
        finally:
            clean_up_libreoffice(profile_directory, port, process, desktop)


def main():
    if len(sys.argv) != 5:
        raise ValueError(
            "Usage: AjouterTransactionsOds.py <classeur.ods> <transactions.json> "
            "<nouveau-classeur> <numero-compte>"
        )

    update_document(
        sys.argv[1],
        sys.argv[2],
        sys.argv[3].lower() == "true",
        sys.argv[4],
    )


if __name__ == "__main__":
    main()
