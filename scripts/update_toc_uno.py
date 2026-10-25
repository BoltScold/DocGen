import os
import sys
import time
from pathlib import Path

import uno

from com.sun.star.beans import PropertyValue


def make_property(name, value):
    prop = PropertyValue()
    prop.Name = name
    prop.Value = value
    return prop


def connect_to_libreoffice(
    host: str,
    port: int,
    max_attempts: int = 40,
    delay_seconds: float = 0.25,
):
    local_context = uno.getComponentContext()

    resolver = (
        local_context.ServiceManager
        .createInstanceWithContext(
            "com.sun.star.bridge.UnoUrlResolver",
            local_context,
        )
    )

    connection_string = (
        f"uno:socket,host={host},port={port};"
        "urp;StarOffice.ComponentContext"
    )

    last_error = None

    for _ in range(max_attempts):
        try:
            return resolver.resolve(connection_string)
        except Exception as exc:
            last_error = exc
            time.sleep(delay_seconds)

    raise RuntimeError(
        "Verbindung zu LibreOffice konnte nicht hergestellt werden. "
        f"Letzter Fehler: {last_error}"
    )


def update_document_indexes(document):
    indexes = document.getDocumentIndexes()
    count = indexes.getCount()

    print(f"Gefundene Dokumentindizes: {count}")

    if count == 0:
        raise RuntimeError(
            "Das Dokument enthält keine von LibreOffice erkannten "
            "Dokumentindizes. Das Inhaltsverzeichnis der Vorlage "
            "muss geprüft werden."
        )

    for index_number in range(count):
        document_index = indexes.getByIndex(index_number)

        try:
            document_index.update()
        except Exception:
            document_index.refresh()

        print(
            f"Dokumentindex {index_number + 1}/{count} aktualisiert."
        )


def update_toc(
    document_path: str,
    host: str = "127.0.0.1",
    port: int = 2002,
):
    document_path = os.path.abspath(document_path)

    if not os.path.isfile(document_path):
        raise FileNotFoundError(
            f"DOCX-Datei nicht gefunden: {document_path}"
        )

    context = connect_to_libreoffice(host, port)

    service_manager = context.ServiceManager

    desktop = service_manager.createInstanceWithContext(
        "com.sun.star.frame.Desktop",
        context,
    )

    document_url = uno.systemPathToFileUrl(document_path)

    load_properties = (
        make_property("Hidden", True),
        make_property("ReadOnly", False),
        make_property("UpdateDocMode", 3),
    )

    document = None

    try:
        document = desktop.loadComponentFromURL(
            document_url,
            "_blank",
            0,
            load_properties,
        )

        if document is None:
            raise RuntimeError(
                f"LibreOffice konnte das Dokument nicht öffnen: "
                f"{document_path}"
            )

        if not document.supportsService(
            "com.sun.star.text.TextDocument"
        ):
            raise RuntimeError(
                "Die Datei wurde von LibreOffice nicht als "
                "Writer-Dokument erkannt."
            )

        update_document_indexes(document)

        try:
            document.calculateAll()
        except Exception:
            # calculateAll ist nicht bei jedem Dokumentmodell relevant.
            pass

        document.store()

        print(
            f"Dokument erfolgreich gespeichert: {document_path}"
        )

    finally:
        if document is not None:
            try:
                document.close(True)
            except Exception:
                try:
                    document.dispose()
                except Exception:
                    pass


def main():
    if len(sys.argv) < 2:
        raise ValueError(
            "Aufruf: update_toc_uno.py <docx_path> [port]"
        )

    document_path = sys.argv[1]

    port = int(sys.argv[2]) if len(sys.argv) >= 3 else 2002

    update_toc(
        document_path=document_path,
        port=port,
    )


if __name__ == "__main__":
    main()