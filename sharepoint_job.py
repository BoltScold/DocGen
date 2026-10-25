# from flask import Flask, jsonify, request
# from concurrent.futures import ThreadPoolExecutor
# import uuid, time

# app = Flask(__name__)
# executor = ThreadPoolExecutor(max_workers=1)
# runs = {}

# def _run_and_store(run_id: str):
#     try:
#         runs[run_id]["status"] = "running"

#         # IMPORTANT: call your batch logic
#         main()

#         runs[run_id]["status"] = "completed"
#         runs[run_id]["finished"] = time.time()

#     except Exception as e:
#         runs[run_id]["status"] = "failed"
#         runs[run_id]["error"] = str(e)
#         runs[run_id]["finished"] = time.time()

# @app.get("/health")
# def health():
#     return jsonify(ok=True), 200,

# @app.post("/run")
# def run():
#     run_id = str(uuid.uuid4())
#     runs[run_id] = {"status": "queued", "started": time.time()}
#     executor.submit(_run_and_store, run_id)
#     return jsonify(run_id=run_id, status="queued"), 202

import os
import zipfile
from lxml import etree
from pypdf import PdfReader
import socket
import sys
import re
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import base64
from PIL import Image
import numpy as np
import subprocess
from io import BytesIO
import tempfile
import shutil
from typing import List, Dict, Any
from dotenv import load_dotenv
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.listitems.caml.query import CamlQuery
from pprint import pprint
from pathlib import Path
from collections import defaultdict

BASE_DIR = Path(__file__).resolve().parent

TEMPLATE_PATH = (
    BASE_DIR
    / "Dachcheck-Inhaltsverzeichnis.docx"
)

UNO_SCRIPT_PATH = (
    BASE_DIR
    / "scripts"
    / "update_toc_uno.py"
)

def find_free_local_port() -> int:
    with socket.socket(
        socket.AF_INET,
        socket.SOCK_STREAM,
    ) as sock:
        sock.bind(("127.0.0.1", 0))
        return sock.getsockname()[1]

def log_document_runtime_environment() -> None:
    """
    Protokolliert die für die DOCX-/TOC-Verarbeitung relevante
    Serverumgebung, jedoch keine geheimen Application Settings.
    """

    import platform

    print("=== Dokument-Runtime-Diagnose ===")
    print(f"Python: {sys.version}")
    print(f"Python executable: {sys.executable}")
    print(f"Platform: {platform.platform()}")
    print(f"Machine: {platform.machine()}")
    print(f"Working directory: {os.getcwd()}")
    print(f"Script directory: {BASE_DIR}")
    print(f"Temp directory: {tempfile.gettempdir()}")
    print(f"Template vorhanden: {TEMPLATE_PATH.is_file()}")
    print(f"Template-Pfad: {TEMPLATE_PATH}")

    candidates = [
        shutil.which("libreoffice"),
        shutil.which("soffice"),
        "/usr/bin/libreoffice",
        "/usr/bin/soffice",
        r"C:\Program Files\LibreOffice\program\soffice.exe",
    ]

    for candidate in candidates:
        if not candidate:
            continue

        print(
            f"LibreOffice-Kandidat: {candidate}; "
            f"vorhanden={os.path.isfile(candidate)}"
        )

    try:
        soffice_path = find_libreoffice_executable()

        print(
            f"LibreOffice gefunden: {soffice_path}"
        )

        version_result = subprocess.run(
            [soffice_path, "--version"],
            capture_output=True,
            text=True,
            timeout=30,
            check=False,
        )

        print(
            f"LibreOffice Exitcode: "
            f"{version_result.returncode}"
        )
        print(
            f"LibreOffice Version: "
            f"{version_result.stdout.strip()}"
        )

        if version_result.stderr:
            print(
                f"LibreOffice STDERR: "
                f"{version_result.stderr.strip()}"
            )

    except Exception as exc:
        print(
            f"LibreOffice nicht verwendbar: "
            f"{type(exc).__name__}: {exc}"
        )

    try:
        import lxml

        print(
            f"lxml importiert: "
            f"{lxml.__version__}"
        )
    except Exception as exc:
        print(
            f"lxml nicht importierbar: "
            f"{type(exc).__name__}: {exc}"
        )

    try:
        import pypdf

        print(
            f"pypdf importiert: "
            f"{pypdf.__version__}"
        )
    except Exception as exc:
        print(
            f"pypdf nicht importierbar: "
            f"{type(exc).__name__}: {exc}"
        )

    print("=== Ende Dokument-Runtime-Diagnose ===")

def render_docx_to_pdf_with_libreoffice(
    docx_path: str,
    output_directory: str,
    timeout_seconds: int = 180,
) -> str:
    """
    Rendert ein DOCX über LibreOffice als PDF.

    Das ursprüngliche DOCX wird nicht geöffnet und gespeichert,
    sondern ausschließlich gelesen und als PDF ausgegeben.
    """

    docx_path = os.path.abspath(docx_path)
    output_directory = os.path.abspath(output_directory)

    if not os.path.isfile(docx_path):
        raise FileNotFoundError(
            f"DOCX-Datei nicht gefunden: {docx_path}"
        )

    os.makedirs(
        output_directory,
        exist_ok=True,
    )

    soffice_path = find_libreoffice_executable()

    profile_directory = tempfile.mkdtemp(
        prefix="libreoffice_render_profile_"
    )

    profile_url = (
        Path(profile_directory)
        .resolve()
        .as_uri()
    )

    command = [
        soffice_path,
        "--headless",
        "--nologo",
        "--nodefault",
        "--nofirststartwizard",
        "--norestore",
        f"-env:UserInstallation={profile_url}",
        "--convert-to",
        "pdf",
        "--outdir",
        output_directory,
        docx_path,
    ]

    try:
        result = subprocess.run(
            command,
            capture_output=True,
            text=True,
            timeout=timeout_seconds,
            check=False,
        )

        expected_pdf_path = os.path.join(
            output_directory,
            Path(docx_path).stem + ".pdf",
        )

        if result.returncode != 0:
            raise RuntimeError(
                "LibreOffice-PDF-Rendering fehlgeschlagen.\n"
                f"Exitcode: {result.returncode}\n"
                f"STDOUT:\n{result.stdout}\n"
                f"STDERR:\n{result.stderr}"
            )

        if not os.path.isfile(expected_pdf_path):
            raise FileNotFoundError(
                "LibreOffice meldete keinen Fehler, "
                "aber die PDF-Datei wurde nicht gefunden: "
                f"{expected_pdf_path}\n"
                f"STDOUT:\n{result.stdout}\n"
                f"STDERR:\n{result.stderr}"
            )

        if os.path.getsize(expected_pdf_path) == 0:
            raise RuntimeError(
                f"Die erzeugte PDF-Datei ist leer: "
                f"{expected_pdf_path}"
            )

        return expected_pdf_path

    finally:
        shutil.rmtree(
            profile_directory,
            ignore_errors=True,
        )

def normalize_match_text(value: str) -> str:
    if not value:
        return ""

    value = value.replace("\u00ad", "")
    value = value.replace("\u00a0", " ")
    value = value.replace("–", "-")
    value = value.replace("—", "-")

    value = re.sub(
        r"\s+",
        " ",
        value,
    )

    return value.strip().casefold()

def normalize_compact_match_text(value: str) -> str:
    """
    Erstellt eine besonders tolerante Vergleichsform.

    Entfernt unter anderem:
    - Leerzeichen
    - Zeilenumbrüche
    - Trennstriche
    - Satzzeichen
    - Soft-Hyphens

    Beispiele:

    "W ÄRMEDÄMMUNG"
        -> "wärmedämmung"

    "PV-\nANLAGE"
        -> "pvanlage"

    "PV-Anlage"
        -> "pvanlage"
    """

    if not value:
        return ""

    value = value.replace("\u00ad", "")
    value = value.replace("\u00a0", " ")
    value = value.casefold()

    # Nur Buchstaben und Ziffern behalten.
    value = "".join(
        character
        for character in value
        if character.isalnum()
    )

    return value

def extract_pageref_entries(
    docx_path: str,
) -> list[dict]:
    """
    Liest sämtliche PAGEREF-Felder des vorhandenen
    Inhaltsverzeichnisses aus.

    Verwendet den sichtbaren TOC-Eintrag statt des Ziel-Bookmarks.
    Dadurch funktionieren auch fehlende oder falsch platzierte
    _Toc-Bookmarks.

    Rückgabe:
        [
            {
                "bookmark": "_Toc123456",
                "title": "Einleitung",
                "display_text": "1 Einleitung"
            },
            ...
        ]
    """

    docx_path = os.path.abspath(docx_path)

    if not os.path.isfile(docx_path):
        raise FileNotFoundError(
            f"DOCX-Datei nicht gefunden: {docx_path}"
        )

    namespaces = {
        "w": (
            "http://schemas.openxmlformats.org/"
            "wordprocessingml/2006/main"
        )
    }

    w_namespace = namespaces["w"]

    with zipfile.ZipFile(
        docx_path,
        "r",
    ) as archive:
        document_xml = archive.read(
            "word/document.xml"
        )

    parser = etree.XMLParser(
        remove_blank_text=False,
        resolve_entities=False,
    )

    root = etree.fromstring(
        document_xml,
        parser=parser,
    )

    entries = []
    seen_bookmarks = set()

    instruction_elements = root.xpath(
        "//w:instrText["
        "contains("
        "translate(., 'pageref', 'PAGEREF'), "
        "'PAGEREF'"
        ")"
        "]",
        namespaces=namespaces,
    )

    for instruction_element in instruction_elements:
        instruction = instruction_element.text or ""

        match = re.search(
            r"\bPAGEREF\s+([^\s\\]+)",
            instruction,
            flags=re.IGNORECASE,
        )

        if not match:
            continue

        bookmark_name = match.group(1)

        if bookmark_name in seen_bookmarks:
            continue

        seen_bookmarks.add(bookmark_name)

        # TOC-Einträge mit \h liegen normalerweise in einem
        # Hyperlink. Falls nicht, verwenden wir den Absatz.
        hyperlink_ancestors = instruction_element.xpath(
            "ancestor::w:hyperlink[1]",
            namespaces=namespaces,
        )

        if hyperlink_ancestors:
            entry_container = hyperlink_ancestors[0]
        else:
            paragraph_ancestors = instruction_element.xpath(
                "ancestor::w:p[1]",
                namespaces=namespaces,
            )

            if not paragraph_ancestors:
                raise RuntimeError(
                    "PAGEREF-Feld liegt weder in einem Hyperlink "
                    "noch in einem Absatz: "
                    f"{bookmark_name}"
                )

            entry_container = paragraph_ancestors[0]

        # Alle sichtbaren Texte vor dem PAGEREF-Feld ergeben:
        # Kapitelnummer + Überschrift.
        text_parts = []

        for element in entry_container.iter():
            if element is instruction_element:
                break

            if (
                etree.QName(element).namespace == w_namespace
                and etree.QName(element).localname == "t"
                and element.text
            ):
                text_parts.append(element.text)

        display_text = "".join(text_parts).strip()

        # Führende Word-Kapitelnummer entfernen:
        #
        # 1 Einleitung             -> Einleitung
        # 2.3 Gesamtbewertung     -> Gesamtbewertung
        # 5.1 Untersuchung ...    -> Untersuchung ...
        title = re.sub(
            r"^\s*\d+(?:\.\d+)*"
            r"(?:[.)])?"
            r"\s*",
            "",
            display_text,
        ).strip()

        if not title:
            raise RuntimeError(
                "Aus dem TOC-Eintrag konnte keine Überschrift "
                f"ermittelt werden: {bookmark_name}; "
                f"Anzeigetext={display_text!r}"
            )

        entries.append(
            {
                "bookmark": bookmark_name,
                "title": title,
                "display_text": display_text,
            }
        )

    if not entries:
        raise RuntimeError(
            "Im DOCX wurden keine PAGEREF-Einträge gefunden."
        )

    return entries

def find_pageref_pages_in_pdf(
    pdf_path: str,
    pageref_entries: list[dict],
    first_content_page: int = 3,
) -> dict[str, int]:
    """
    Ermittelt die PDF-Seite jedes sichtbaren TOC-Eintrags.

    Suchstufen:

    1. Normalisierter Textvergleich
    2. Kompakter Vergleich ohne Leerzeichen, Trennstriche
       und Satzzeichen

    Die TOC-Einträge werden in ihrer natürlichen Reihenfolge
    verarbeitet. Eine folgende Überschrift darf nur auf derselben
    oder einer späteren Seite liegen.
    """

    pdf_path = os.path.abspath(pdf_path)

    if not os.path.isfile(pdf_path):
        raise FileNotFoundError(
            f"PDF-Datei nicht gefunden: {pdf_path}"
        )

    reader = PdfReader(pdf_path)

    page_texts = []

    for page_number, page in enumerate(
        reader.pages,
        start=1,
    ):
        original_text = page.extract_text() or ""

        page_texts.append(
            {
                "page_number": page_number,
                "original_text": original_text,
                "normalized_text": normalize_match_text(
                    original_text
                ),
                "compact_text": normalize_compact_match_text(
                    original_text
                ),
            }
        )

    result = {}
    unresolved_entries = []

    # Die erste Überschrift darf frühestens auf der ersten
    # eigentlichen Inhaltsseite liegen.
    minimum_page = first_content_page

    for entry in pageref_entries:
        bookmark_name = entry["bookmark"]
        title = entry["title"]

        normalized_title = normalize_match_text(title)
        compact_title = normalize_compact_match_text(title)

        selected_page = None
        selected_method = None

        # -----------------------------------------------------
        # Suchstufe 1:
        # Normalisierter, aber weiterhin lesbarer Volltext
        # -----------------------------------------------------
        for page_data in page_texts:
            page_number = page_data["page_number"]

            if page_number < minimum_page:
                continue

            if (
                normalized_title
                and normalized_title
                in page_data["normalized_text"]
            ):
                selected_page = page_number
                selected_method = "normalisiert"
                break

        # -----------------------------------------------------
        # Suchstufe 2:
        # Kompakte Suche ohne Leerzeichen und Bindestriche
        #
        # Beispiele:
        # W ÄRMEDÄMMUNG -> wärmedämmung
        # PV-\nANLAGE   -> pvanlage
        # -----------------------------------------------------
        if selected_page is None:
            for page_data in page_texts:
                page_number = page_data["page_number"]

                if page_number < minimum_page:
                    continue

                if (
                    compact_title
                    and compact_title
                    in page_data["compact_text"]
                ):
                    selected_page = page_number
                    selected_method = "kompakt"
                    break

        if selected_page is None:
            unresolved_entries.append(entry)

            print(
                f"TOC-Seite nicht gefunden: "
                f"{entry['display_text']}"
            )

            continue

        result[bookmark_name] = selected_page

        # Die nächste Überschrift darf auf derselben Seite stehen,
        # aber nicht auf einer davorliegenden Seite.
        minimum_page = selected_page

        print(
            f"TOC-Seite ermittelt: "
            f"{entry['display_text']} "
            f"-> {selected_page} "
            f"[{selected_method}]"
        )

    if unresolved_entries:
        formatted_missing = "\n".join(
            (
                f"- {entry['bookmark']}: "
                f"{entry['display_text']!r}"
            )
            for entry in unresolved_entries
        )

        raise RuntimeError(
            "Nicht alle TOC-Einträge wurden im PDF gefunden:\n"
            f"{formatted_missing}"
        )

    return result

def patch_pageref_results(
    docx_path: str,
    bookmark_page_numbers: dict[str, int],
) -> None:
    """
    Aktualisiert ausschließlich die sichtbaren Ergebnisse der
    vorhandenen PAGEREF-Felder im Word-Inhaltsverzeichnis.

    Die Funktion unterstützt:
    - PAGEREF-Felder innerhalb von w:hyperlink
    - verschachtelte Word-Felder
    - das äußere TOC-Feld
    - mehrere PAGEREF-Felder im Dokument

    Alle anderen DOCX-Bestandteile bleiben unverändert.
    """

    docx_path = os.path.abspath(docx_path)

    if not os.path.isfile(docx_path):
        raise FileNotFoundError(
            f"DOCX-Datei nicht gefunden: {docx_path}"
        )

    namespaces = {
        "w": (
            "http://schemas.openxmlformats.org/"
            "wordprocessingml/2006/main"
        )
    }

    w_namespace = namespaces["w"]

    temporary_docx = docx_path + ".toc-update.tmp"

    updated_fields = 0
    found_pageref_fields = 0
    unresolved_bookmarks = []
    fields_without_result = []

    with zipfile.ZipFile(
        docx_path,
        "r",
    ) as source_archive:

        document_xml = source_archive.read(
            "word/document.xml"
        )

        xml_parser = etree.XMLParser(
            remove_blank_text=False,
            resolve_entities=False,
        )

        root = etree.fromstring(
            document_xml,
            parser=xml_parser,
        )

        paragraphs = root.xpath(
            "//w:p[.//w:instrText[contains("
            "translate(., 'pageref', 'PAGEREF'), "
            "'PAGEREF'"
            ")]]",
            namespaces=namespaces,
        )

        for paragraph in paragraphs:
            # Stack für verschachtelte Word-Felder.
            field_stack = []

            # iter() durchläuft auch Runs innerhalb von w:hyperlink.
            for element in paragraph.iter():
                local_name = etree.QName(element).localname
                namespace = etree.QName(element).namespace

                if namespace != w_namespace:
                    continue

                if local_name == "fldChar":
                    field_type = element.get(
                        f"{{{w_namespace}}}fldCharType"
                    )

                    if field_type == "begin":
                        field_stack.append(
                            {
                                "instruction_parts": [],
                                "result_text_elements": [],
                                "phase": "instruction",
                            }
                        )

                    elif field_type == "separate":
                        if field_stack:
                            field_stack[-1]["phase"] = "result"

                    elif field_type == "end":
                        if not field_stack:
                            continue

                        completed_field = field_stack.pop()

                        instruction = "".join(
                            completed_field["instruction_parts"]
                        ).strip()

                        match = re.search(
                            r"\bPAGEREF\s+([^\s\\]+)",
                            instruction,
                            flags=re.IGNORECASE,
                        )

                        if not match:
                            continue

                        found_pageref_fields += 1

                        bookmark_name = match.group(1)

                        if bookmark_name not in bookmark_page_numbers:
                            unresolved_bookmarks.append(
                                bookmark_name
                            )
                            continue

                        result_text_elements = completed_field[
                            "result_text_elements"
                        ]

                        if not result_text_elements:
                            fields_without_result.append(
                                bookmark_name
                            )
                            continue

                        page_number = str(
                            bookmark_page_numbers[bookmark_name]
                        )

                        # Nur das erste sichtbare Ergebnis erhält
                        # die neue Seitenzahl.
                        result_text_elements[0].text = page_number

                        # Eventuell vorhandene weitere Textfragmente
                        # desselben Feldergebnisses leeren.
                        for text_element in result_text_elements[1:]:
                            text_element.text = ""

                        updated_fields += 1

                elif local_name == "instrText":
                    if (
                        field_stack
                        and field_stack[-1]["phase"]
                        == "instruction"
                    ):
                        field_stack[-1][
                            "instruction_parts"
                        ].append(
                            element.text or ""
                        )

                elif local_name == "t":
                    if (
                        field_stack
                        and field_stack[-1]["phase"]
                        == "result"
                    ):
                        field_stack[-1][
                            "result_text_elements"
                        ].append(element)

        print(
            f"PAGEREF-Felder im DOCX gefunden: "
            f"{found_pageref_fields}"
        )

        print(
            f"PAGEREF-Felder aktualisiert: "
            f"{updated_fields}"
        )

        if unresolved_bookmarks:
            unresolved_string = ", ".join(
                sorted(set(unresolved_bookmarks))
            )

            raise RuntimeError(
                "Für folgende PAGEREF-Bookmarks wurde keine "
                f"Seite ermittelt: {unresolved_string}"
            )

        if fields_without_result:
            missing_result_string = ", ".join(
                sorted(set(fields_without_result))
            )

            raise RuntimeError(
                "Folgende PAGEREF-Felder besitzen kein sichtbares "
                f"Ergebnis: {missing_result_string}"
            )

        if found_pageref_fields == 0:
            raise RuntimeError(
                "Im DOCX wurden keine PAGEREF-Felder gefunden."
            )

        if updated_fields != found_pageref_fields:
            raise RuntimeError(
                "Nicht alle PAGEREF-Felder konnten aktualisiert werden. "
                f"Gefunden: {found_pageref_fields}, "
                f"aktualisiert: {updated_fields}."
            )

        updated_document_xml = etree.tostring(
            root,
            xml_declaration=True,
            encoding="UTF-8",
            standalone=True,
        )

        try:
            with zipfile.ZipFile(
                temporary_docx,
                "w",
                compression=zipfile.ZIP_DEFLATED,
            ) as target_archive:

                for source_item in source_archive.infolist():
                    item_data = source_archive.read(
                        source_item.filename
                    )

                    if (
                        source_item.filename
                        == "word/document.xml"
                    ):
                        item_data = updated_document_xml

                    target_archive.writestr(
                        source_item,
                        item_data,
                    )

        except Exception:
            if os.path.exists(temporary_docx):
                os.remove(temporary_docx)
            raise

    os.replace(
        temporary_docx,
        docx_path,
    )

    print(
        f"{updated_fields} TOC-Seitenzahlen "
        "erfolgreich in das DOCX geschrieben."
    )

def update_toc_page_numbers_without_roundtrip(
    docx_path: str,
    first_content_page: int = 3,
) -> str:
    """
    Aktualisiert ausschließlich die sichtbaren Seitenzahlen des
    vorhandenen Word-Inhaltsverzeichnisses.

    Ablauf:
    1. Sichtbare TOC-Einträge und PAGEREF-Felder lesen.
    2. DOCX mit LibreOffice nur als temporäre PDF rendern.
    3. Seiten der Überschriften im PDF ermitteln.
    4. Nur die PAGEREF-Ergebnisse im DOCX aktualisieren.

    LibreOffice speichert das DOCX zu keinem Zeitpunkt.
    """

    docx_path = os.path.abspath(docx_path)

    if not os.path.isfile(docx_path):
        raise FileNotFoundError(
            f"DOCX-Datei nicht gefunden: {docx_path}"
        )

    work_directory = tempfile.mkdtemp(
        prefix="toc_page_calculation_"
    )

    try:
        print(
            "Ermittle TOC-Einträge aus dem DOCX."
        )

        pageref_entries = extract_pageref_entries(
            docx_path
        )

        print(
            f"{len(pageref_entries)} TOC-Einträge gefunden."
        )

        for entry in pageref_entries:
            print(
                f"  {entry['bookmark']}: "
                f"{entry['display_text']}"
            )

        print(
            "Rendere temporäre PDF-Datei zur "
            "Seitenermittlung."
        )

        pdf_path = render_docx_to_pdf_with_libreoffice(
            docx_path=docx_path,
            output_directory=work_directory,
        )

        bookmark_pages = find_pageref_pages_in_pdf(
            pdf_path=pdf_path,
            pageref_entries=pageref_entries,
            first_content_page=first_content_page,
        )

        print(
            f"Für {len(bookmark_pages)} TOC-Einträge "
            "wurden Seiten ermittelt."
        )

        if len(bookmark_pages) != len(pageref_entries):
            raise RuntimeError(
                "Die Anzahl der ermittelten Seiten stimmt nicht "
                "mit der Anzahl der TOC-Einträge überein. "
                f"TOC-Einträge: {len(pageref_entries)}, "
                f"Seiten: {len(bookmark_pages)}."
            )

        patch_pageref_results(
            docx_path=docx_path,
            bookmark_page_numbers=bookmark_pages,
        )

        print(
            f"TOC-Seitenzahlen erfolgreich aktualisiert: "
            f"{docx_path}"
        )

        return docx_path

    finally:
        shutil.rmtree(
            work_directory,
            ignore_errors=True,
        )

def find_libreoffice_executable() -> str:
    configured_path = os.getenv("LIBREOFFICE_PATH")

    candidates = [
        configured_path,
        shutil.which("soffice"),
        shutil.which("libreoffice"),
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        "/usr/bin/libreoffice",
        "/usr/bin/soffice",
    ]

    for candidate in candidates:
        if not candidate:
            continue

        candidate_path = Path(candidate)

        if candidate_path.is_file():
            return str(candidate_path.resolve())

    raise FileNotFoundError(
        "LibreOffice wurde nicht gefunden. "
        "Installiere LibreOffice oder setze LIBREOFFICE_PATH."
    )

def find_uno_python() -> str:
    configured_path = os.getenv("UNO_PYTHON_PATH")

    candidates = [
        configured_path,
        r"C:\Program Files\LibreOffice\program\python.exe",
        r"C:\Program Files (x86)\LibreOffice\program\python.exe",
        "/usr/bin/python3",
        sys.executable,
    ]

    for candidate in candidates:
        if not candidate:
            continue

        candidate_path = Path(candidate)

        if not candidate_path.is_file():
            continue

        test_result = subprocess.run(
            [
                str(candidate_path),
                "-c",
                "import uno; print('UNO_OK')",
            ],
            capture_output=True,
            text=True,
            timeout=20,
            check=False,
        )

        if (
            test_result.returncode == 0
            and "UNO_OK" in test_result.stdout
        ):
            return str(candidate_path.resolve())

    raise RuntimeError(
        "Kein Python-Interpreter mit LibreOffice-UNO-Unterstützung "
        "gefunden. Setze UNO_PYTHON_PATH auf einen Interpreter, "
        "bei dem 'import uno' funktioniert."
    )

def update_toc_with_libreoffice(
    docx_path: str,
    timeout_seconds: int = 120,
) -> str:
    """
    Öffnet das DOCX headless in LibreOffice, aktualisiert
    alle Dokumentindizes und speichert es wieder.

    Pro Aufruf wird ein eigenes LibreOffice-Benutzerprofil
    verwendet, damit parallele oder hängende lokale Sitzungen
    den Job nicht blockieren.
    """

    docx_path = os.path.abspath(docx_path)

    if not os.path.isfile(docx_path):
        raise FileNotFoundError(
            f"DOCX-Datei nicht gefunden: {docx_path}"
        )

    if not UNO_SCRIPT_PATH.is_file():
        raise FileNotFoundError(
            f"UNO-Skript nicht gefunden: {UNO_SCRIPT_PATH}"
        )

    soffice_path = find_libreoffice_executable()
    uno_python_path = find_uno_python()
    port = find_free_local_port()

    profile_directory = tempfile.mkdtemp(
        prefix="libreoffice_profile_"
    )

    profile_url = (
        Path(profile_directory)
        .resolve()
        .as_uri()
    )

    libreoffice_command = [
        soffice_path,
        "--headless",
        "--nologo",
        "--nodefault",
        "--nofirststartwizard",
        "--norestore",
        f"-env:UserInstallation={profile_url}",
        (
            f"--accept=socket,host=127.0.0.1,port={port};"
            "urp;StarOffice.ServiceManager"
        ),
    ]

    office_process = None

    try:
        print(
            f"Starte LibreOffice für TOC-Aktualisierung "
            f"auf Port {port}."
        )

        office_process = subprocess.Popen(
            libreoffice_command,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
        )

        uno_command = [
            uno_python_path,
            str(UNO_SCRIPT_PATH),
            docx_path,
            str(port),
        ]

        result = subprocess.run(
            uno_command,
            capture_output=True,
            text=True,
            timeout=timeout_seconds,
            check=False,
        )

        if result.stdout:
            print(result.stdout.strip())

        if result.returncode != 0:
            raise RuntimeError(
                "LibreOffice-UNO-Aktualisierung fehlgeschlagen.\n"
                f"Exitcode: {result.returncode}\n"
                f"STDOUT:\n{result.stdout}\n"
                f"STDERR:\n{result.stderr}"
            )

        if not os.path.isfile(docx_path):
            raise FileNotFoundError(
                "Nach der LibreOffice-Verarbeitung fehlt die Datei: "
                f"{docx_path}"
            )

        print(
            f"Inhaltsverzeichnis wurde aktualisiert: "
            f"{docx_path}"
        )

        return docx_path

    finally:
        if office_process is not None:
            office_process.terminate()

            try:
                office_process.wait(timeout=10)
            except subprocess.TimeoutExpired:
                office_process.kill()
                office_process.wait(timeout=5)

        shutil.rmtree(
            profile_directory,
            ignore_errors=True,
        )

def upload_file_to_sharepoint_folder(
    ctx,
    local_file_path: str,
    folder_server_relative_url: str,
    target_file_name: str | None = None,
    status_list_title: str = None,
    status_item_id: int = None,
    status_field_internal_name: str = "field_1",
    status_value: str = "Gutachten erstellt",
):
    if not os.path.isfile(local_file_path):
        raise FileNotFoundError(f"File not found: {local_file_path}")

    file_name = target_file_name or os.path.basename(local_file_path)

    with open(local_file_path, "rb") as f:
        file_content = f.read()

    target_folder = ctx.web.get_folder_by_server_relative_url(folder_server_relative_url)
    uploaded_file = target_folder.upload_file(file_name, file_content).execute_query()

    if uploaded_file and status_list_title and status_item_id is not None:
        sp_list = ctx.web.lists.get_by_title(status_list_title)
        sp_item = sp_list.items.get_by_id(status_item_id)
        sp_item.set_property(status_field_internal_name, status_value)
        sp_item.update()
        ctx.execute_query()

    return uploaded_file

def extract_from_sharepoint(gutachten_texte_filtered_items, texts) -> Dict[str, List]:
    paragraphs_begutachtungsziel                    = []
    paragraphs_dachproben_sichtbegutachtung         = []
    paragraphs_dachflaechen_aufmass                 = []
    paragraphs_dachprobenanalyse                    = []
    paragraphs_dachsichtbegutachtung                = []
    paragraphs_dachsichtbegutachtung                = []
    paragraphs_gesamtbewertung                      = []
    paragraphs_dachaufbau_analyse                   = []
    paragraphs_u_wert_berechnung                    = []
    paragraphs_effizienz_beurteilung                = []
    paragraphs_energiebilanz_empfehlungen           = []
    paragraphs_statikpruefung                       = []
    paragraphs_dachkonstruktion_ueberpruefung       = []
    paragraphs_waermedaemmung_druckfestigkeit       = []
    paragraphs_untersuchung_waermedaemmung          = []
    paragraphs_feststellung_schwachstellen          = []
    paragraphs_empfehlungen_belastbarkeit           = []
    paragraphs_gebaeudegeometrie                    = []
    paragraphs_brandschutz                          = []
    paragraphs_gesamtbewertung_ds_pv                = []
    paragraphs_handlungsempfehlungen                = []
    tables_by_chapter                               = defaultdict(list)
    
    #Ergebnis durchlaufen
    for item in gutachten_texte_filtered_items:
        props       = item.properties
        level       = props.get("field_2", "")
        data        = props.get("field_4", "")
        data_type   = props.get("field_3", "text").lower()
        chapter = props.get("field_5", "")

        #print(f"Level: {level}, Data: {data}, Type: {data_type}, Chapter: {chapter}")

        # Build the same dict‐shape that your JSON loader used before.
        if data_type == "image":
            para_dict = {"level": 0, "type": "image", "src": data}
            # append to the correct paragraph list (same as before)
        elif data_type == "table":
            # parse table rows and store under this chapter
            para_dict = extract_table_rows_from_data_field(data)
            #print("WICHTIG")
            print(para_dict)
            #print("WICHTIG")
        else:
            para_dict = {"level": level, "type": "text", "text": data}

        # Decide which of the 15 lists to append to, based on chapter.
        # (Adjust these string‐matches to exactly whatever chapter holds.)

        if chapter == "Begutachtungsziel":
            paragraphs_begutachtungsziel.append(para_dict)
        elif chapter == "Dachproben & Sichtbegutachtung":
            paragraphs_dachproben_sichtbegutachtung.append(para_dict)
        elif chapter == "Dachflächen Aufmaß":
            paragraphs_dachflaechen_aufmass.append(para_dict)
        elif chapter == "Dachprobenanalyse":
            paragraphs_dachprobenanalyse.append(para_dict)
        elif chapter == "Dachsichtbegutachtung":
            paragraphs_dachsichtbegutachtung.append(para_dict)
        elif chapter == "Gesamtbewertung":
            paragraphs_gesamtbewertung.append(para_dict)
        elif chapter == "Dachaufbau & Analyse":
            paragraphs_dachaufbau_analyse.append(para_dict)
        elif chapter == "U-Wert Berechnung":
            paragraphs_u_wert_berechnung.append(para_dict)
        elif chapter == "Effizienz Beurteilung":
            paragraphs_effizienz_beurteilung.append(para_dict)
        elif chapter == "Energiebilanz & Empfehlungen":
            paragraphs_energiebilanz_empfehlungen.append(para_dict)
        elif chapter == "Statikprüfung":
            paragraphs_statikpruefung.append(para_dict)
        elif chapter == "Dachkonstruktion Überprüfung":
            paragraphs_dachkonstruktion_ueberpruefung.append(para_dict)
        elif chapter == "Wärmedämmung & Druckfestigkeit":
            paragraphs_waermedaemmung_druckfestigkeit.append(para_dict)
        elif chapter == "Untersuchung Wärmedämmung":
            paragraphs_untersuchung_waermedaemmung.append(para_dict)
        elif chapter == "Feststellung Schwachstellen":
            paragraphs_feststellung_schwachstellen.append(para_dict)
        elif chapter == "Empfehlungen Belastbarkeit":
            paragraphs_empfehlungen_belastbarkeit.append(para_dict)
        elif chapter == "Gebäudegeometrie":
            paragraphs_gebaeudegeometrie.append(para_dict)
        elif chapter == "Brandschutz":
            paragraphs_brandschutz.append(para_dict)
        elif chapter == "Gesamtbewertung DS PV":
            paragraphs_gesamtbewertung_ds_pv.append(para_dict)
        elif chapter == "Handlungsempfehlungen":
            paragraphs_handlungsempfehlungen.append(para_dict)

        else:
            # If there are any extra levels (or missing ones), you can either skip or log them.
            # print(f"Warning: Unmapped Level '{level}'. Skipping.")
            continue

    return {
        "paragraphs_begutachtungsziel":                    paragraphs_begutachtungsziel,
        "paragraphs_dachproben_sichtbegutachtung":         paragraphs_dachproben_sichtbegutachtung,
        "paragraphs_dachflaechen_aufmass":                 paragraphs_dachflaechen_aufmass,
        "paragraphs_dachprobenanalyse":                    paragraphs_dachprobenanalyse,
        "paragraphs_dachsichtbegutachtung":                paragraphs_dachsichtbegutachtung,
        "paragraphs_gesamtbewertung":                      paragraphs_gesamtbewertung,
        "paragraphs_dachaufbau_analyse":                   paragraphs_dachaufbau_analyse,
        "paragraphs_u_wert_berechnung":                    paragraphs_u_wert_berechnung,
        "paragraphs_effizienz_beurteilung":                paragraphs_effizienz_beurteilung,
        "paragraphs_energiebilanz_empfehlungen":           paragraphs_energiebilanz_empfehlungen,
        "paragraphs_statikpruefung":                       paragraphs_statikpruefung,
        "paragraphs_dachkonstruktion_ueberpruefung":       paragraphs_dachkonstruktion_ueberpruefung,
        "paragraphs_waermedaemmung_druckfestigkeit":       paragraphs_waermedaemmung_druckfestigkeit,
        "paragraphs_untersuchung_waermedaemmung":          paragraphs_untersuchung_waermedaemmung,
        "paragraphs_feststellung_schwachstellen":          paragraphs_feststellung_schwachstellen,
        "paragraphs_empfehlungen_belastbarkeit":           paragraphs_empfehlungen_belastbarkeit,
        "paragraphs_gebaeudegeometrie":                    paragraphs_gebaeudegeometrie,
        "paragraphs_brandschutz":                          paragraphs_brandschutz,
        "paragraphs_gesamtbewertung_ds_pv":                paragraphs_gesamtbewertung_ds_pv,
        "paragraphs_handlungsempfehlungen":                paragraphs_handlungsempfehlungen,
        "texts":                                           texts
    }

def extract_table_rows_from_data_field(raw_data: str) -> List[Dict[str, List]]:
    rows: List[Dict[str, List]] = []
    if not raw_data:
        return rows

    # Split exactly on the delimiter. Do NOT remove interior newlines in tokens.
    parts = [p.strip() for p in re.split(r"_\|_", raw_data) if p is not None]
    # Remove purely-empty tokens
    parts = [p for p in parts if p != ""]

    # Group into triplets (nr, description, image_src)
    i = 0
    while i < len(parts):
        nr = parts[i] if i < len(parts) else ""
        description = parts[i + 1] if (i + 1) < len(parts) else ""
        image_src = parts[i + 2] if (i + 2) < len(parts) else ""

        rows.append({
            "cols": [nr, description, image_src]
        })
        i += 3
    return rows

def build_report(
    data: dict,
    template_path: str,
    output_path: str,
    export_pdf: bool,
    _temp_img_dir: str,
    ctx
):
    tpl = DocxTemplate(template_path)

    # texts fallback (avoid KeyError used later with **texts)
    texts = data.get("texts", {})

    # Helper to normalize newlines in all string fields of a paragraph
    def normalize_newlines(paragraphs: List[Dict[str, Any]]):
        out = []
        for p in (paragraphs or []):
            if not isinstance(p, dict):
                out.append(p)
                continue
            p2 = {}
            for k, v in p.items():
                if isinstance(v, str):
                    # convert escaped newline tokens to real newline
                    p2[k] = v.replace("\\n", "\n").replace(r"\n", "/n")
                else:
                    p2[k] = v
            out.append(p2)
        return out

    # make one temp folder per build
    _image_cache = {}

    def make_image(img_src: str, *, force_width_cm: float = None, max_height_cm: float = np.inf, _temp_img_dir: str):
        """
        Fetch image from SharePoint (using ctx.web...), cache as PIL,
        save locally to temp dir and return an InlineImage sized appropriately.
        """
        if not img_src:
            return None

        # Build server relative path (adjust if your SharePoint path differs)
        server_url = "/Freigegebene%20Dokumente/Dev_Vincent/Images/"
        server_path = server_url + img_src
        # print(f"Fetching image from {server_path}")

        # 2) Download once into PIL
        if server_path not in _image_cache:
            buffer = BytesIO()
            try:
                ctx.web \
                   .get_file_by_server_relative_url(server_path) \
                   .download(buffer) \
                   .execute_query()
            except Exception as e:
                # don't crash the whole build — return None so caller can provide placeholder text
                print(f"Warning: Cannot fetch {server_path}: {e}")
                return None

            buffer.seek(0)
            try:
                pil_img = Image.open(buffer).convert("RGB")
            except Exception as e:
                print(f"Warning: failed to open image buffer for {server_path}: {e}")
                return None
            _image_cache[server_path] = pil_img
        else:
            pil_img = _image_cache[server_path]

        # Save locally under temp dir
        local_filename = os.path.basename(img_src)
        # ensure unique file name in temp dir to avoid collisions
        local_path = os.path.join(_temp_img_dir, local_filename)
        try:
            pil_img.save(local_path, format="JPEG", quality=90)
        except Exception:
            # fallback to PNG if JPEG fails
            pil_img.save(local_path, format="PNG")

        # Compute dimensions in mm to pass to InlineImage
        width_px, height_px = pil_img.size
        aspect = width_px / height_px if height_px != 0 else 1.0

        if force_width_cm is not None:
            w_mm = float(force_width_cm) * 10.0
            h_mm = w_mm / aspect
            if (h_mm / 10.0) > max_height_cm:
                h_mm = float(max_height_cm) * 10.0
                w_mm = h_mm * aspect
        else:
            # convert px -> mm (empirical). Then clamp to reasonable limits.
            PX_TO_MM = 1.759047619047619
            w_mm = float(np.round((width_px * PX_TO_MM), 2))
            h_mm = float(np.round((height_px * PX_TO_MM), 2))
            # clamp to page-friendly maxima (in mm)
            MAX_W_MM = 173.6
            MAX_H_MM = 200.0
            if w_mm > MAX_W_MM:
                w_mm = MAX_W_MM
                h_mm = float(np.round(w_mm / aspect, 2))
            if h_mm > MAX_H_MM:
                h_mm = MAX_H_MM
                w_mm = float(np.round(h_mm * aspect, 2))

        # docx InlineImage expects Length; pass Mm
        return InlineImage(tpl, local_path, width=Mm(w_mm), height=Mm(h_mm))

    # Normalize paragraphs (existing keys in your original code)
    paragraphs_begutachtungsziel = normalize_newlines(data.get("paragraphs_begutachtungsziel", []))
    paragraphs_dachproben_sichtbegutachtung = normalize_newlines(data.get("paragraphs_dachproben_sichtbegutachtung", []))
    paragraphs_dachflaechen_aufmass = normalize_newlines(data.get("paragraphs_dachflaechen_aufmass", []))
    paragraphs_dachprobenanalyse = data.get("paragraphs_dachprobenanalyse", [])[0]
    paragraphs_dachsichtbegutachtung = data.get("paragraphs_dachsichtbegutachtung", [])[0]
    paragraphs_gesamtbewertung = normalize_newlines(data.get("paragraphs_gesamtbewertung", []))
    paragraphs_dachaufbau_analyse = normalize_newlines(data.get("paragraphs_dachaufbau_analyse", []))
    paragraphs_u_wert_berechnung = normalize_newlines(data.get("paragraphs_u_wert_berechnung", []))
    paragraphs_effizienz_beurteilung = normalize_newlines(data.get("paragraphs_effizienz_beurteilung", []))
    paragraphs_energiebilanz_empfehlungen = normalize_newlines(data.get("paragraphs_energiebilanz_empfehlungen", []))
    paragraphs_statikpruefung = normalize_newlines(data.get("paragraphs_statikpruefung", []))
    paragraphs_dachkonstruktion_ueberpruefung = normalize_newlines(data.get("paragraphs_dachkonstruktion_ueberpruefung", []))
    paragraphs_waermedaemmung_druckfestigkeit = normalize_newlines(data.get("paragraphs_waermedaemmung_druckfestigkeit", []))
    paragraphs_untersuchung_waermedaemmung = normalize_newlines(data.get("paragraphs_untersuchung_waermedaemmung", []))
    paragraphs_feststellung_schwachstellen = normalize_newlines(data.get("paragraphs_feststellung_schwachstellen", []))
    paragraphs_empfehlungen_belastbarkeit = normalize_newlines(data.get("paragraphs_empfehlungen_belastbarkeit", []))
    paragraphs_gebaeudegeometrie = normalize_newlines(data.get("paragraphs_gebaeudegeometrie", []))
    paragraphs_brandschutz = normalize_newlines(data.get("paragraphs_brandschutz", []))
    paragraphs_gesamtbewertung_ds_pv = normalize_newlines(data.get("paragraphs_gesamtbewertung_ds_pv", []))
    paragraphs_handlungsempfehlungen = normalize_newlines(data.get("paragraphs_handlungsempfehlungen", []))

    # Build images inside generic paragraph lists if type == image
    all_paragraphs = (paragraphs_begutachtungsziel + paragraphs_dachproben_sichtbegutachtung +
                      paragraphs_dachflaechen_aufmass + paragraphs_dachprobenanalyse + paragraphs_dachsichtbegutachtung + paragraphs_gesamtbewertung +
                      paragraphs_dachaufbau_analyse + paragraphs_u_wert_berechnung +
                      paragraphs_effizienz_beurteilung + paragraphs_energiebilanz_empfehlungen +
                      paragraphs_statikpruefung + paragraphs_dachkonstruktion_ueberpruefung +
                      paragraphs_waermedaemmung_druckfestigkeit + paragraphs_untersuchung_waermedaemmung +
                      paragraphs_feststellung_schwachstellen + paragraphs_empfehlungen_belastbarkeit +
                      paragraphs_gebaeudegeometrie + paragraphs_brandschutz + paragraphs_gesamtbewertung_ds_pv + paragraphs_handlungsempfehlungen)

    for p in all_paragraphs:
        #print(f"Processing paragraph: {p}")
        if isinstance(p, dict):
            if p.get("type") == "image" and "src" in p:
                p["image"] = make_image(p.pop("src"), force_width_cm=16.0, max_height_cm=8.0, _temp_img_dir=_temp_img_dir) or "Bild nicht verfügbar"
            elif "cols" in p and isinstance(p["cols"], list) and len(p["cols"]) > 2:
                # Check if the last column is an image source
                image_src = p["cols"][-1]
                if isinstance(image_src, str) and image_src.lower().endswith(('.png', '.jpg', '.jpeg')):
                    p["cols"][-1] = make_image(image_src, force_width_cm=6.0, max_height_cm=8.0, _temp_img_dir=_temp_img_dir) or "Bild nicht verfügbar"

    # -----------------------------
    # New: process tables from data.get("tables_by_chapter")
    # Expect: data["tables_by_chapter"] = { chapter_name: [ {nr, description, image_src}, ... ], ... }
    # Convert image_src -> InlineImage (or text placeholder) and expose as context['tables']
    # -----------------------------

    # 3. Build context and render
    context = {
        "page_break": "\f",  # Word’s page-break marker
        "paragraphs_begutachtungsziel": paragraphs_begutachtungsziel,
        "paragraphs_dachproben_sichtbegutachtung": paragraphs_dachproben_sichtbegutachtung,
        "paragraphs_dachflaechen_aufmass": paragraphs_dachflaechen_aufmass,
        "paragraphs_dachprobenanalyse": paragraphs_dachprobenanalyse,
        "paragraphs_dachsichtbegutachtung": paragraphs_dachsichtbegutachtung,
        "paragraphs_gesamtbewertung": paragraphs_gesamtbewertung,
        "paragraphs_dachaufbau_analyse": paragraphs_dachaufbau_analyse,
        "paragraphs_u_wert_berechnung": paragraphs_u_wert_berechnung,
        "paragraphs_effizienz_beurteilung": paragraphs_effizienz_beurteilung,
        "paragraphs_energiebilanz_empfehlungen": paragraphs_energiebilanz_empfehlungen,
        "paragraphs_statikpruefung": paragraphs_statikpruefung,
        "paragraphs_dachkonstruktion_ueberpruefung": paragraphs_dachkonstruktion_ueberpruefung,
        "paragraphs_waermedaemmung_druckfestigkeit": paragraphs_waermedaemmung_druckfestigkeit,
        "paragraphs_untersuchung_waermedaemmung": paragraphs_untersuchung_waermedaemmung,
        "paragraphs_feststellung_schwachstellen": paragraphs_feststellung_schwachstellen,
        "paragraphs_empfehlungen_belastbarkeit": paragraphs_empfehlungen_belastbarkeit,
        "paragraphs_gebaeudegeometrie": paragraphs_gebaeudegeometrie,
        "paragraphs_brandschutz": paragraphs_brandschutz,
        "paragraphs_gesamtbewertung_ds_pv": paragraphs_gesamtbewertung_ds_pv,
        "paragraphs_handlungsempfehlungen": paragraphs_handlungsempfehlungen,
        **texts
    }
    if "normen_und_richtlinien" in context:
        context["normen_lines"] = context["normen_und_richtlinien"].splitlines()

    #print(f"Context - paragraphs_dachprobenanalyse:\n{context['paragraphs_dachprobenanalyse']}")
    #exit()
    tpl.render(context)

    # now post-process the generated document to set "page break before"
    doc = tpl.docx  # underlying python-docx Document

    first = True
    for para in doc.paragraphs:
        text = para.text.strip()
        if text.startswith("2.3.1 "):
            if first:
                first = False
            else:
                para.paragraph_format.page_break_before = True

    tpl.save(output_path)

    print(
        f"DOCX aus Template erzeugt: {output_path}"
    )

    update_toc_page_numbers_without_roundtrip(
        docx_path=output_path,
        first_content_page=3,
    )

    print(
        f"TOC-Seitenzahlen aktualisiert: {output_path}"
    )

    if not export_pdf:
        return
    pdf_output = os.path.splitext(output_path)[0] + ".pdf"
    try:
        result = subprocess.run([
            "soffice", "--headless", "--convert-to", "pdf", "--outdir",
            os.path.dirname(os.path.abspath(output_path)) or ".", output_path
        ], check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        if os.path.exists(pdf_output):
            print(f"PDF exported: {pdf_output}")
        else:
            print("PDF export failed: Output file not found.")
    except Exception as e:
        print("PDF export failed. Please ensure LibreOffice is installed and 'soffice' is in your PATH.")
        print(f"Error: {e}")

def main():
    if not TEMPLATE_PATH.is_file():
        raise FileNotFoundError(
            f"Template nicht gefunden: {TEMPLATE_PATH}"
        )
    load_dotenv() # Load from .env file

    log_document_runtime_environment()

    SITE_URL = os.getenv("SITE_URL")
    TENANT_ID = os.getenv("TENANT_ID")
    CLIENT_ID = os.getenv("CLIENT_ID")
    THUMBPRINT = os.getenv("THUMBPRINT")
    cert_b64 = os.getenv("CERT_BASE64")
    cert_path_env = os.getenv("CERT_PATH")

    if cert_b64:
        # ✅ Azure: use Base64 env variable
        cert_file_path = "/tmp/cert.pem"
        with open(cert_file_path, "wb") as f:
            f.write(base64.b64decode(cert_b64))
        CERT_PATH = Path(cert_file_path)

    elif cert_path_env:
        # ✅ Local: use file path from .env
        CERT_PATH = Path(cert_path_env)

    else:
        raise ValueError("No certificate configuration found (CERT_BASE64 or CERT_PATH missing)")


    if not SITE_URL:
        raise ValueError("SITE_URL missing")

    missing = [k for k, v in {
        "SITE_URL": SITE_URL,
        "TENANT_ID": TENANT_ID,
        "CLIENT_ID": CLIENT_ID,
        "THUMBPRINT": THUMBPRINT,
        "CERT_PATH": CERT_PATH
    }.items() if not v]

    if missing:
        raise ValueError(f"Missing env variables: {missing}")

    # if not CERT_PATH.exists():
    #     CERT_PATH = r"C:\Users\Vincent\Desktop\NorproofDocPopCert.pem"
    #     #raise FileNotFoundError(f"Certificate not found: {CERT_PATH}")

    url = os.getenv("URL")
    ctx = ClientContext(base_url=url).with_client_certificate(
        tenant=TENANT_ID, 
        client_id=CLIENT_ID,
        thumbprint=THUMBPRINT,
        cert_path=CERT_PATH
    )

    # Liste "Dev_VC_Projekt" nur mit Status = "Offen"
    project_list = ctx.web.lists.get_by_title("Tab_Projekte")

    # Erstelle eine CAML-Abfrage, die nur Einträge mit Status "Dachprobenfotos erhalten" zurückliefert
    caml = CamlQuery()
    caml.ViewXml = """
    <View>
    <Query>
        <Where>
        <Eq>
            <FieldRef Name='Status' />
            <Value Type='Text'>Dachprobenfotos erhalten</Value>
        </Eq>
        </Where>
    </Query>
    </View>
    """

    # Führe die Abfrage aus
    filtered_projects = project_list.get_items(caml).execute_query()

    from datetime import datetime

    caml = CamlQuery()
    for project in filtered_projects:
        props             = project.properties
        #print(props)
        ID = props["ID"]
        project_id  = props["Projekt_ID"]
        # Format stand from "2024-06-07T07:00:00Z" to "07.06.2024"
        stand = datetime.now().strftime("%d.%m.%Y")
        project_titel  = props["Projekt_ID"]
        project_strasse  = props["Stra_x00df_e"]
        project_hausnummer  = props["Hausnummer"]
        project_plz  = props["PLZ"]
        project_ort  = props["Stadt"]
        auftraggeber_id  = props["Title"]
        _temp_img_dir = tempfile.mkdtemp(prefix="report_images_")

        # Füge die Abfrage für den Auftraggeber hinzu
        customer_list = ctx.web.lists.get_by_title("Dev_VC_Kunde")
        caml.ViewXml = f"""
        <View>
        <Query>
        <Where>
            <Eq>
            <FieldRef Name='Title' />
            <Value Type='Text'>{auftraggeber_id}</Value>
            </Eq>
        </Where>
        </Query>
        </View>
        """
        # Führe die Abfrage aus
        filtered_customer = customer_list.get_items(caml).execute_query()
        if not filtered_customer:
            print(f"Kein Kunde mit Title='{auftraggeber_id}' gefunden!")
            continue
        props = filtered_customer[0].properties
        customer_name       = props["Title"]
        customer_strasse    = props["field_4"]
        customer_hausnummer = props["field_5"]
        customer_plz        = props["field_3"]
        customer_ort        = props["field_2"]
        
        texts = {
            "stand": stand,
            "projektnummer": project_id,
            "projekt_titel": project_titel,
            "projekt_strasse": project_strasse + " " + str(project_hausnummer),
            "projekt_plz": project_plz + " " + project_ort,
            "auftraggeber_name": customer_name,
            "auftraggeber_strasse": customer_strasse + " " + str(customer_hausnummer),
            "auftraggeber_plz": str(customer_plz) + " " + customer_ort,
            "aufstellungs_ort": "58089 Hagen",
            "aufstellungs_datum": "14.06.2024"
        }

        gutachten_texte_list = ctx.web.lists.get_by_title("Dev_VC_Gutachten-Texte")
        caml.ViewXml = f"""
        <View>
        <Query>
            <Where>
            <Eq>
                <FieldRef Name='Title' /> 
                <Value Type='Text'>{project_id}</Value>
            </Eq>
            </Where>
        </Query>
        </View>
        """

        #Gefilterte Items abfragen
        gutachten_texte_filtered_items = gutachten_texte_list.get_items(caml).execute_query()
        dict_result = extract_from_sharepoint(gutachten_texte_filtered_items, texts)

        with tempfile.NamedTemporaryFile(prefix=f"{project_id}_", suffix=".docx", delete=False) as f:
            output_docx = f.name
            
        input_docx = os.path.join(os.path.dirname(__file__), output_docx)
        build_report(
        data=dict_result,
        template_path=str(TEMPLATE_PATH),
        output_path=output_docx,
        export_pdf=False,
        _temp_img_dir=_temp_img_dir,
        ctx=ctx
        )
        uploaded = upload_file_to_sharepoint_folder(
            ctx=ctx,
            local_file_path=output_docx,
            folder_server_relative_url=f"/sites/Norproof/Projektdaten/{project_id}",
            target_file_name="Gutachten.docx",
            status_list_title="Tab_Projekte",
            status_item_id=ID,
            status_field_internal_name="Status",
            status_value="Gutachten erstellt",
        )
        shutil.rmtree(_temp_img_dir, ignore_errors=True)

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"Job failed: {e}", file=sys.stderr)
        raise  # oder sys.exit(1)

# if __name__ == "__main__":
#     test_document = (
#         BASE_DIR
#         / "Output-TOC-Test-2.docx"
#     )

#     update_toc_page_numbers_without_roundtrip(
#         docx_path=str(test_document),
#         first_content_page=3,
#     )

#     print(
#         f"Test abgeschlossen: {test_document}"
#     )