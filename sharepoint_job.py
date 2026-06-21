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
import sys
#import win32com.client as win32
#from win32com.client import constants
import re
from typing import List, Dict
from docxtpl import InlineImage
from docx.shared import Mm
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm, Cm
from PIL import Image
import numpy as np
import subprocess
import os
from io import BytesIO
import tempfile
import shutil
from typing import List, Dict, Any
from dotenv import load_dotenv
from office365.sharepoint.client_context import ClientContext
import os
from office365.sharepoint.listitems.caml.query import CamlQuery
from pprint import pprint
from typing import List, Dict
from pathlib import Path
from collections import defaultdict

def upload_file_to_sharepoint_folder(
    ctx,
    local_file_path: str,
    folder_server_relative_url: str,
    status_list_title: str = None,
    status_item_id: int = None,
    status_field_internal_name: str = "field_1",  # likely your "Status" choice column
    status_value: str = "Gutachten erstellt",
):
    if not os.path.isfile(local_file_path):
        raise FileNotFoundError(f"File not found: {local_file_path}")

    file_name = os.path.basename(local_file_path)
    with open(local_file_path, "rb") as f:
        file_content = f.read()

    # Upload file
    target_folder = ctx.web.get_folder_by_server_relative_url(folder_server_relative_url)
    uploaded_file = target_folder.upload_file(file_name, file_content).execute_query()
    #print(f"Uploaded: {uploaded_file.serverRelativeUrl}")

    # Update SharePoint list item only if upload succeeded and row info was provided
    if uploaded_file and status_list_title and status_item_id is not None:
        sp_list = ctx.web.lists.get_by_title(status_list_title)
        sp_item = sp_list.items.get_by_id(status_item_id)
        sp_item.set_property(status_field_internal_name, status_value)
        sp_item.update()
        ctx.execute_query()
        #print(
        #    f"Updated list '{status_list_title}', item ID {status_item_id}: "
        #    f"{status_field_internal_name} = '{status_value}'"
        #)

    return uploaded_file

def update_toc(
    input_path: str,
    output_path: str | None = None,
    visible: bool = False,
    read_only: bool = False
) -> str:
    """
    Opens a Word document, updates all Tables of Contents and fields, and saves.
    Returns the path of the saved document.
    """
    if not os.path.isfile(input_path):
        raise FileNotFoundError(f"Input file not found: {input_path}")

    if output_path is None:
        # Save back to the same file by default
        output_path = input_path

    word = None
    doc = None

    try:
        word = win32.gencache.EnsureDispatch("Word.Application")
        word.Visible = visible
        word.DisplayAlerts = False  # Suppress prompts

        # Open the document
        doc = word.Documents.Open(input_path, ReadOnly=read_only)

        # --- Update all TOCs (there can be multiple) ---
        # Update page numbers first (faster), then full content if needed.
        for toc in doc.TablesOfContents:
            toc.UpdatePageNumbers()
        for toc in doc.TablesOfContents:
            toc.Update()

        # --- Update all fields throughout the document (headers/footers too) ---
        # This is often necessary if you have cross-references, captions, etc.
        doc.Fields.Update()

        # Update fields in headers/footers across all sections
        for section in doc.Sections:
            for header in section.Headers:
                header.Range.Fields.Update()
            for footer in section.Footers:
                footer.Range.Fields.Update()

        # Save (use SaveAs2 if changing format or path)
        if os.path.abspath(output_path) == os.path.abspath(input_path):
            doc.Save()
        else:
            # 16 = wdFormatXMLDocument (.docx)
            doc.SaveAs2(output_path, FileFormat=16)

        return output_path

    finally:
        if doc is not None:
            doc.Close(SaveChanges=False)
        if word is not None:
            word.Quit()

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
    _temp_img_dir = tempfile.mkdtemp(prefix="report_images_")
    _image_cache = {}

    def make_image(img_src: str, *, force_width_cm: float = None, max_height_cm: float = np.inf):
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
    paragraphs_dachprobenanalyse = data.get("paragraphs_dachprobenanalyse", [])#[0]
    paragraphs_dachsichtbegutachtung = data.get("paragraphs_dachsichtbegutachtung", [])#[0]
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
                p["image"] = make_image(p.pop("src"), force_width_cm=16.0, max_height_cm=8.0) or "Bild nicht verfügbar"
            elif "cols" in p and isinstance(p["cols"], list) and len(p["cols"]) > 2:
                # Check if the last column is an image source
                image_src = p["cols"][-1]
                if isinstance(image_src, str) and image_src.lower().endswith(('.png', '.jpg', '.jpeg')):
                    p["cols"][-1] = make_image(image_src, force_width_cm=6.0, max_height_cm=8.0) or "Bild nicht verfügbar"

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
    #print(f"Report generated: {output_path}")

    # cleanup temp images
    try:
        shutil.rmtree(_temp_img_dir)
    except Exception:
        pass

    # ── 4. Export to PDF elegantly ──────────────────────────────────────
    if not export_pdf:
        #print("PDF export skipped.")
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
    load_dotenv() # Load from .env file

    # === Konfiguration ===
    # App-Registrierung in Azure AD
    tenant_id = os.getenv("TENANT_ID", "your-tenant-id")
    client_id = os.getenv("CLIENT_ID", "your-client-id")

    # Client-Secret (geheim)
    thumbprint = os.getenv("THUMBPRINT", "your-thumbprint")
    passphrase = os.getenv("PASSPHRASE", "your-passphrase")

    # SharePoint-Details
    site_url = os.getenv("SITE_URL", "https://test")
    list_name = os.getenv("LIST_NAME", "Tab_Kundendaten")

    # === Token holen ===
    authority = f"{os.getenv('AUTHORITY')}/{tenant_id}"
    scope = [os.getenv("SCOPE")]
    # app = ConfidentialClientApplication(client_id, authority=authority, client_credential=client_secret)
    # token = app.acquire_token_for_client(scopes=scope)

    # Get Cert File
    cert_path = Path(
        os.getenv("CERT_PATH", "/run/secrets/mycert.pem")
    )

    # if not CERT_PATH.exists():
    #     CERT_PATH = r"C:\Users\Vincent\Desktop\NorproofDocPopCert.pem"
    #     #raise FileNotFoundError(f"Certificate not found: {CERT_PATH}")

    url = os.getenv("URL")
    ctx = ClientContext(base_url=url).with_client_certificate(
        tenant=tenant_id, 
        client_id=client_id,
        thumbprint=thumbprint,
        cert_path=cert_path
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
        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as f:
            output_docx = f.name
        input_docx = os.path.join(os.path.dirname(__file__), output_docx)
        build_report(
        data=dict_result,
        template_path="Dachcheck-Inhaltsverzeichnis.docx",
        output_path=output_docx,
        export_pdf=False,
        ctx=ctx
        )
        #saved = update_toc(input_docx, output_path=None, visible=False, read_only=False)
        uploaded = upload_file_to_sharepoint_folder(
        ctx=ctx,
        local_file_path=output_docx,
        folder_server_relative_url=f"/sites/Norproof/Projektdaten/{project_id}",
        status_list_title="Tab_Projekte",
        status_item_id=ID,
        status_field_internal_name="Status",
        status_value="Gutachten erstellt",
        )
        os.remove(output_docx)  # optional: clean up local file after upload
        # pdf_output = "Output" + ".pdf"
        # try:
        #     result = subprocess.run([
        #         "soffice", "--headless", "--convert-to", "pdf", "--outdir",
        #         os.path.dirname(os.path.abspath("Output.docx")) or ".", "Output.docx"
        #     ], check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        #     if os.path.exists(pdf_output):
        #         print(f"PDF exported: {pdf_output}")
        #     else:
        #         print("PDF export failed: Output file not found.")
        # except Exception as e:
        #     print("PDF export failed. Please ensure LibreOffice is installed and 'soffice' is in your PATH.")
        #     print(f"Error: {e}")
        # print(f"Updated TOC and fields saved to: {saved}")


# def main():
#     # Example usage
#     input_docx = os.path.join(os.path.dirname(__file__), "Output.docx")
#     dict_result = extract_from_sharepoint(gutachten_texte_filtered_items, texts)
#     build_report(
#     data=dict_result,
#     template_path="Dachcheck-Inhaltsverzeichnis.docx",
#     output_path="Output.docx",
#     export_pdf=False,
#     ctx=ctx
#     )
#     #saved = update_toc(input_docx, output_path=None, visible=False, read_only=False)
#     uploaded = upload_file_to_sharepoint_folder(
#     ctx=ctx,
#     local_file_path="Output.docx",
#     folder_server_relative_url=f"/sites/Norproof/Projektdaten/{project_id}",
#     status_list_title="Tab_Projekte",
#     status_item_id=ID,
#     status_field_internal_name="Status",
#     status_value="Gutachten erstellt",
#     )
#     os.remove("Output.docx")  # optional: clean up local file after upload
#     # pdf_output = "Output" + ".pdf"
#     # try:
#     #     result = subprocess.run([
#     #         "soffice", "--headless", "--convert-to", "pdf", "--outdir",
#     #         os.path.dirname(os.path.abspath("Output.docx")) or ".", "Output.docx"
#     #     ], check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
#     #     if os.path.exists(pdf_output):
#     #         print(f"PDF exported: {pdf_output}")
#     #     else:
#     #         print("PDF export failed: Output file not found.")
#     # except Exception as e:
#     #     print("PDF export failed. Please ensure LibreOffice is installed and 'soffice' is in your PATH.")
#     #     print(f"Error: {e}")
#     # print(f"Updated TOC and fields saved to: {saved}")

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"Job failed: {e}", file=sys.stderr)
        raise  # oder sys.exit(1)