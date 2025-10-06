# certificados.py
import io
import re
import zipfile
import shutil
import tempfile
import subprocess
import unicodedata
from zipfile import ZipFile
from pathlib import Path

import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate

# ============== NUEVO: PDF nativo (sin Word/LibreOffice) ==============
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import inch

# ===================== Utilidades =====================

def strip_accents_upper(s: str) -> str:
    s = unicodedata.normalize("NFKD", s)
    s = "".join(c for c in s if not unicodedata.combining(c))
    return s.upper()

def normalize_key(s: str) -> str:
    if s is None:
        return ""
    s = str(s).strip().replace("\n", " ")
    s = re.sub(r"\s+", " ", s)
    return strip_accents_upper(s)

def sanitize_filename(name: str) -> str:
    name = re.sub(r'[<>:"/\\|?*\x00-\x1F]', "_", str(name))
    name = name.strip().strip(".")
    return (name or "documento")[:200]

def render_docx_from_template(template_bytes: bytes, context: dict) -> bytes:
    tpl = DocxTemplate(io.BytesIO(template_bytes))
    tpl.render(context)
    out = io.BytesIO()
    tpl.save(out)
    return out.getvalue()

# ============== Detección (best effort, opcional) ==============

PLACEHOLDER_RE = re.compile(r"{{\s*([^{}}]+?)\s*}}")

def extract_placeholders_best_effort(docx_bytes: bytes):
    """
    Devuelve una lista *posible* de placeholders leyendo el XML.
    OJO: Word a veces parte las llaves en "runs" y no aparecen completas.
    Igual usamos esto SOLO para sugerir; el usuario puede añadir manualmente.
    """
    placeholders = set()
    try:
        with zipfile.ZipFile(io.BytesIO(docx_bytes), "r") as z:
            for name in z.namelist():
                if name.startswith("word/") and name.endswith(".xml"):
                    xml = z.read(name).decode("utf-8", errors="ignore")
                    for m in PLACEHOLDER_RE.findall(xml):
                        placeholders.add(m.strip())
    except Exception:
        pass
    candidates = [p for p in placeholders if len(p) <= 80]
    return sorted(candidates, key=lambda x: normalize_key(x))

# ===================== PDF nativo simple =====================

def crear_pdf_certificado(nombre_archivo_base: str, datos_dict: dict) -> bytes:
    """
    Genera un PDF simple con los datos del certificado.
    No depende de Word ni de LibreOffice.
    """
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    width, height = A4

    # Título
    c.setFont("Helvetica-Bold", 20)
    c.drawCentredString(width / 2, height - 1.6 * inch, "CERTIFICADO DE PARTICIPACIÓN")

    # Si existe un campo tipo 'Nombre', lo destacamos al centro
    nombre_keys = ["NOMBRE", "NOMBRE COMPLETO", "NOMBRE Y APELLIDO", "ALUMNO", "ESTUDIANTE", "PARTICIPANTE", "NAME", "FULL NAME"]
    nombre_val = ""
    for k, v in datos_dict.items():
        if normalize_key(k) in [normalize_key(x) for x in nombre_keys]:
            nombre_val = str(v)
            break
    if nombre_val:
        c.setFont("Helvetica-Bold", 16)
        c.drawCentredString(width / 2, height - 2.2 * inch, nombre_val)

    # Cuerpo de campos (izquierda)
    c.setFont("Helvetica", 12)
    y = height - 3.0 * inch
    margen_x = 1.25 * inch
    for k, v in datos_dict.items():
        k_clean = str(k).strip().strip("{} ")
        texto = f"{k_clean}: {v}"
        c.drawString(margen_x, y, texto)
        y -= 0.35 * inch
        if y < 1.3 * inch:  # salto si se acaba la página
            c.showPage()
            c.setFont("Helvetica", 12)
            y = height - 1.5 * inch

    # Pie
    c.setFont("Helvetica-Oblique", 10)
    c.drawString(margen_x, 1.0 * inch, "Emitido automáticamente por el generador de certificados.")

    c.showPage()
    c.save()
    pdf_data = buffer.getvalue()
    buffer.close()
    return pdf_data

# ===================== App =====================

st.set_page_config(page_title="Generador de Certificados DOCX/PDF", layout="wide")
st.title("🧾 Generador masivo de certificados (Word/PDF)")

with st.sidebar:
    st.markdown("### Instrucciones")
    st.write("1) Sube tu **machote .docx** con placeholders como `{{Nombre}}`, `{{Cédula}}`, `{{Calificación}}`.")
    st.write("2) Sube tu **Excel** con columnas de datos.")
    st.write("3) **Mapea** cada placeholder → columna del Excel (o un valor fijo).")
    st.write("4) Descarga **ZIP de DOCX** o **ZIP de PDF nativo** (no requiere Word).")
    st.caption("Si un placeholder no aparece en 'detectados', agrégalo manualmente abajo.")

col1, col2 = st.columns([1, 1])

with col1:
    tpl_file = st.file_uploader("Sube el machote (.docx)", type=["docx"])
    xls_file = st.file_uploader("Sube el Excel de datos", type=["xlsx", "xls"])

with col2:
    sheet_name = None
    if xls_file:
        try:
            x = pd.ExcelFile(xls_file)
            sheet_name = st.selectbox("Hoja del Excel", x.sheet_names, index=0)
        except Exception as e:
            st.error(f"No se pudo leer el Excel: {e}")

st.markdown("---")

if "mappings" not in st.session_state:
    st.session_state.mappings = []   # cada item: {"placeholder":"Cédula", "column":"Cédula", "default":""}

if tpl_file and xls_file and sheet_name:
    # --- Leer Excel ---
    try:
        df = pd.read_excel(xls_file, sheet_name=sheet_name, dtype=str).fillna("")
    except Exception as e:
        st.error(f"Error leyendo la hoja '{sheet_name}': {e}")
        st.stop()

    cols_original = list(df.columns)
    cols_norm_map = {c: normalize_key(c) for c in cols_original}

    st.subheader("🧾 Columnas del Excel")
    st.write(", ".join(map(str, cols_original)))

    # --- Leer placeholders sugeridos del Word ---
    tpl_bytes = tpl_file.read()
    suggested_placeholders = extract_placeholders_best_effort(tpl_bytes)
    if suggested_placeholders:
        st.subheader("🔎 Placeholders detectados (sugerencias)")
        st.write(", ".join(suggested_placeholders))
    else:
        st.info("No se detectaron placeholders automáticamente. Puedes agregarlos manualmente abajo.")

    st.subheader("🔗 Mapear placeholders del Word ↔ columnas del Excel")

    def add_mapping_if_missing(ph: str, col_guess: str | None):
        for m in st.session_state.mappings:
            if m["placeholder"] == ph:
                return
        st.session_state.mappings.append({
            "placeholder": ph,          # tal cual en Word (con acentos/may/min)
            "column": col_guess or "",  # nombre de columna ORIGINAL
            "default": ""               # valor fijo si la celda viene vacía
        })

    cta_cols = st.columns([1, 1, 2])
    with cta_cols[0]:
        if st.button("➕ Agregar mapeo vacío"):
            st.session_state.mappings.append({"placeholder": "", "column": "", "default": ""})
    with cta_cols[1]:
        if st.button("✨ Autocompletar desde placeholders"):
            for ph in suggested_placeholders:
                ph_norm = normalize_key(ph)
                best = None
                for c in cols_original:
                    if cols_norm_map[c] == ph_norm:
                        best = c
                        break
                add_mapping_if_missing(ph, best)

    new_mappings = []
    for idx, m in enumerate(st.session_state.mappings):
        st.markdown(f"**Mapeo {idx+1}**")
        col_a, col_b, col_c, col_d = st.columns([2, 2, 2, 1])
        placeholder = col_a.text_input("Placeholder (tal como en el Word, sin llaves)", value=m["placeholder"], key=f"ph_{idx}")
        column = col_b.selectbox("Columna del Excel", options=[""] + cols_original, index=([""] + cols_original).index(m["column"]) if m["column"] in cols_original else 0, key=f"col_{idx}")
        default = col_c.text_input("Valor por defecto (si la celda está vacía)", value=m["default"], key=f"def_{idx}")
        remove = col_d.checkbox("Eliminar", value=False, key=f"rm_{idx}")

        if not remove:
            new_mappings.append({"placeholder": placeholder, "column": column, "default": default})

        st.divider()

    st.session_state.mappings = new_mappings

    # Columna que se usará para el NOMBRE DE ARCHIVO
    st.subheader("👤 Columna para el **nombre del archivo**")
    candidatos_nombre = ["NOMBRE", "NOMBRE COMPLETO", "NOMBRE Y APELLIDO", "ALUMNO", "ESTUDIANTE", "PARTICIPANTE", "NAME", "FULL NAME"]
    auto_idx = 0
    for i, c in enumerate(cols_original):
        if cols_norm_map[c] in [normalize_key(x) for x in candidatos_nombre]:
            auto_idx = i
            break
    nombre_col_original = st.selectbox("Selecciona la columna que contiene el nombre de la persona", options=cols_original, index=auto_idx)
    st.caption(f"Los archivos se guardarán como **{{Nombre}} - Certificado** usando la columna: **{nombre_col_original}**.")

    # ========= Botones de generación =========
    c1, c2 = st.columns(2)

    # Validación básica de mapeos
    valid_mappings = [m for m in st.session_state.mappings if m["placeholder"].strip() and (m["column"] or m["default"])]
    if not valid_mappings:
        st.info("Agrega al menos un mapeo (placeholder → columna o valor por defecto) para generar certificados.")

    # ---------------- DOCX ----------------
    with c1:
        if st.button("⬇️ Generar ZIP de DOCX", type="primary", disabled=(len(valid_mappings) == 0)):
            with st.spinner("Generando documentos DOCX..."):
                memory_zip = io.BytesIO()
                with ZipFile(memory_zip, "w", compression=zipfile.ZIP_DEFLATED) as zf:
                    for i in range(len(df)):
                        # Construir contexto EXACTO como el Word espera
                        ctx = {}
                        for m in valid_mappings:
                            key = m["placeholder"]  # EXACTO (con acentos/may/min)
                            if m["column"]:
                                val = df.iloc[i][m["column"]]
                                if pd.isna(val) or val == "":
                                    val = m["default"]
                            else:
                                val = m["default"]
                            ctx[key] = "" if val is None else val

                        # Render y escribir
                        out_bytes = render_docx_from_template(tpl_bytes, ctx)

                        # Nombre de archivo
                        base_name_val = df.iloc[i][nombre_col_original]
                        base_name_val = sanitize_filename(base_name_val) if base_name_val else f"documento_{i+1}"
                        out_name = f"{base_name_val} - Certificado.docx"
                        zf.writestr(out_name, out_bytes)

                memory_zip.seek(0)
            st.download_button(
                "Descargar DOCX.zip",
                data=memory_zip,
                file_name="certificados_docx.zip",
                mime="application/zip"
            )

    # ---------------- PDF nativo (ReportLab) ----------------
    with c2:
        if st.button("⬇️ Generar ZIP de PDF (nativo, sin Word) ", type="secondary", disabled=(len(valid_mappings) == 0)):
            with st.spinner("Generando PDFs..."):
                memory_zip = io.BytesIO()
                with ZipFile(memory_zip, "w", compression=zipfile.ZIP_DEFLATED) as zf:
                    for i in range(len(df)):
                        # Contexto de datos para imprimir en el PDF
                        ctx = {}
                        for m in valid_mappings:
                            key = m["placeholder"]
                            if m["column"]:
                                val = df.iloc[i][m["column"]]
                                if pd.isna(val) or val == "":
                                    val = m["default"]
                            else:
                                val = m["default"]
                            ctx[key] = "" if val is None else val

                        base_name_val = df.iloc[i][nombre_col_original]
                        base_name_val = sanitize_filename(base_name_val) if base_name_val else f"documento_{i+1}"

                        pdf_bytes = crear_pdf_certificado(base_name_val, ctx)
                        zf.writestr(f"{base_name_val} - Certificado.pdf", pdf_bytes)

                memory_zip.seek(0)
            st.download_button(
                "Descargar PDF.zip",
                data=memory_zip,
                file_name="certificados_pdf.zip",
                mime="application/zip"
            )

st.markdown("---")
st.caption("Si algún placeholder no aparece en 'detectados', agrégalo manualmente en los mapeos. "
           "Esta versión genera PDF nativo (ReportLab), por lo que funciona también en la nube.")
