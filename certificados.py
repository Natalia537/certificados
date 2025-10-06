# certificados.py
import io
import os
import re
import zipfile
import shutil
import tempfile
import platform
import subprocess
import unicodedata
from zipfile import ZipFile
from pathlib import Path

import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate

# ------------------ Utilidades ------------------

# Detecta {{ PLACEHOLDER }} en document.xml, headers y footers
PLACEHOLDER_RE = re.compile(r"{{\s*([A-Za-z0-9_ÁÉÍÓÚÑáéíóúüÜ\-\./ ]+?)\s*}}")

def strip_accents_upper(s: str) -> str:
    s = unicodedata.normalize("NFKD", s)
    s = "".join(c for c in s if not unicodedata.combining(c))
    return s.upper()

def normalize_key(s: str) -> str:
    if s is None:
        return ""
    s = str(s).strip().replace("\n", " ")
    s = re.sub(r"\s+", " ", s)  # compacta espacios
    return strip_accents_upper(s)

def sanitize_filename(name: str) -> str:
    name = re.sub(r'[<>:"/\\|?*\x00-\x1F]', "_", str(name))
    name = name.strip().strip(".")
    return (name or "documento")[:200]

def extract_placeholders_full(docx_bytes: bytes):
    """
    Lee el .docx y devuelve:
    - placeholders_set: set de placeholders NORMALIZADOS
    - norm_to_original: dict {NORMALIZADO: un_ejemplar_original}
    """
    placeholders_set = set()
    norm_to_original = {}
    with zipfile.ZipFile(io.BytesIO(docx_bytes), "r") as z:
        candidates = ["word/document.xml"]
        for name in z.namelist():
            if name.startswith("word/header") and name.endswith(".xml"):
                candidates.append(name)
            if name.startswith("word/footer") and name.endswith(".xml"):
                candidates.append(name)

        for part in candidates:
            if part in z.namelist():
                xml = z.read(part).decode("utf-8", errors="ignore")
                for m in PLACEHOLDER_RE.findall(xml):
                    original = m.strip()
                    norm = normalize_key(original)
                    placeholders_set.add(norm)
                    # Conserva la primera forma "original" encontrada
                    norm_to_original.setdefault(norm, original)
    return placeholders_set, norm_to_original

def can_convert_pdf() -> bool:
    """¿Existe docx2pdf (Word) o LibreOffice en el entorno?"""
    try:
        from docx2pdf import convert  # noqa: F401
        return True
    except Exception:
        pass
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    return bool(soffice)

def try_docx_to_pdf(input_docx: Path, output_pdf: Path) -> bool:
    """
    Convierte DOCX a PDF.
    - Windows/Mac: intenta docx2pdf (requiere Word).
    - Linux: intenta LibreOffice (soffice).
    Retorna True si logró convertir.
    """
    # 1) docx2pdf
    try:
        from docx2pdf import convert as docx2pdf_convert
        docx2pdf_convert(str(input_docx), str(output_pdf))
        return output_pdf.exists()
    except Exception:
        pass

    # 2) LibreOffice
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if soffice:
        try:
            outdir = output_pdf.parent
            cmd = [
                soffice, "--headless", "--convert-to", "pdf",
                "--outdir", str(outdir), str(input_docx)
            ]
            subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            gen_path = input_docx.with_suffix(".pdf")
            gen_file = outdir / gen_path.name
            if gen_file.exists():
                if gen_file != output_pdf:
                    gen_file.replace(output_pdf)
                return True
        except Exception:
            pass

    return False

def render_docx_from_template(template_bytes: bytes, context: dict) -> bytes:
    tpl = DocxTemplate(io.BytesIO(template_bytes))
    tpl.render(context)
    out = io.BytesIO()
    tpl.save(out)
    return out.getvalue()

def build_context_for_render(row: pd.Series,
                             defaults_norm: dict,
                             placeholders_norm: set,
                             norm_to_original: dict) -> dict:
    """
    Construye el contexto para docxtpl con CLAVES **ORIGINALES** tal como aparecen en el .docx.
    Los valores se buscan en la fila por clave NORMALIZADA.
    """
    context = {}

    # Index rápido: {col_norm: valor}
    row_map = {}
    for col_name, val in row.items():
        col_norm = normalize_key(col_name)
        row_map[col_norm] = "" if pd.isna(val) else val

    # Para cada placeholder detectado en el documento:
    for ph_norm in placeholders_norm:
        ph_original = norm_to_original.get(ph_norm, ph_norm)  # clave que espera docxtpl
        # Valor desde la fila (por nombre normalizado)
        val = row_map.get(ph_norm, None)
        if (val is None or val == "") and ph_norm in defaults_norm:
            val = defaults_norm[ph_norm]
        # Si aún None, deja vacío
        if val is None:
            val = ""
        context[ph_original] = val

    return context

# ------------------ UI ------------------

st.set_page_config(page_title="Generador de Certificados DOCX/PDF", layout="wide")
st.title("🧾 Generador masivo de certificados (Word/PDF)")

with st.sidebar:
    st.markdown("### Instrucciones rápidas")
    st.write("- En tu **machote .docx**, usa variables como `{{NOMBRE}}`, `{{FECHA}}`...")
    st.write("- En tu **Excel**, crea columnas con esos mismos nombres (sin llaves).")
    st.write("- Los nombres se comparan **sin acentos y sin mayúsculas**.")
    st.write("- Puedes definir **valores por defecto** si faltan columnas en el Excel.")

col1, col2 = st.columns([1, 1])

with col1:
    tpl_file = st.file_uploader("Sube el machote (.docx)", type=["docx"])
    xls_file = st.file_uploader("Sube el Excel de datos", type=["xlsx", "xls"])

with col2:
    defaults_box = st.expander("➕ Valores por defecto (si faltan en el Excel)", expanded=False)
    defaults_input = {}
    with defaults_box:
        st.write("Formato: `CLAVE=valor`, una por línea. Ej.:")
        st.code("FECHA=2025-10-15\nTIPO_DE_CHARLA=Magistral", language="text")
        default_kvs = st.text_area("Pega aquí tus valores por defecto", value="")
        for line in default_kvs.splitlines():
            if "=" in line:
                k, v = line.split("=", 1)
                defaults_input[normalize_key(k)] = v.strip()

    sheet_name = None
    if xls_file:
        try:
            x = pd.ExcelFile(xls_file)
            sheet_name = st.selectbox("Hoja del Excel", x.sheet_names, index=0)
        except Exception as e:
            st.error(f"No se pudo leer el Excel: {e}")

st.markdown("---")

if tpl_file and xls_file and sheet_name:
    # Leer bytes del template y placeholders detectados
    tpl_bytes = tpl_file.read()
    placeholders_norm, norm_to_original = extract_placeholders_full(tpl_bytes)

    st.subheader("🔎 Placeholders detectados en el machote")
    if placeholders_norm:
        # Muestra las formas originales únicas en orden alfabético por su versión normalizada
        originals_sorted = [norm_to_original[n] for n in sorted(placeholders_norm)]
        st.write(", ".join(originals_sorted))
    else:
        st.info("No se detectaron placeholders `{{...}}` en el documento. Se generarán copias tal cual.")

    # Leer datos
    try:
        df = pd.read_excel(xls_file, sheet_name=sheet_name, dtype=str)
        df = df.fillna("")
    except Exception as e:
        st.error(f"Error leyendo la hoja '{sheet_name}': {e}")
        st.stop()

    # Normalizar encabezados solo para facilitar matching (pero mantenemos df con encabezados originales para mostrar)
    df_norm = df.copy()
    df_norm.columns = [normalize_key(c) for c in df.columns]

    st.subheader("🧾 Columnas en el Excel")
    st.write(", ".join(map(str, df.columns)))

    # Mostrar coincidencias
    st.subheader("🧭 Coincidencias (placeholder ⇄ columna)")
    hits, missing = [], []
    for ph_norm in placeholders_norm:
        if ph_norm in df_norm.columns:
            hits.append(norm_to_original.get(ph_norm, ph_norm))
        else:
            missing.append(norm_to_original.get(ph_norm, ph_norm))

    if hits:
        st.success("Encontradas columnas para: " + ", ".join(hits))
    if missing:
        st.info("Faltan columnas para: " + ", ".join(missing) + ". Puedes cubrirlas con valores por defecto.")

    # Detectar automáticamente columna de NOMBRE
    candidatos_nombre = [
        "NOMBRE", "NOMBRE COMPLETO", "NOMBRE Y APELLIDO",
        "ALUMNO", "ESTUDIANTE", "PARTICIPANTE",
        "NAME", "FULL NAME"
    ]
    auto_nombre = None
    for c in df_norm.columns:
        if c in candidatos_nombre:
            auto_nombre = c
            break
    if auto_nombre is None and len(df_norm.columns) > 0:
        auto_nombre = df_norm.columns[0]  # fallback

    st.subheader("👤 Columna a usar como NOMBRE para el archivo")
    # Para el selectbox mostramos los encabezados originales, pero seleccionamos por su versión normalizada
    opciones_originales = list(df.columns)
    # Mapa original->normalizado
    original_to_norm = {orig: normalize_key(orig) for orig in df.columns}
    # Encontrar índice sugerido
    suggested_idx = 0
    if auto_nombre:
        for i, orig in enumerate(opciones_originales):
            if original_to_norm[orig] == auto_nombre:
                suggested_idx = i
                break

    nombre_col_original = st.selectbox(
        "Selecciona la columna que contiene el nombre de la persona (para nombrar los archivos)",
        options=opciones_originales,
        index=suggested_idx
    )
    nombre_col_norm = normalize_key(nombre_col_original)
    st.caption(f"Los archivos se guardarán como **{{Nombre}} - Certificado** usando la columna: **{nombre_col_original}**.")

    # ------------- Botones de generación -------------
    c1, c2 = st.columns(2)

    with c1:
        if st.button("⬇️ Generar ZIP de DOCX", type="primary"):
            with st.spinner("Generando documentos DOCX..."):
                memory_zip = io.BytesIO()
                with ZipFile(memory_zip, "w", compression=zipfile.ZIP_DEFLATED) as zf:
                    for i in range(len(df_norm)):
                        row_original = df.iloc[i]
                        row_norm = df_norm.iloc[i]
                        # Construir contexto para RENDER con claves ORIGINALES del docx
                        context = build_context_for_render(
                            row=row_norm,
                            defaults_norm=defaults_input,
                            placeholders_norm=placeholders_norm,
                            norm_to_original=norm_to_original
                        )
                        # Renderizar DOCX
                        out_bytes = render_docx_from_template(tpl_bytes, context)

                        # Nombre de archivo basado en la columna elegida
                        base_name_val = row_norm.get(nombre_col_norm, "")
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

    with c2:
        pdf_ok = can_convert_pdf()
        pdf_btn = st.button("⬇️ Generar ZIP de PDF", disabled=not pdf_ok)

        if not pdf_ok:
            st.info(
                "⚠️ Conversión a PDF no disponible en este entorno.\n\n"
                "- En Windows/Mac, instala **Microsoft Word** para usar `docx2pdf`.\n"
                "- En Linux, instala **LibreOffice** (`soffice`) y asegúrate de que esté en el PATH."
            )

        if pdf_btn:
            with st.spinner("Generando documentos PDF..."):
                tempdir = tempfile.TemporaryDirectory()
                outdir = Path(tempdir.name)
                pdf_zip = io.BytesIO()

                docx_paths = []
                for i in range(len(df_norm)):
                    row_original = df.iloc[i]
                    row_norm = df_norm.iloc[i]

                    context = build_context_for_render(
                        row=row_norm,
                        defaults_norm=defaults_input,
                        placeholders_norm=placeholders_norm,
                        norm_to_original=norm_to_original
                    )
                    # Nombre base
                    base_name_val = row_norm.get(nombre_col_norm, "")
                    base_name_val = sanitize_filename(base_name_val) if base_name_val else f"documento_{i+1}"
                    docx_path = outdir / f"{base_name_val} - Certificado.docx"

                    # Render y guardar a disco
                    doc_bytes = render_docx_from_template(tpl_bytes, context)
                    docx_path.write_bytes(doc_bytes)
                    docx_paths.append((base_name_val, docx_path))

                # Convertir a PDF
                pdf_paths = []
                for base_name_val, docx_path in docx_paths:
                    pdf_path = outdir / f"{base_name_val} - Certificado.pdf"
                    ok = try_docx_to_pdf(docx_path, pdf_path)
                    if ok and pdf_path.exists():
                        pdf_paths.append(pdf_path)

                if not pdf_paths:
                    st.error(
                        "No se pudieron generar PDFs.\n\n"
                        "- En Windows/Mac, instala Microsoft Word para usar `docx2pdf`.\n"
                        "- En Linux, instala **LibreOffice** (soffice) y asegúrate de que esté en el PATH."
                    )
                else:
                    with ZipFile(pdf_zip, "w", compression=zipfile.ZIP_DEFLATED) as zf:
                        for p in pdf_paths:
                            zf.write(p, arcname=p.name)
                    pdf_zip.seek(0)
                    st.download_button(
                        "Descargar PDF.zip",
                        data=pdf_zip,
                        file_name="certificados_pdf.zip",
                        mime="application/zip"
                    )

                tempdir.cleanup()

# ------------- Footer -------------
st.markdown("---")
st.caption(
    "Tip: usa placeholders como `{{NOMBRE}}`, `{{FECHA}}`, `{{CONFERENCIA}}` en el Word. "
    "La app empata nombres **sin acentos y sin mayúsculas**. "
    "Si faltan columnas, puedes cubrirlos con valores por defecto arriba."
)
