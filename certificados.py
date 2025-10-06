import io
import os
import re
import zipfile
import shutil
import tempfile
import platform
import subprocess
from zipfile import ZipFile
from pathlib import Path

import streamlit as st
import pandas as pd

from docxtpl import DocxTemplate

# ---------- Utilidades ----------
PLACEHOLDER_RE = re.compile(r"{{\s*([A-Za-z0-9_ÁÉÍÓÚÑáéíóúüÜ\-\./ ]+?)\s*}}")

def normalize_key(s: str) -> str:
    if s is None:
        return ""
    # Normaliza: mayúsculas, quita dobles espacios, reemplaza espacios por _
    s = s.strip().replace("\n", " ")
    s = re.sub(r"\s+", " ", s)
    return s

def sanitize_filename(name: str) -> str:
    # Quita caracteres problemáticos
    name = re.sub(r'[<>:"/\\|?*\x00-\x1F]', "_", name)
    name = name.strip().strip(".")
    if not name:
        name = "documento"
    return name[:200]

def extract_placeholders_from_docx(docx_bytes: bytes) -> set:
    """Lee el XML de Word (document, headers, footers) y extrae {{PLACEHOLDER}} robustamente."""
    ph = set()
    with zipfile.ZipFile(io.BytesIO(docx_bytes), "r") as z:
        candidates = ["word/document.xml"]
        # headers / footers
        for name in z.namelist():
            if name.startswith("word/header") and name.endswith(".xml"):
                candidates.append(name)
            if name.startswith("word/footer") and name.endswith(".xml"):
                candidates.append(name)

        for part in candidates:
            if part in z.namelist():
                xml = z.read(part).decode("utf-8", errors="ignore")
                for m in PLACEHOLDER_RE.findall(xml):
                    ph.add(normalize_key(m))
    return ph

def try_docx_to_pdf(input_docx: Path, output_pdf: Path) -> bool:
    """
    Convierte DOCX a PDF.
    - Windows/Mac: intenta docx2pdf (requiere Word).
    - Linux: intenta LibreOffice (soffice).
    Retorna True si logró convertir.
    """
    sys = platform.system()

    # Intento con docx2pdf si está instalado
    try:
        from docx2pdf import convert as docx2pdf_convert  # type: ignore
        # docx2pdf requiere path, no BytesIO
        docx2pdf_convert(str(input_docx), str(output_pdf))
        return output_pdf.exists()
    except Exception:
        pass

    # Fallback LibreOffice (Linux o si el usuario lo tiene)
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if soffice:
        try:
            # LibreOffice solo acepta carpeta de salida, no nombre de archivo exacto
            outdir = output_pdf.parent
            cmd = [
                soffice, "--headless", "--convert-to", "pdf",
                "--outdir", str(outdir), str(input_docx)
            ]
            subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            # LibreOffice genera <nombre>.pdf en outdir con el mismo nombre base
            gen_pdf = input_docx.with_suffix(".pdf").name
            gen_path = outdir / gen_pdf
            if gen_path.exists():
                # Renombrar si hace falta
                if gen_path != output_pdf:
                    gen_path.replace(output_pdf)
                return True
        except Exception:
            pass

    return False

def build_context(row: pd.Series, default_values: dict) -> dict:
    ctx = {}
    # normalizamos nombres: Excel y placeholders se comparan en mayúsculas y con espacios limpios
    for k, v in row.items():
        nk = normalize_key(str(k))
        ctx[nk] = "" if pd.isna(v) else v

    # Defaults para placeholders no presentes en el Excel
    for dk, dv in default_values.items():
        if dk not in ctx or (isinstance(ctx[dk], str) and ctx[dk] == ""):
            ctx[dk] = dv

    return ctx

def render_one_docx(template_bytes: bytes, context: dict) -> bytes:
    tpl = DocxTemplate(io.BytesIO(template_bytes))
    tpl.render(context)
    out = io.BytesIO()
    tpl.save(out)
    return out.getvalue()

def resolve_filename(pattern: str, context: dict, idx: int) -> str:
    """Construye nombre de archivo con el patrón usando {{VAR}} del contexto."""
    def repl(m):
        key = normalize_key(m.group(1))
        val = context.get(key, "")
        return str(val) if val is not None else ""

    name = re.sub(PLACEHOLDER_RE, repl, pattern)
    name = name if name.strip() else f"documento_{idx+1}"
    return sanitize_filename(name)

# ---------- UI ----------
st.set_page_config(page_title="Generador de Certificados DOCX/PDF", layout="wide")
st.title("🧾 Generador masivo de certificados (Word/PDF)")

with st.sidebar:
    st.markdown("### Instrucciones rápidas")
    st.write("- En tu **machote .docx**, usa variables como `{{NOMBRE}}`, `{{FECHA}}`...")
    st.write("- En tu **Excel**, crea columnas con esos mismos nombres (sin llaves).")
    st.write("- Opcional: define **valores por defecto** para placeholders que no vengan en el Excel.")
    st.write("- Define el **patrón de nombre** de salida: por ejemplo `{{NOMBRE}} - {{CONFERENCIA}}`.")

col1, col2 = st.columns([1,1])

with col1:
    tpl_file = st.file_uploader("Sube el machote (.docx)", type=["docx"])
    xls_file = st.file_uploader("Sube el Excel de datos", type=["xlsx", "xls"])
    filename_pattern = st.text_input(
        "Patrón de nombre de archivo (sin extensión)",
        value="{{NOMBRE}} - Certificado"
    )
    st.caption("Puedes usar cualquier placeholder del machote, ej.: `{{NOMBRE}} - {{CONFERENCIA}}`")

with col2:
    default_expander = st.expander("➕ Valores por defecto (si faltan en Excel)", expanded=False)
    defaults_input = {}
    with default_expander:
        st.write("Agrega pares *Placeholder → Valor por defecto* (opcional).")
        # Pequeño UI para capturar pares dinámicos
        default_kvs = st.text_area(
            "Formato: una por línea. Ej.:  \nFECHA=2025-10-15\nTIPO_DE_CHARLA=Magistral",
            value=""
        )
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
    placeholders = extract_placeholders_from_docx(tpl_bytes)
    st.subheader("🔎 Placeholders detectados en el machote")
    if placeholders:
        st.write(", ".join(sorted(placeholders)))
    else:
        st.warning("No se detectaron placeholders `{{...}}` en el documento. Revisa tu machote.")

    # Leer datos
    try:
        df = pd.read_excel(xls_file, sheet_name=sheet_name, dtype=str)
        df = df.fillna("")
    except Exception as e:
        st.error(f"Error leyendo la hoja '{sheet_name}': {e}")
        st.stop()

    # Normalizar encabezados del DF para facilitar matching
    norm_columns = [normalize_key(c) for c in df.columns]
    df.columns = norm_columns

    st.subheader("🧾 Columnas en el Excel")
    st.write(", ".join(df.columns))

    # Mostrar mapeo simple: qué placeholders tienen columna matching
    st.subheader("🧭 Coincidencias (placeholder ⇄ columna)")
    hits, missing = [], []
    for ph in placeholders:
        if ph in df.columns:
            hits.append(ph)
        else:
            missing.append(ph)

    if hits:
        st.success("Encontradas columnas para: " + ", ".join(hits))
    if missing:
        st.info("Faltan columnas para: " + ", ".join(missing) + ". Puedes cubrirlos con valores por defecto.")

    # Botones de generación
    c1, c2 = st.columns(2)

    with c1:
        if st.button("⬇️ Generar ZIP de DOCX", type="primary"):
            with st.spinner("Generando documentos DOCX..."):
                memory_zip = io.BytesIO()
                with ZipFile(memory_zip, "w", compression=zipfile.ZIP_DEFLATED) as zf:
                    for i, row in df.iterrows():
                        ctx = build_context(row, defaults_input)
                        out_bytes = render_one_docx(tpl_bytes, ctx)
                        out_name = resolve_filename(filename_pattern, ctx, i) + ".docx"
                        zf.writestr(out_name, out_bytes)
                memory_zip.seek(0)
            st.download_button(
                "Descargar DOCX.zip",
                data=memory_zip,
                file_name="certificados_docx.zip",
                mime="application/zip"
            )

    with c2:
        if st.button("⬇️ Generar ZIP de PDF"):
            with st.spinner("Generando documentos PDF..."):
                tempdir = tempfile.TemporaryDirectory()
                outdir = Path(tempdir.name)
                pdf_zip = io.BytesIO()

                # 1) generar todos los DOCX a disco
                docx_paths = []
                rows = list(df.iterrows())
                for i, row in rows:
                    ctx = build_context(row, defaults_input)
                    out_name = resolve_filename(filename_pattern, ctx, i)
                    docx_path = outdir / f"{out_name}.docx"
                    # render
                    doc_bytes = render_one_docx(tpl_bytes, ctx)
                    docx_path.write_bytes(doc_bytes)
                    docx_paths.append((out_name, docx_path))

                # 2) convertir cada DOCX a PDF
                pdf_paths = []
                for out_name, docx_path in docx_paths:
                    pdf_path = outdir / f"{out_name}.pdf"
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
                    # 3) empaquetar ZIP en memoria
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
