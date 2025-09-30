import io
import os
import re
import json
import zipfile
import unicodedata
from typing import Dict, List, Any, Iterable

import pandas as pd
import streamlit as st
from docx import Document
from docx.oxml.ns import qn
from docx.shared import Pt

# ==========================
# Utilidades de texto/chaves
# ==========================

PLACEHOLDER_REGEX = re.compile(r"\{\{([^}]+)\}\}")

def strip_accents(s: str) -> str:
    return "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")

def normalize_key(s: str) -> str:
    s = strip_accents(s)
    s = re.sub(r"[^A-Za-z0-9]+", "_", s)
    return s.strip("_").upper()

def collect_placeholders_from_text(text: str) -> List[str]:
    return PLACEHOLDER_REGEX.findall(text or "")

def iter_all_paragraphs(doc: Document) -> Iterable:
    for p in doc.paragraphs:
        yield p
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    yield p

def paragraph_contains_any(p, needles: List[str]) -> bool:
    t = p.text or ""
    return any((n.strip().lower() in t.lower()) for n in needles if n.strip())

def apply_font_to_paragraph(p, font_name: str, size_pt: int, bold: bool):
    for run in p.runs:
        run.font.name = font_name
        rPr = run._element.rPr
        if rPr is not None and rPr.rFonts is not None:
            rPr.rFonts.set(qn('w:ascii'), font_name)
            rPr.rFonts.set(qn('w:hAnsi'), font_name)
            rPr.rFonts.set(qn('w:eastAsia'), font_name)
        run.font.size = Pt(size_pt)
        run.font.bold = bold

def replace_placeholders_in_paragraph(p, mapping: Dict[str, Any]) -> bool:
    full = "".join(run.text for run in p.runs)
    phs = collect_placeholders_from_text(full)
    if not phs:
        return False
    new_full = full
    for raw in phs:
        raw_key = raw.strip()
        val = mapping.get(raw_key)
        if val is None:
            val = mapping.get(normalize_key(raw_key))
        if val is None:
            val = ""
        new_full = new_full.replace("{{" + raw_key + "}}", str(val))
    if new_full != full:
        if not p.runs:
            p.add_run(new_full)
        else:
            p.runs[0].text = new_full
            for r in p.runs[1:]:
                r.text = ""
        return True
    return False

def build_mapping_for_record(rec: Dict[str, Any]) -> Dict[str, Any]:
    mapping: Dict[str, Any] = {}
    for k, v in rec.items():
        mapping[k] = v
        mapping[normalize_key(k)] = v

    # Aliases úteis (p/ seu modelo de fornecedores)
    aliases = {
        "RAZAO_SOCIAL_DO_FORNECEDOR": ["RAZAO_SOCIAL_DO_FORNECEDOR","RAZAO_SOCIAL","RAZÃO_SOCIAL","NOME_FORNECEDOR"],
        "RAZÃO_SOCIAL_DO_FORNECEDOR": ["RAZAO_SOCIAL_DO_FORNECEDOR","RAZAO_SOCIAL","RAZÃO_SOCIAL","NOME_FORNECEDOR"],
        "Razão_Social_do_Fornecedor": ["RAZAO_SOCIAL_DO_FORNECEDOR","RAZAO_SOCIAL","RAZÃO_SOCIAL","NOME_FORNECEDOR"],
        "CNPJ_DO_FORNECEDOR": ["CNPJ","CNPJ_DO_FORNECEDOR"],
        "CNPJ_do_Fornecedor": ["CNPJ","CNPJ_DO_FORNECEDOR"],
        "ENDERECO_DO_FORNECEDOR": ["ENDERECO","ENDEREÇO","ENDERECO_DO_FORNECEDOR"],
        "Endereço_do_Fornecedor": ["ENDERECO","ENDEREÇO","ENDERECO_DO_FORNECEDOR"],
        "DATA": ["DATA","DATA_EMISSAO","DATA_ATUAL"]
    }
    for target, candidates in aliases.items():
        for c in candidates:
            if c in rec:
                mapping[target] = rec[c]
                mapping[normalize_key(target)] = rec[c]
                break
            nc = normalize_key(c)
            if nc in mapping:
                mapping[target] = mapping[nc]
                mapping[normalize_key(target)] = mapping[nc]
                break

    return mapping

def safe_filename(name: str) -> str:
    return re.sub(r'[\\/:*?"<>|]+', "-", str(name)).strip()

# ==========================
# Processamento DOCX/PDF
# ==========================

def process_one(template_bytes: bytes,
                rec: Dict[str, Any],
                font_name: str = "Calibri",
                font_size: int = 11,
                bold_all_by_default: bool = True,
                exceptions_contains: List[str] = None,
                exceptions_placeholders: List[str] = None,
                filename_template: str = "Homologacao - {RAZAO_SOCIAL_DO_FORNECEDOR}.docx",
                try_pdf: bool = False) -> Dict[str, bytes]:

    exceptions_contains = exceptions_contains or []
    exceptions_placeholders_norm = {normalize_key(k) for k in (exceptions_placeholders or [])}

    mapping = build_mapping_for_record(rec)

    # Abrir modelo
    bio = io.BytesIO(template_bytes)
    doc = Document(bio)

    # Replace em corpo e tabelas
    for p in iter_all_paragraphs(doc):
        replace_placeholders_in_paragraph(p, mapping)

    # Replace em headers/footers
    for sec in doc.sections:
        for hf in [sec.header, sec.first_page_header, sec.even_page_header,
                   sec.footer, sec.first_page_footer, sec.even_page_footer]:
            if hf:
                for p in hf.paragraphs:
                    replace_placeholders_in_paragraph(p, mapping)

    # Tipografia e negrito conforme regra
    for p in iter_all_paragraphs(doc):
        is_exc_text = paragraph_contains_any(p, exceptions_contains)
        tnorm = normalize_key(p.text or "")
        is_exc_ph = any(k in tnorm for k in exceptions_placeholders_norm)
        bold = (not is_exc_text) and (not is_exc_ph) and bold_all_by_default
        apply_font_to_paragraph(p, font_name, font_size, bold)

    for sec in doc.sections:
        for hf in [sec.header, sec.first_page_header, sec.even_page_header,
                   sec.footer, sec.first_page_footer, sec.even_page_footer]:
            if hf:
                for p in hf.paragraphs:
                    is_exc_text = paragraph_contains_any(p, exceptions_contains)
                    tnorm = normalize_key(p.text or "")
                    is_exc_ph = any(k in tnorm for k in exceptions_placeholders_norm)
                    bold = (not is_exc_text) and (not is_exc_ph) and bold_all_by_default
                    apply_font_to_paragraph(p, font_name, font_size, bold)

    # Salvar DOCX em memória
    out_docx = io.BytesIO()
    doc.save(out_docx)
    out_docx.seek(0)

    # Nome do arquivo
    out_name = filename_template
    for k, v in mapping.items():
        out_name = out_name.replace("{" + k + "}", safe_filename(v))
        out_name = out_name.replace("{" + normalize_key(k) + "}", safe_filename(v))
    out_name = " ".join(out_name.split())
    if not out_name.lower().endswith(".docx"):
        out_name += ".docx"

    result = {out_name: out_docx.getvalue()}

    # PDF (opcional)
    if try_pdf:
        # Em ambientes sem Word/LibreOffice, PDF pode não funcionar.
        # Vamos tentar python-docx->docx2pdf localmente; se não der, ignoramos silenciosamente.
        try:
            from docx2pdf import convert  # noqa
            # docx2pdf não aceita bytes diretamente; salvar temp local
            import tempfile, os, shutil
            with tempfile.TemporaryDirectory() as td:
                temp_docx = os.path.join(td, "tmp.docx")
                with open(temp_docx, "wb") as f:
                    f.write(out_docx.getvalue())
                pdf_path = os.path.splitext(temp_docx)[0] + ".pdf"
                convert(temp_docx, pdf_path)
                with open(pdf_path, "rb") as f:
                    pdf_bytes = f.read()
                pdf_name = os.path.splitext(out_name)[0] + ".pdf"
                result[pdf_name] = pdf_bytes
        except Exception:
            pass

    return result

# ==========================
# UI (Streamlit)
# ==========================

st.set_page_config(page_title="Homologação de Fornecedores - Automação", page_icon="🧩", layout="wide")
st.title("Homologação de Fornecedores — Automação DOCX → (DOCX/PDF)")

with st.sidebar:
    st.markdown("### Configurações padrão")
    font_name = st.text_input("Fonte", value="Calibri")
    font_size = st.number_input("Tamanho da fonte (pt)", min_value=8, max_value=20, value=11, step=1)
    bold_all = st.checkbox("Negrito por padrão em todo o documento", value=True)
    try_pdf = st.checkbox("Tentar exportar PDF (pode não funcionar no Cloud)", value=False)
    filename_tpl = st.text_input("Nome do arquivo de saída",
                                 value="Homologacao - {RAZAO_SOCIAL_DO_FORNECEDOR}.docx")
    st.markdown("---")
    st.caption("Dica: use chaves do seu dataset no nome (ex.: {RAZAO_SOCIAL} ou {DATA}).")

st.markdown("#### 1) Envie o **modelo DOCX** (com placeholders `{{CHAVE}}`)")
template_file = st.file_uploader("Modelo .docx", type=["docx"], accept_multiple_files=False)

st.markdown("#### 2) Envie os **dados** (CSV ou JSON)")
data_file = st.file_uploader("Arquivo .csv ou .json", type=["csv", "json"], accept_multiple_files=False)

st.markdown("#### 3) Configure **exceções** (trechos que NÃO devem ficar em negrito)")
exceptions_contains = st.text_area(
    "Exceções por texto que o parágrafo contém (1 por linha).",
    value="Descrever aqui os produtos ou serviços aprovados."
).splitlines()

exceptions_placeholders = st.text_area(
    "Exceções por **placeholder** (a chave cujos valores não ficam em negrito, 1 por linha).",
    value="DESCRICAO_PRODUTOS_SERVICOS"
).splitlines()

# Preview de placeholders do modelo
if template_file:
    try:
        # leitura rápida
        tmp_doc = Document(template_file)
        found = set()
        for p in iter_all_paragraphs(tmp_doc):
            for ph in collect_placeholders_from_text(p.text):
                found.add(ph.strip())
        # headers/footers
        for sec in tmp_doc.sections:
            for hf in [sec.header, sec.first_page_header, sec.even_page_header,
                       sec.footer, sec.first_page_footer, sec.even_page_footer]:
                if hf:
                    for p in hf.paragraphs:
                        for ph in collect_placeholders_from_text(p.text):
                            found.add(ph.strip())
        st.success(f"Placeholders encontrados no modelo: {sorted(found) if found else '—'}")
    except Exception as e:
        st.warning(f"Não consegui inspecionar placeholders: {e}")

st.markdown("#### 4) Processar")
go = st.button("Gerar documentos")

def read_datafile_to_records(uploaded) -> List[Dict[str, Any]]:
    if uploaded is None:
        return []
    suffix = uploaded.name.lower().split(".")[-1]
    if suffix == "json":
        data = json.loads(uploaded.read().decode("utf-8"))
        if isinstance(data, dict):
            data = data.get("registros", [])
        if not isinstance(data, list):
            raise ValueError("JSON deve ser uma lista de objetos ou {'registros': [...]} .")
        return data
    elif suffix == "csv":
        df = pd.read_csv(uploaded)
        return df.fillna("").to_dict(orient="records")
    else:
        raise ValueError("Arquivo de dados deve ser .csv ou .json")

if go:
    if not template_file or not data_file:
        st.error("Envie o **modelo DOCX** e o **arquivo de dados**.")
    else:
        try:
            template_bytes = template_file.read()
            records = read_datafile_to_records(data_file)
            if not records:
                st.warning("Nenhum registro encontrado no arquivo de dados.")
            else:
                st.info(f"Processando {len(records)} registro(s)...")
                all_outputs: Dict[str, bytes] = {}
                for rec in records:
                    out_map = process_one(
                        template_bytes=template_bytes,
                        rec=rec,
                        font_name=font_name,
                        font_size=int(font_size),
                        bold_all_by_default=bold_all,
                        exceptions_contains=[s for s in exceptions_contains if s.strip()],
                        exceptions_placeholders=[s for s in exceptions_placeholders if s.strip()],
                        filename_template=filename_tpl,
                        try_pdf=try_pdf
                    )
                    all_outputs.update(out_map)

                # Mostrar lista e botões individuais
                st.success(f"Gerado(s) {len(all_outputs)} arquivo(s).")
                for name, bts in all_outputs.items():
                    st.download_button(f"⬇️ Baixar {name}", data=bts, file_name=name, mime="application/octet-stream")

                # Zip de tudo
                zip_buf = io.BytesIO()
                with zipfile.ZipFile(zip_buf, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
                    for name, bts in all_outputs.items():
                        zf.writestr(name, bts)
                zip_buf.seek(0)
                st.download_button("📦 Baixar todos (ZIP)", data=zip_buf, file_name="saidas_homologacao.zip", mime="application/zip")

        except Exception as e:
            st.exception(e)

st.markdown("---")
st.caption("Observação: exportar PDF pode exigir Microsoft Word (docx2pdf) ou LibreOffice no host. Em alguns ambientes (ex.: Streamlit Cloud), essa etapa pode não estar disponível.")
