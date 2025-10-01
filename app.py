import io
import re
from datetime import datetime
from typing import Dict, List, Any, Iterable

import streamlit as st
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn

# Regex para {{CHAVE}}
PH_RE = re.compile(r"\{\{([^}]+)\}\}")

# --------------------------
# Utilitários
# --------------------------
def normalize_key(s: str) -> str:
    """Remove acentos e normaliza para UPPER com _."""
    import unicodedata
    s = "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")
    s = re.sub(r"[^A-Za-z0-9]+", "_", s)
    return s.strip("_").upper()

def iter_all_paragraphs(doc: Document) -> Iterable:
    # Corpo
    for p in doc.paragraphs:
        yield p
    # Tabelas
    for t in doc.tables:
        for r in t.rows:
            for c in r.cells:
                for p in c.paragraphs:
                    yield p

def collect_placeholders(doc: Document) -> List[str]:
    found = set()
    # Corpo + tabelas
    for p in iter_all_paragraphs(doc):
        for ph in PH_RE.findall(p.text or ""):
            found.add(ph.strip())
    # Headers/footers
    for sec in doc.sections:
        for hf in [sec.header, sec.first_page_header, sec.even_page_header,
                   sec.footer, sec.first_page_footer, sec.even_page_footer]:
            if hf:
                for p in hf.paragraphs:
                    for ph in PH_RE.findall(p.text or ""):
                        found.add(ph.strip())
    return sorted(found, key=str.lower)

def apply_font_family_and_size(paragraph, font_name: str, size_pt: int):
    """Uniformiza fonte/tamanho (NÃO mexe em negrito/itálico)."""
    for run in paragraph.runs:
        run.font.name = font_name
        rPr = run._element.rPr
        if rPr is not None and rPr.rFonts is not None:
            rPr.rFonts.set(qn('w:ascii'), font_name)
            rPr.rFonts.set(qn('w:hAnsi'), font_name)
            rPr.rFonts.set(qn('w:eastAsia'), font_name)
        run.font.size = Pt(size_pt)

def replace_within_run_text(run, mapping: Dict[str, Any]) -> bool:
    """Substitui {{CHAVE}} dentro de UM run mantendo o estilo do run."""
    txt = run.text or ""
    if "{{" not in txt:
        return False
    placeholders = PH_RE.findall(txt)
    if not placeholders:
        return False
    new_txt = txt
    for raw in placeholders:
        k = raw.strip()
        val = mapping.get(k)
        if val is None:
            val = mapping.get(normalize_key(k))
        if val is None:
            val = ""
        new_txt = new_txt.replace("{{" + k + "}}", str(val))
    if new_txt != txt:
        run.text = new_txt
        return True
    return False

def replace_across_runs_preserving_first_style(paragraph, mapping: Dict[str, Any]) -> bool:
    """
    Se {{CHAVE}} estiver quebrada em vários runs, junta o bloco,
    substitui e grava tudo no PRIMEIRO run (preservando o estilo do primeiro).
    """
    runs = list(paragraph.runs)
    if not runs:
        return False

    changed_any = False
    i = 0
    while i < len(runs):
        txt = runs[i].text or ""
        if "{{" in txt and "}}" in txt:
            i += 1
            continue
        if "{{" in txt and "}}" not in txt:
            j = i
            block_text = txt
            while j + 1 < len(runs) and "}}" not in block_text:
                j += 1
                block_text += runs[j].text or ""
            if "}}" in block_text:
                new_block = block_text
                placeholders = PH_RE.findall(block_text)
                for raw in placeholders:
                    k = raw.strip()
                    val = mapping.get(k)
                    if val is None:
                        val = mapping.get(normalize_key(k))
                    if val is None:
                        val = ""
                    new_block = new_block.replace("{{" + k + "}}", str(val))
                if new_block != block_text:
                    runs[i].text = new_block
                    for x in range(i + 1, j + 1):
                        runs[x].text = ""
                    changed_any = True
                i = j + 1
                continue
        i += 1

    return changed_any

def replace_placeholders_preserving_bold(paragraph, mapping: Dict[str, Any]) -> bool:
    """
    1) Tenta substituir dentro de cada run (preserva bold/itálico daquele run).
    2) Trata casos de placeholder quebrado em múltiplos runs (preserva estilo do 1º run).
    """
    changed = False
    for run in paragraph.runs:
        if replace_within_run_text(run, mapping):
            changed = True
    if replace_across_runs_preserving_first_style(paragraph, mapping):
        changed = True
    return changed

def process_document(template_bytes: bytes,
                     mapping: Dict[str, Any],
                     font_name: str,
                     font_size: int) -> bytes:
    doc = Document(io.BytesIO(template_bytes))

    # Corpo + tabelas
    for p in iter_all_paragraphs(doc):
        replace_placeholders_preserving_bold(p, mapping)
        apply_font_family_and_size(p, font_name, font_size)

    # Headers/Footers
    for sec in doc.sections:
        for hf in [sec.header, sec.first_page_header, sec.even_page_header,
                   sec.footer, sec.first_page_footer, sec.even_page_footer]:
            if hf:
                for p in hf.paragraphs:
                    replace_placeholders_preserving_bold(p, mapping)
                    apply_font_family_and_size(p, font_name, font_size)

    out = io.BytesIO()
    doc.save(out)
    out.seek(0)
    return out.getvalue()

def safe_filename(name: str) -> str:
    return re.sub(r'[\\/:*?"<>|]+', "-", str(name)).strip()

# --------------------------
# UI
# --------------------------
st.set_page_config(page_title="Carta de Homologação — Preenchimento", page_icon="🧰", layout="wide")
st.title("Carta de Homologação de Fornecedores — Preenchimento por {{Chaves}}")
st.caption("Suba o modelo .docx → detectamos as {{CHAVES}} → você preenche → baixamos o .docx final. "
           "Negrito é preservado exatamente como no modelo; apenas fonte/tamanho são uniformizados para Calibri 11.")

st.divider()

template_file = st.file_uploader("Modelo .docx (com placeholders {{CHAVE}})", type=["docx"], accept_multiple_files=False)

with st.sidebar:
    st.header("Configurações")
    font_name = st.text_input("Fonte", value="Calibri")
    font_size = st.number_input("Tamanho (pt)", min_value=8, max_value=20, value=11, step=1)
    st.caption("Obs.: não alteramos negrito; apenas fonte/tamanho para uniformizar o documento.")

# Sugerir data de hoje no formato DD/MM/AAAA
def default_value_for_key(k: str) -> str:
    nk = normalize_key(k)
    if "DATA" in nk:
        return datetime.now().strftime("%d/%m/%Y")
    return ""

if template_file:
    try:
        tmp_doc = Document(template_file)
        keys = collect_placeholders(tmp_doc)
        st.success(f"Placeholders encontrados ({len(keys)}): {keys if keys else '—'}")

        with st.form("form"):
            st.subheader("Preencha os valores para as {{Chaves}} detectadas")
            values: Dict[str, Any] = {}

            col1, col2 = st.columns(2)
            half = (len(keys) + 1) // 2
            left, right = keys[:half], keys[half:]

            with col1:
                for k in left:
                    values[k] = st.text_input(k, value=default_value_for_key(k))
            with col2:
                for k in right:
                    values[k] = st.text_input(k, value=default_value_for_key(k))

            out_name = st.text_input(
                "Nome do arquivo (.docx)",
                value="Homologacao - {RAZAO_SOCIAL_DO_FORNECEDOR}.docx"
            )
            submitted = st.form_submit_button("Gerar documento")

        if submitted:
            # mapping com chaves originais e normalizadas
            mapping: Dict[str, Any] = {}
            for k, v in values.items():
                mapping[k] = v
                mapping[normalize_key(k)] = v

            # nome final com tokens
            final_name = out_name
            for k, v in values.items():
                final_name = final_name.replace("{"+k+"}", safe_filename(v))
                final_name = final_name.replace("{"+normalize_key(k)+"}", safe_filename(v))
            if not final_name.lower().endswith(".docx"):
                final_name += ".docx"
            final_name = safe_filename(final_name) or "Homologacao.docx"

            # reabrir bytes do modelo (reposiciona ponteiro)
            template_file.seek(0)
            template_bytes = template_file.read()

            # processar
            docx_bytes = process_document(
                template_bytes=template_bytes,
                mapping=mapping,
                font_name=font_name,
                font_size=int(font_size),
            )

            st.download_button("⬇️ Baixar DOCX", data=docx_bytes,
                               file_name=final_name,
                               mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

    except Exception as e:
        st.error(f"Erro ao ler o modelo: {e}")
else:
    st.info("Envie o modelo .docx para começar.")
