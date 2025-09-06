# app.py
import streamlit as st
import pandas as pd
import io
import re
from pathlib import Path
from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from lxml import etree
from datetime import datetime

st.set_page_config(page_title="Estrattore evidenziati e commenti .docx", layout="wide")

NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

def iter_paragraphs(doc):
    for p in doc.paragraphs:
        yield p
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                seen = set()
                for p in cell.paragraphs:
                    pid = id(p._element)
                    if pid in seen:
                        continue
                    seen.add(pid)
                    yield p

def sentence_window(text, start_index):
    parts = re.split(r'(?<=[\.\?\!])\s+', text)
    cum = 0
    for i, s in enumerate(parts):
        if cum <= start_index < cum + len(s) + 1:
            return s.strip()
        cum += len(s) + 1
    return text.strip()

def extract_highlights(doc, filename):
    rows = []
    for p in iter_paragraphs(doc):
        para_text = p.text or ""
        if not para_text.strip():
            continue
        idx = 0
        for r in p.runs:
            text = r.text or ""
            if text:
                if r.font.highlight_color is not None:
                    start_idx = idx
                    context = sentence_window(para_text, start_idx)
                    rows.append({
                        "filename": filename,
                        "type": "highlight",
                        "highlight_color": str(r.font.highlight_color),
                        "extracted_text": text.strip(),
                        "context": context,
                        "paragraph": para_text.strip()
                    })
                idx += len(text)
    return rows

def get_comments_map(doc):
    comments = {}
    part = None
    for rel in doc.part.rels.values():
        if rel.reltype == RT.COMMENTS:
            part = rel.target_part
            break
    if part is None:
        return comments
    root = etree.fromstring(part.blob)
    for c in root.xpath(".//w:comment", namespaces=NS):
        cid = int(c.get("{%s}id" % NS["w"]))
        author = c.get("{%s}author" % NS["w"]) or ""
        date = c.get("{%s}date" % NS["w"]) or ""
        text_runs = c.xpath(".//w:t", namespaces=NS)
        ctext = "".join([t.text or "" for t in text_runs]).strip()
        comments[cid] = {"author": author, "date": date, "comment_text": ctext}
    return comments

def get_commented_spans(doc):
    body = doc.element.body
    open_ids = []
    spans = {}
    def walk(el):
        nonlocal open_ids
        if el.tag == "{%s}commentRangeStart" % NS["w"]:
            cid = int(el.get("{%s}id" % NS["w"]))
            open_ids.append(cid)
            spans.setdefault(cid, [])
        elif el.tag == "{%s}commentRangeEnd" % NS["w"]:
            cid = int(el.get("{%s}id" % NS["w"]))
            if cid in open_ids:
                open_ids.remove(cid)
        if el.tag == "{%s}t" % NS["w"] and el.text:
            for cid in open_ids:
                spans.setdefault(cid, []).append(el.text)
        for child in el:
            walk(child)
    walk(body)
    joined = {cid: "".join(parts).strip() for cid, parts in spans.items() if "".join(parts).strip()}
    return joined

# ---- MULTI CODE/LABEL PARSING ----
def parse_code_labels(text):
    """
    Returns a list of (code, label) pairs extracted from the comment text.
    Supports multiple formats and multiple pairs per comment.

    Supported patterns:
    - Repeated XML-like pairs:
        <codice>CE_P</codice><commento>coinvolgimento progettuale</commento>
        (can appear multiple times)
    - Line-based pairs (one per line):
        CE_P: coinvolgimento progettuale
        CE_S - coinvolgimento scolastico
        [CE_A] altro label
    - Semicolon or pipe separated:
        CE_P: coinvolgimento progettuale; CE_S: coinvolgimento scolastico
    - key=value style:
        codice=CE_P; commento=coinvolgimento progettuale
    """
    results = []

    if not text:
        return results

    t = text.strip()

    # 1) XML-like tags, possibly repeated
    xml_pairs = re.findall(
        r"<\s*codice\s*>(.*?)<\s*/?\\?\s*codice\s*>\s*<\s*commento\s*>(.*?)<\s*/?\\?\s*commento\s*>",
        t, flags=re.IGNORECASE | re.DOTALL
    )
    for code, label in xml_pairs:
        code = code.strip()
        label = label.strip()
        if code or label:
            results.append((code or None, label or None))

    # 2) key=value within the text (can repeat)
    kv_pairs = re.findall(
        r"(?:codice|code)\s*=\s*([A-Za-z0-9_]+)\s*(?:;|,|\||\n|\r|\s)+(?:commento|label)\s*=\s*([^;,\|\n\r]+)",
        t, flags=re.IGNORECASE
    )
    for code, label in kv_pairs:
        code = code.strip()
        label = label.strip()
        if code or label:
            results.append((code or None, label or None))

    # Normalize separators into lines for further parsing
    normalized = re.sub(r"[;|]+", "\n", t)

    # 3) Line-based / bracketed code + label
    for line in normalized.splitlines():
        s = line.strip()
        if not s:
            continue
        # [CODE] label
        m_br = re.match(r"\[\s*([A-Za-z0-9_]+)\s*\]\s+(.+)$", s)
        if m_br:
            results.append((m_br.group(1).strip(), m_br.group(2).strip()))
            continue
        # CODE : label OR CODE - label
        m_col = re.match(r"([A-Za-z0-9_]+)\s*[:\-–]\s*(.+)$", s)
        if m_col:
            results.append((m_col.group(1).strip(), m_col.group(2).strip()))
            continue

    # Deduplicate while preserving order
    seen = set()
    deduped = []
    for code, label in results:
        key = (code or "", label or "")
        if key not in seen and (code or label):
            seen.add(key)
            deduped.append((code, label))

    return deduped

def extract_comments(doc, filename):
    comments_meta = get_comments_map(doc)
    quoted_map = get_commented_spans(doc)
    rows = []
    for cid, meta in comments_meta.items():
        raw = meta.get("comment_text", "")
        pairs = parse_code_labels(raw)
        if not pairs:
            # still record the raw comment without parsed pairs
            rows.append({
                "filename": filename,
                "type": "comment",
                "comment_id": cid,
                "author": meta.get("author", ""),
                "date": meta.get("date", ""),
                "quoted_text": quoted_map.get(cid, ""),
                "comment_text": raw,
                "code": None,
                "label": None,
            })
        else:
            for code, label in pairs:
                rows.append({
                    "filename": filename,
                    "type": "comment",
                    "comment_id": cid,
                    "author": meta.get("author", ""),
                    "date": meta.get("date", ""),
                    "quoted_text": quoted_map.get(cid, ""),
                    "comment_text": raw,
                    "code": code,
                    "label": label,
                })
    return rows

def process_docx_bytes(file_bytes, filename):
    doc = Document(io.BytesIO(file_bytes))
    rows = []
    rows += extract_highlights(doc, filename)
    rows += extract_comments(doc, filename)
    return rows

def export_excel(high_df, com_df, link_df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        if high_df is not None and not high_df.empty:
            high_order = ["filename","extracted_text","highlight_color","context","paragraph"]
            for col in high_order:
                if col not in high_df.columns:
                    high_df[col] = ""
            high_df[high_order].to_excel(writer, index=False, sheet_name="Highlights")
        if com_df is not None and not com_df.empty:
            com_order = ["filename","comment_id","author","date","quoted_text","comment_text","code","label"]
            for col in com_order:
                if col not in com_df.columns:
                    com_df[col] = ""
            com_df[com_order].to_excel(writer, index=False, sheet_name="Commenti (esplosi)")
        if link_df is not None and not link_df.empty:
            link_order = ["filename","comment_id","author","date","code","label","quoted_text","highlight_matches","highlights_concat"]
            for col in link_order:
                if col not in link_df.columns:
                    link_df[col] = ""
            link_df[link_order].to_excel(writer, index=False, sheet_name="Annotazioni collegate")
        summary = pd.DataFrame({
            "sheet": ["Highlights","Commenti (esplosi)","Annotazioni collegate"],
            "rows": [
                0 if high_df is None else len(high_df),
                0 if com_df is None else len(com_df),
                0 if link_df is None else len(link_df)
            ]
        })
        summary.to_excel(writer, index=False, sheet_name="Riepilogo")
    output.seek(0)
    return output

st.title("Estrattore evidenziati e commenti da Word (.docx)")
st.caption("Supporto a più coppie codice/etichetta per singolo commento, filtri avanzati e collegamento con evidenziati.")

uploaded_files = st.file_uploader(
    "Carica file .docx",
    accept_multiple_files=True,
    type=["docx"]
)

if uploaded_files:
    rows = []
    for up in uploaded_files:
        try:
            file_bytes = up.getbuffer()
            rows.extend(process_docx_bytes(file_bytes, up.name))
        except Exception as e:
            st.error(f"Errore nel file {up.name}: {e}")
    if not rows:
        st.warning("Nessuna evidenziazione o commento trovati nei documenti caricati.")
    else:
        df = pd.DataFrame(rows)

        st.sidebar.header("Filtri")
        filenames = sorted(df["filename"].dropna().unique().tolist())
        sel_files = st.sidebar.multiselect("File", filenames, default=filenames)

        types = sorted(df["type"].dropna().unique().tolist())
        sel_types = st.sidebar.multiselect("Tipo annotazione", types, default=types)

        colors = sorted(df.loc[df["type"]=="highlight","highlight_color"].dropna().unique().tolist())
        sel_colors = st.sidebar.multiselect("Colore evidenziazione", colors, default=colors)

        authors = sorted(df.loc[df["type"]=="comment","author"].dropna().unique().tolist())
        sel_authors = st.sidebar.multiselect("Autore commento", authors, default=authors)

        codes = sorted(df.loc[df["type"]=="comment","code"].dropna().unique().tolist())
        sel_codes = st.sidebar.multiselect("Codice (dal commento)", codes, default=codes)

        labels = sorted(df.loc[df["type"]=="comment","label"].dropna().unique().tolist())
        sel_labels = st.sidebar.multiselect("Etichetta (dal commento)", labels, default=labels)

        search_text = st.sidebar.text_input("Cerca testo (qualsiasi campo)")
        date_min = st.sidebar.text_input("Data minima (YYYY-MM-DD)")
        date_max = st.sidebar.text_input("Data massima (YYYY-MM-DD)")

        fdf = df.copy()
        if sel_files:
            fdf = fdf[fdf["filename"].isin(sel_files)]
        if sel_types:
            fdf = fdf[fdf["type"].isin(sel_types)]
        if sel_colors:
            fdf = fdf[(fdf["type"]!="highlight") | (fdf["highlight_color"].isin(sel_colors))]
        if sel_authors:
            fdf = fdf[(fdf["type"]!="comment") | (fdf["author"].isin(sel_authors))]
        if sel_codes:
            fdf = fdf[(fdf["type"]!="comment") | (fdf["code"].isin(sel_codes))]
        if sel_labels:
            fdf = fdf[(fdf["type"]!="comment") | (fdf["label"].isin(sel_labels))]
        if search_text:
            mask = pd.Series(False, index=fdf.index)
            for col in ["extracted_text","paragraph","context","comment_text","quoted_text","author","filename","code","label"]:
                if col in fdf.columns:
                    mask = mask | fdf[col].astype(str).str.contains(search_text, case=False, na=False)
            fdf = fdf[mask]

        def parse_date(val):
            if not val:
                return None
            try:
                return datetime.fromisoformat(val.replace("Z","").replace("+00:00",""))
            except:
                for fmt in ("%Y-%m-%d", "%Y-%m-%dT%H:%M:%S"):
                    try:
                        return datetime.strptime(val[:len(fmt)], fmt)
                    except:
                        pass
            return None
        dmin = parse_date(date_min) if date_min else None
        dmax = parse_date(date_max) if date_max else None
        if dmin or dmax:
            if "date" in fdf.columns:
                dt_series = fdf["date"].apply(parse_date)
                if dmin:
                    fdf = fdf[(dt_series.isna()) | (dt_series >= dmin)]
                if dmax:
                    fdf = fdf[(dt_series.isna()) | (dt_series <= dmax)]

        high_df = fdf[fdf["type"]=="highlight"].copy()
        com_df = fdf[fdf["type"]=="comment"].copy()

        # Build linked table per pair
        link_rows = []
        if not com_df.empty:
            all_high = df[df["type"]=="highlight"].copy()
            for _, row in com_df.iterrows():
                quoted = str(row.get("quoted_text") or "").strip()
                fname = row.get("filename")
                matches = []
                if quoted:
                    cand = all_high[(all_high["filename"]==fname)]
                    for _, hr in cand.iterrows():
                        htext = str(hr.get("extracted_text") or "").strip()
                        if not htext:
                            continue
                        qn = re.sub(r"\s+", " ", quoted).strip()
                        hn = re.sub(r"\s+", " ", htext).strip()
                        if hn and qn and (hn in qn or qn in hn):
                            matches.append(htext)
                link_rows.append({
                    "filename": fname,
                    "comment_id": row.get("comment_id"),
                    "author": row.get("author"),
                    "date": row.get("date"),
                    "code": row.get("code"),
                    "label": row.get("label"),
                    "quoted_text": quoted,
                    "highlight_matches": len(matches),
                    "highlights_concat": " | ".join(sorted(set(matches))) if matches else ""
                })
        link_df = pd.DataFrame(link_rows) if link_rows else pd.DataFrame()

        st.subheader("Annotazioni collegate (Commento ↔ Evidenziati)")
        if not link_df.empty:
            default_link_cols = ["filename","comment_id","author","date","code","label","quoted_text","highlight_matches","highlights_concat"]
            link_cols = [c for c in default_link_cols if c in link_df.columns]
            sel_link_cols = st.multiselect("Colonne da mostrare (annotazioni collegate)", link_df.columns.tolist(), default=link_cols, key="link_cols")
            st.dataframe(link_df[sel_link_cols], use_container_width=True, hide_index=True)
        else:
            st.info("Nessuna annotazione collegata dopo i filtri.")

        st.subheader("Evidenziati")
        if not high_df.empty:
            default_high_cols = ["filename","extracted_text","highlight_color","context","paragraph"]
            high_cols = [c for c in default_high_cols if c in high_df.columns]
            sel_high_cols = st.multiselect("Colonne da mostrare (evidenziati)", high_df.columns.tolist(), default=high_cols, key="high_cols")
            st.dataframe(high_df[sel_high_cols], use_container_width=True, hide_index=True)
        else:
            st.info("Nessun evidenziato dopo i filtri.")

        st.subheader("Commenti (esplosi per codice/etichetta)")
        if not com_df.empty:
            default_com_cols = ["filename","comment_id","author","date","quoted_text","comment_text","code","label"]
            com_cols = [c for c in default_com_cols if c in com_df.columns]
            sel_com_cols = st.multiselect("Colonne da mostrare (commenti)", com_df.columns.tolist(), default=com_cols, key="com_cols")
            st.dataframe(com_df[sel_com_cols], use_container_width=True, hide_index=True)
        else:
            st.info("Nessun commento dopo i filtri.")

        st.divider()
        st.subheader("Esportazione")
        excel_bytes = export_excel(high_df, com_df, link_df)
        st.download_button(
            label="Scarica Excel filtrato",
            data=excel_bytes,
            file_name="estrazioni_filtrate.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        if not link_df.empty:
            st.download_button(
                label="Scarica CSV Annotazioni collegate",
                data=link_df.to_csv(index=False).encode("utf-8"),
                file_name="annotazioni_collegate.csv",
                mime="text/csv"
            )
        if not high_df.empty:
            st.download_button(
                label="Scarica CSV Evidenziati",
                data=high_df.to_csv(index=False).encode("utf-8"),
                file_name="evidenziati_filtrati.csv",
                mime="text/csv"
            )
        if not com_df.empty:
            st.download_button(
                label="Scarica CSV Commenti (esplosi)",
                data=com_df.to_csv(index=False).encode("utf-8"),
                file_name="commenti_esplosi.csv",
                mime="text/csv"
            )
else:
    st.info("Carica uno o più file .docx per iniziare.")
