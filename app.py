# --- bootstrap: ensure pygooglenews is available without its legacy deps ---
import sys, subprocess
def _pip(*args):
    subprocess.check_call([sys.executable, "-m", "pip", *args])

try:
    import pygooglenews  # noqa: F401
except Exception:
    _pip("install", "pygooglenews==0.1.2", "--no-deps")
# ---------------------------------------------------------------------------

import re
from datetime import datetime
from io import BytesIO

import streamlit as st
from pygooglenews import GoogleNews

from docx import Document
from docx.shared import Pt
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import docx.opc.constants as constants

# -----------------------------
# Config & constants
# -----------------------------
DEFAULT_TERMS = [
    "Jerry Moran", "Senator Jerry Moran", "Senator Moran",
    "Sen. Moran", "Sen. Jerry Moran", "Sens. Moran", "Sens. Jerry Moran"
]

KANSAS_OUTLETS = [
    'Kansas Reflector', 'The Topeka Capital-Journal', 'The Wichita Eagle',
    'KCLY Radio', 'KSN-TV', 'KWCH', 'Kansas City Star',
    'Lawrence Journal-World', 'The Garden City Telegram', 'KSNT 27 News',
    'The Hutchinson News', 'Salina Journal', 'Hays Daily News',
    'Hays Post', 'Emporia Gazette', 'JC Post', 'WIBW'
]

EXCLUDE_SOURCES_CONTAINS = ['.gov', 'Quiver Quantitative', 'MSN', 'Twin States News']

# -----------------------------
# Helpers
# -----------------------------
def clean_text(text):
    return re.sub(r'[^\w\s\-\.\,\:\;\!\?\(\)\'\"]+', '', text).strip()

def normalize_title_for_duplicate_detection(title):
    normalized = title.lower().strip()
    normalized = re.sub(r'^(breaking:?\s*|update:?\s*|exclusive:?\s*)', '', normalized)
    normalized = re.sub(r'\s+', ' ', normalized)
    return normalized

def add_hyperlink(paragraph, url, text):
    part = paragraph.part
    r_id = part.relate_to(url, constants.RELATIONSHIP_TYPE.HYPERLINK, is_external=True)
    hyperlink = OxmlElement('w:hyperlink')
    hyperlink.set(qn('r:id'), r_id)

    new_run = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')

    color = OxmlElement('w:color'); color.set(qn('w:val'), '0000FF'); rPr.append(color)
    u = OxmlElement('w:u'); u.set(qn('w:val'), 'single'); rPr.append(u)

    new_run.append(rPr)
    text_elem = OxmlElement('w:t'); text_elem.text = text; new_run.append(text_elem)

    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)
    return hyperlink

def process_entries_with_duplicates(all_entries):
    title_groups = {}
    for entry in all_entries:
        media = entry['source']['title']
        if any(x in media for x in EXCLUDE_SOURCES_CONTAINS):
            continue

        raw_title = re.sub(rf" - {re.escape(media)}$", "", entry['title'])
        title = clean_text(raw_title)
        normalized_title = normalize_title_for_duplicate_detection(title)

        title_groups.setdefault(normalized_title, []).append({
            'title': title, 'media': clean_text(media), 'link': entry['link'], 'entry': entry
        })

    processed = []
    for group in title_groups.values():
        if not group: continue
        primary, duplicates = group[0], group[1:]
        media_string = primary['media']
        if duplicates:
            dup_outlets = [d['media'] for d in duplicates]
            media_string += f" (also ran in {', '.join(dup_outlets[:-1])} and {dup_outlets[-1]})" if len(dup_outlets) > 1 else f" (also ran in {dup_outlets[0]})"
        is_kansas = any(k in primary['media'] for k in KANSAS_OUTLETS)
        processed.append({'title': primary['title'], 'media_string': media_string, 'link': primary['link'], 'is_kansas': is_kansas})
    return processed

def build_docx_bytes(processed_entries):
    now = datetime.now()
    filename = f"Sen Moran in the News {now.month}-{now.day}.docx"
    doc = Document()
    doc.add_paragraph(f"Generated on: {now.strftime('%B %d, %Y at %I:%M %p')}")
    doc.add_paragraph(f"Total articles found: {len(processed_entries)}")
    doc.add_paragraph("* indicates Kansas news outlet")
    doc.add_paragraph()
    doc.add_heading('News Articles', level=1)

    for entry in processed_entries:
        prefix = '*' if entry['is_kansas'] else ''
        p = doc.add_paragraph()
        p.add_run(f"{prefix}{entry['media_string']}: ")
        add_hyperlink(p, entry['link'], entry['title'])
        url_run = p.add_run(f" [{entry['link']}]"); url_run.font.italic = True; url_run.font.size = Pt(8)

    bio = BytesIO(); doc.save(bio); bio.seek(0)
    return filename, bio

# -----------------------------
# Streamlit UI
# -----------------------------
st.title("Jerry Moran News Search")
st.caption("Powered by Google News via pygooglenews")

terms_text = st.text_area("Search terms (comma-separated):", value=", ".join(DEFAULT_TERMS), height=90)
when_choice = st.selectbox("Time window", ["1d", "3d", "7d", "14d", "30d"], index=0, help="pygooglenews `when` parameter")

if st.button("Search"):
    search_terms = [t.strip() for t in terms_text.split(",") if t.strip()]
    gn = GoogleNews(lang='en', country='US')
    all_entries, seen_links = [], set()

    with st.spinner("Searching Google News…"):
        for term in search_terms:
            try:
                results = gn.search(term, when=when_choice)
                for e in results.get('entries', []):
                    if e['link'] not in seen_links:
                        all_entries.append(e); seen_links.add(e['link'])
            except Exception as ex:
                st.warning(f"Error for '{term}': {ex}")

    st.write(f"**Found {len(all_entries)} unique articles**")
    processed_entries = process_entries_with_duplicates(all_entries)

    md_lines = ["# Jerry Moran News", ""]
    for i, entry in enumerate(processed_entries, 1):
        star = "*" if entry['is_kansas'] else ""
        md_lines.append(f"{i}. {star}{entry['media_string']}: [{entry['title']}]({entry['link']})")
    st.markdown("\n".join(md_lines))

    filename, bio = build_docx_bytes(processed_entries)
    st.download_button("Download Word Document", data=bio, file_name=filename, mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
