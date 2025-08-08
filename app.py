import re
from datetime import datetime
from io import BytesIO
from urllib.parse import quote_plus

import streamlit as st
import feedparser

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
# Google News RSS helper
# -----------------------------
def google_news_rss(term: str, when: str = "1d", lang="en-US", country="US"):
    """
    Build a Google News RSS query.
    - 'when' uses the Google News query operator (e.g., when:1d, when:7d).
    """
    query = f"{term} when:{when}"
    q = quote_plus(query)
    # hl=en-US, gl=US, ceid=US:en controls locale
    return f"https://news.google.com/rss/search?q={q}&hl={lang}&gl={country}&ceid={country}:en"

def fetch_entries(search_terms, when="1d"):
    """Fetch and de-duplicate entries across search terms using Google News RSS + feedparser."""
    all_entries, seen_links = [], set()
    for term in search_terms:
        url = google_news_rss(term, when=when)
        feed = feedparser.parse(url)
        for e in feed.entries:
            link = e.get("link")
            if not link or link in seen_links:
                continue
            seen_links.add(link)

            # Try to read the source/outlet
            media = None
            # feedparser commonly exposes source.title
            src = getattr(e, "source", None)
            if src and isinstance(src, dict):
                media = src.get("title")

            # Fallbacks
            if not media:
                media = e.get("author") or "Unknown"

            title = e.get("title", "").strip()
            all_entries.append({
                "title": title,
                "link": link,
                "source": {"title": media}
            })
    return all_entries

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
        media = entry['source']['title'] or "Unknown"

        if any(x in media for x in EXCLUDE_SOURCES_CONTAINS):
            continue

        raw_title = re.sub(rf" - {re.escape(media)}$", "", entry['title'] or "")
        title = clean_text(raw_title)
        normalized_title = normalize_title_for_duplicate_detection(title)

        title_groups.setdefault(normalized_title, []).append({
            'title': title, 'media': clean_text(media), 'link': entry['link'], 'entry': entry
        })

    processed = []
    for group in title_groups.values():
        if not group: 
            continue
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

# --- page config & light CSS ---
st.set_page_config(
    page_title="Jerry Moran News Search",
    page_icon="📰",
    layout="wide",
)

st.markdown("""
<style>
/* tighten list spacing and make links pop a bit */
.block-container { padding-top: 1.5rem; }
h1.hero { font-size: 2rem; margin: 0 0 .25rem 0; }
p.sub { color:#374151; margin:0 0 1rem 0; }
.badge { display:inline-block; padding:.15rem .5rem; font-size:.75rem; border-radius:999px; background:#e8ecff; color:#1f3a8a; margin-right:.4rem; }
hr.soft { border:none; border-top:1px solid #e5e7eb; margin:1rem 0 1.25rem 0; }
ol li { margin-bottom:.35rem; }
a { text-decoration: underline; }
</style>
""", unsafe_allow_html=True)

# --- PAGE HEADER ---
st.set_page_config(
    page_title="Jerry Moran News Search",
    page_icon="📰",
    layout="wide",
)

# Full-width centered image at top
st.image(
    "assets/jerry-moran.jpg",  # replace with your local file path
    use_column_width=True
)

st.markdown(
    "<h1 style='text-align:center;'>Jerry Moran — News Tracker</h1>",
    unsafe_allow_html=True
)
st.markdown(
    "<p style='text-align:center;'>Live Google News RSS search with smart deduping, Kansas-outlet highlighting, and one-click DOCX export.</p>",
    unsafe_allow_html=True
)
st.markdown("<hr>", unsafe_allow_html=True)

# --- SEARCH CONTROLS CENTERED ---
col1, col2, col3 = st.columns([1, 2, 1])
with col2:
    terms_text = st.text_area(
        "Search terms (comma-separated)",
        value=", ".join(DEFAULT_TERMS),
        height=110
    )
    when_choice = st.selectbox(
        "Time window",
        options=["1d", "3d", "7d", "14d", "30d"],
        index=0,
        help="Google News query operator `when:`"
    )
    run_search = st.button("🔎 Run Search", type="primary", use_container_width=True)

# --- RESULTS ---
if run_search:
    search_terms = [t.strip() for t in terms_text.split(",") if t.strip()]
    with st.spinner("Searching Google News…"):
        all_entries = fetch_entries(search_terms, when=when_choice)

    processed_entries = process_entries_with_duplicates(all_entries)

    st.markdown(
        f"<p style='text-align:center;'><strong>Found {len(all_entries)} items before dedupe — After dedupe: {len(processed_entries)}</strong></p>",
        unsafe_allow_html=True
    )

    # Center results
    col1, col2, col3 = st.columns([0.5, 3, 0.5])
    with col2:
        md_lines = []
        for i, entry in enumerate(processed_entries, 1):
            star = "*" if entry['is_kansas'] else ""
            md_lines.append(f"{i}. {star}{entry['media_string']}: [{entry['title']}]({entry['link']})")
        st.markdown("\n".join(md_lines))

        # Download DOCX
        filename, bio = build_docx_bytes(processed_entries)
        st.download_button(
            "⬇️ Download Word Document",
            data=bio,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
else:
    st.info("Enter search terms above and click **Run Search**.")
