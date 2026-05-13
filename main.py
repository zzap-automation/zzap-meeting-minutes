import os
import re
import tempfile
import uuid
from datetime import datetime
from io import BytesIO

import httpx
from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import Response, StreamingResponse

# ── optional Supabase ────────────────────────────────────────────────────────
try:
    from supabase import create_client
    _sb_url = os.environ.get("SUPABASE_URL", "")
    _sb_key = os.environ.get("SUPABASE_KEY", "")
    supabase = create_client(_sb_url, _sb_key) if _sb_url and _sb_key else None
except Exception:
    supabase = None

# ── optional python-docx ─────────────────────────────────────────────────────
try:
    from docx import Document as DocxDoc
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Pt, RGBColor
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False

# ────────────────────────────────────────────────────────────────────────────
GROQ_API_KEY = os.environ.get("GROQ_API_KEY", "")
GROQ_BASE    = "https://api.groq.com/openai/v1"
LLM_MODEL    = "llama-3.3-70b-versatile"
STT_MODEL    = "whisper-large-v3-turbo"

SYSTEM_PROMPT = """You are a professional meeting minutes writer for an architecture and planning firm.
Convert meeting transcripts into structured minutes using EXACTLY the format below.

# Meeting Minutes - [Concise Descriptive Title]

**Date:** [Extract from transcript, or "Not specified"]
**Attendees:** [Full names comma-separated, identified from transcript]
**Duration:** [Approximate, inferred from context, or "Not specified"]

---

## Meeting Discussion

| Agenda | Item | Description |
|--------|------|-------------|
| [Topic area — 1-4 words, repeated across rows sharing same broad topic] | [Specific sub-topic — 3-7 words] | [2-4 sentence factual prose. Third person. No bullet points inside cells.] |

---

## Action Items Summary

| Action Item | Details | Owner | Timeline |
|-------------|---------|-------|----------|
| [Short action title] | [What needs to be done] | [Person responsible] | [e.g. "This week", "ASAP", "TBD"] |

---

## Key Decisions

1. [Definitively agreed items only — not items to be discussed later]

---

## Additional Notes

- [Important context, background, or observations not captured in the tables]

RULES:
- Agenda: 1-4 words. Repeat across rows sharing the same broad topic.
- Item: 3-7 words, specific and scannable.
- Description: factual prose, complete sentences, third person voice.
- Action Items: every commitment or follow-up. Be specific on who, what, when.
- Key Decisions: definitively agreed items only.
- Additional Notes: context helping readers understand without having attended.
- Produce ALL five sections. Use "None identified." if a section has no content.
- Speaker labels like "Ife:" or "Stephen:" identify attendees — use them for ownership and names."""

# ── app ──────────────────────────────────────────────────────────────────────
app = FastAPI(title="zzap Meeting Minutes API")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)


# ── helpers ──────────────────────────────────────────────────────────────────
def _groq_headers():
    if not GROQ_API_KEY:
        raise HTTPException(500, "GROQ_API_KEY is not configured.")
    return {"Authorization": f"Bearer {GROQ_API_KEY}"}


def _clean_md(text: str) -> str:
    """Strip accidental code fences from LLM output."""
    return re.sub(r"^```(?:markdown)?\n?", "", text.strip()).rstrip("`").strip()


# ── routes ───────────────────────────────────────────────────────────────────

@app.get("/health")
async def health():
    return {"status": "ok", "docx_available": DOCX_AVAILABLE, "supabase": supabase is not None}


@app.post("/transcribe")
async def transcribe(audio: UploadFile = File(...)):
    """Transcribe an uploaded audio file using Groq Whisper."""
    data = await audio.read()
    if len(data) > 25 * 1024 * 1024:
        raise HTTPException(413, "File exceeds 25 MB free-tier limit. Compress or trim the recording.")

    ext = os.path.splitext(audio.filename or "audio.mp3")[1] or ".mp3"
    async with httpx.AsyncClient(timeout=180) as client:
        resp = await client.post(
            f"{GROQ_BASE}/audio/transcriptions",
            headers=_groq_headers(),
            files={"file": (f"audio{ext}", data, audio.content_type or "audio/mpeg")},
            data={"model": STT_MODEL, "response_format": "text"},
        )
    if not resp.is_success:
        raise HTTPException(resp.status_code, f"Groq STT error: {resp.text}")
    return {"transcript": resp.text.strip()}


@app.post("/generate")
async def generate(
    transcript: str = Form(...),
    title: str = Form(""),
    attendees: str = Form(""),
    date: str = Form(""),
):
    """Generate formatted minutes from a transcript using Groq LLM."""
    user_msg = "Please produce meeting minutes from this transcript.\n"
    if title:     user_msg += f"\nMeeting title: {title}"
    if attendees: user_msg += f"\nAttendees: {attendees}"
    if date:      user_msg += f"\nDate: {date}"
    user_msg += f"\n\nTRANSCRIPT:\n{transcript}"

    async with httpx.AsyncClient(timeout=180) as client:
        resp = await client.post(
            f"{GROQ_BASE}/chat/completions",
            headers={**_groq_headers(), "Content-Type": "application/json"},
            json={
                "model": LLM_MODEL,
                "messages": [
                    {"role": "system", "content": SYSTEM_PROMPT},
                    {"role": "user",   "content": user_msg},
                ],
                "max_tokens": 4000,
                "temperature": 0.2,
            },
        )
    if not resp.is_success:
        raise HTTPException(resp.status_code, f"Groq LLM error: {resp.text}")
    minutes_md = _clean_md(resp.json()["choices"][0]["message"]["content"])
    return {"minutes_md": minutes_md}


@app.post("/save")
async def save_minutes(
    title: str      = Form(...),
    date: str       = Form(""),
    attendees: str  = Form(""),
    transcript: str = Form(""),
    minutes_md: str = Form(...),
):
    """Persist minutes to Supabase."""
    if supabase is None:
        raise HTTPException(503, "Supabase is not configured on this deployment.")
    result = supabase.table("meetings").insert({
        "title":      title,
        "date":       date,
        "attendees":  attendees,
        "transcript": transcript,
        "minutes_md": minutes_md,
    }).execute()
    return {"id": result.data[0]["id"]}


@app.get("/minutes")
async def list_minutes():
    """Return a summary list of all archived meetings."""
    if supabase is None:
        return {"meetings": []}
    result = (
        supabase.table("meetings")
        .select("id,title,date,attendees,created_at")
        .order("created_at", desc=True)
        .execute()
    )
    return {"meetings": result.data}


@app.get("/minutes/{meeting_id}")
async def get_minutes(meeting_id: str):
    """Return full record for one meeting."""
    if supabase is None:
        raise HTTPException(503, "Supabase is not configured.")
    result = supabase.table("meetings").select("*").eq("id", meeting_id).execute()
    if not result.data:
        raise HTTPException(404, "Meeting not found.")
    return result.data[0]


@app.post("/export/docx")
async def export_docx(minutes_md: str = Form(...), filename: str = Form("meeting-minutes")):
    """Convert markdown minutes to a formatted .docx file."""
    if not DOCX_AVAILABLE:
        raise HTTPException(503, "python-docx is not installed on this server.")

    buf = _md_to_docx(minutes_md)
    safe_name = re.sub(r"[^\w\-]", "-", filename).strip("-") or "meeting-minutes"
    return Response(
        content=buf.getvalue(),
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={"Content-Disposition": f'attachment; filename="{safe_name}.docx"'},
    )


# ── markdown → docx converter ────────────────────────────────────────────────

def _set_cell_shading(cell, fill_hex: str):
    """Apply background shading to a table cell."""
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd  = OxmlElement("w:shd")
    shd.set(qn("w:val"),   "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"),  fill_hex)
    tcPr.append(shd)


def _add_bottom_border(paragraph, color_hex: str = "CCCCCC"):
    """Draw a thin bottom rule under a paragraph (used after section headers)."""
    pPr  = paragraph._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bot  = OxmlElement("w:bottom")
    bot.set(qn("w:val"),   "single")
    bot.set(qn("w:sz"),    "4")
    bot.set(qn("w:space"), "1")
    bot.set(qn("w:color"), color_hex)
    pBdr.append(bot)
    pPr.append(pBdr)


def _parse_inline(text: str):
    """Return list of (text, bold) tuples from a markdown inline string."""
    parts, buf, bold = [], "", False
    i = 0
    while i < len(text):
        if text[i:i+2] == "**":
            parts.append((buf, bold)); buf = ""; bold = not bold; i += 2
        else:
            buf += text[i]; i += 1
    parts.append((buf, bold))
    return [(t, b) for t, b in parts if t]


def _add_paragraph_runs(para, text: str, base_bold=False, base_size_pt=11):
    for chunk, chunk_bold in _parse_inline(text):
        run = para.add_run(chunk)
        run.bold = base_bold or chunk_bold
        run.font.size = Pt(base_size_pt)


def _md_to_docx(md: str) -> BytesIO:
    if not DOCX_AVAILABLE:
        raise RuntimeError("python-docx not available")

    doc = DocxDoc()

    # ── page margins ──────────────────────────────────────────────────────
    for section in doc.sections:
        section.top_margin    = Pt(72)   # 1 inch
        section.bottom_margin = Pt(72)
        section.left_margin   = Pt(72)
        section.right_margin  = Pt(72)

    # ── default style ─────────────────────────────────────────────────────
    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(11)

    lines     = md.split("\n")
    i         = 0
    in_table  = False

    # We collect table rows then flush them to a real docx table
    table_rows: list[list[str]] = []

    def flush_table():
        nonlocal in_table, table_rows
        if not table_rows:
            return
        col_count = max(len(r) for r in table_rows)
        # Determine column widths as fractions based on header
        # Content width ~9360 twips (US letter, 1" margins)
        # Three-col discussion: 20% / 25% / 55%
        # Four-col action items: 20% / 35% / 20% / 25%
        if col_count == 3:
            widths = [1800, 2200, 5360]
        elif col_count == 4:
            widths = [1800, 3200, 1800, 2560]
        else:
            each = 9360 // col_count
            widths = [each] * col_count

        from docx.shared import Pt as _Pt
        from docx.oxml.ns import qn as _qn
        from docx.oxml import OxmlElement as _el

        tbl = doc.add_table(rows=0, cols=col_count)
        tbl.style = "Table Grid"

        for row_idx, cells in enumerate(table_rows):
            row = tbl.add_row()
            for col_idx in range(col_count):
                cell = row.cells[col_idx]
                text = cells[col_idx] if col_idx < len(cells) else ""
                # Set width
                tc = cell._tc
                tcPr = tc.get_or_add_tcPr()
                tcW = _el("w:tcW")
                tcW.set(_qn("w:w"),    str(widths[col_idx]))
                tcW.set(_qn("w:type"), "dxa")
                tcPr.append(tcW)
                # Header row shading
                if row_idx == 0:
                    _set_cell_shading(cell, "1565a8")
                    para = cell.paragraphs[0]
                    run  = para.add_run(text)
                    run.bold       = True
                    run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                    run.font.size  = Pt(10)
                else:
                    para = cell.paragraphs[0]
                    _add_paragraph_runs(para, text, base_size_pt=10)
                    if row_idx % 2 == 0:
                        _set_cell_shading(cell, "F5F5F5")

        doc.add_paragraph()  # spacing after table
        in_table   = False
        table_rows.clear()

    while i < len(lines):
        line = lines[i]

        # ── H1 title ──────────────────────────────────────────────────────
        if line.startswith("# "):
            flush_table()
            text = line[2:].strip()
            p    = doc.add_paragraph()
            run  = p.add_run(text)
            run.bold           = True
            run.font.size      = Pt(18)
            run.font.color.rgb = RGBColor(0x15, 0x65, 0xa8)  # dark navy
            p.paragraph_format.space_after = Pt(4)
            i += 1; continue

        # ── H2 section header ─────────────────────────────────────────────
        if line.startswith("## "):
            flush_table()
            text = line[3:].strip()
            p    = doc.add_paragraph()
            run  = p.add_run(text.upper())
            run.bold           = True
            run.font.size      = Pt(10)
            run.font.color.rgb = RGBColor(0x15, 0x65, 0xa8)
            p.paragraph_format.space_before = Pt(12)
            p.paragraph_format.space_after  = Pt(2)
            _add_bottom_border(p, "1565a8")
            i += 1; continue

        # ── horizontal rule ───────────────────────────────────────────────
        if line.strip() in ("---", "***", "___"):
            flush_table()
            i += 1; continue

        # ── meta lines (**Date:** etc.) ───────────────────────────────────
        if line.startswith("**") and ":**" in line:
            flush_table()
            p = doc.add_paragraph()
            p.paragraph_format.space_after = Pt(0)
            _add_paragraph_runs(p, line)
            i += 1; continue

        # ── table row ─────────────────────────────────────────────────────
        if line.strip().startswith("|"):
            in_table = True
            cells = [c.strip() for c in line.strip().strip("|").split("|")]
            # skip separator rows
            if all(re.match(r"^[-: ]+$", c) for c in cells if c):
                i += 1; continue
            table_rows.append(cells)
            i += 1; continue
        elif in_table:
            flush_table()

        # ── numbered list ─────────────────────────────────────────────────
        m = re.match(r"^\d+\.\s+(.*)", line)
        if m:
            flush_table()
            p    = doc.add_paragraph(style="List Number")
            _add_paragraph_runs(p, m.group(1))
            p.paragraph_format.space_after = Pt(2)
            i += 1; continue

        # ── bullet list ───────────────────────────────────────────────────
        if re.match(r"^[-*]\s+", line):
            flush_table()
            text = re.sub(r"^[-*]\s+", "", line)
            p    = doc.add_paragraph(style="List Bullet")
            _add_paragraph_runs(p, text)
            p.paragraph_format.space_after = Pt(2)
            i += 1; continue

        # ── blank line ────────────────────────────────────────────────────
        if not line.strip():
            flush_table()
            i += 1; continue

        # ── plain paragraph ───────────────────────────────────────────────
        flush_table()
        p = doc.add_paragraph()
        _add_paragraph_runs(p, line)
        i += 1

    flush_table()

    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf
