# =============================================================================
# DIUD — Decision Intelligence Using Data
# FastAPI backend: Claude + ClickHouse HTTP proxy + PDF/PPTX export
#
# Your ClickHouse proxy (OpenAPI spec confirmed):
#   POST /query   body: {"query": "<SQL>", "limit": <int|null>}
#   Authorization: Bearer <token>
#
# ENV variables needed in .env:
#   ANTHROPIC_API_KEY=sk-ant-...
#   CLICKHOUSE_API_URL=https://clickhouse-api-j55l.onrender.com
#   CLICKHOUSE_API_TOKEN=your_bearer_token_here
# =============================================================================

import io
import json
import os
import re
import traceback
from datetime import date
from typing import List, Literal

import httpx
import anthropic
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import HTMLResponse, StreamingResponse
from pydantic import BaseModel
from dotenv import load_dotenv

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import (
    BaseDocTemplate,
    Frame,
    PageTemplate,
    PageBreak,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
)

# =============================================================================
# Load ENV
# =============================================================================

load_dotenv()

# =============================================================================
# FastAPI App
# =============================================================================

app = FastAPI(title="DIUD", description="Decision Intelligence Using Data", version="1.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# =============================================================================
# Claude client
# =============================================================================

_ai_client = anthropic.Anthropic(api_key=os.getenv("ANTHROPIC_API_KEY"))
_CLAUDE_MODEL = "claude-sonnet-4-5"

# =============================================================================
# ClickHouse HTTP proxy connector
#
# Confirmed from your OpenAPI spec:
#   POST /query
#   Request body (QueryRequest schema):
#     { "query": "<SQL string>", "limit": <integer or null> }
#   Auth: Bearer token via HTTPBearer
#
# Other available endpoints (no query body needed):
#   GET  /databases              → list all databases
#   GET  /tables/{database}      → list tables in a database
#   GET  /schema/{database}/{table} → get column schema for a table
# =============================================================================

def _base_url() -> str:
    return (os.getenv("CLICKHOUSE_API_URL") or "").rstrip("/")

def _token() -> str:
    return os.getenv("CLICKHOUSE_API_TOKEN") or ""

def _auth_headers() -> dict:
    return {
        "Authorization": f"Bearer {_token()}",
        "Content-Type": "application/json",
    }

FORBIDDEN_KEYWORDS = ["INSERT", "UPDATE", "DELETE", "DROP", "ALTER", "TRUNCATE", "CREATE"]


def run_clickhouse_query(sql: str) -> str:
    """
    Execute a SELECT/WITH query via POST /query on your ClickHouse HTTP proxy.
    Body: {"query": "<SQL>"}   (field name is "query", NOT "sql" — confirmed from OpenAPI spec)
    Returns a plain-text table string, or an error message starting with ERROR:/DATABASE.
    Never raises — always returns a string so Claude can report failures gracefully.
    """
    base_url = _base_url()
    token    = _token()

    if not base_url:
        return (
            "DATABASE CONNECTION FAILED: CLICKHOUSE_API_URL is not set. "
            "Add it to your .env file and restart. Example:\n"
            "CLICKHOUSE_API_URL=https://clickhouse-api-j55l.onrender.com"
        )
    if not token:
        return (
            "DATABASE CONNECTION FAILED: CLICKHOUSE_API_TOKEN is not set. "
            "Add your Bearer token to .env and restart."
        )

    # Safety check — only allow read queries
    stripped = sql.strip().upper()
    if not (stripped.startswith("SELECT") or stripped.startswith("WITH")):
        return "ERROR: Only SELECT/WITH queries are permitted."
    for kw in FORBIDDEN_KEYWORDS:
        if re.search(rf'\b{kw}\b', stripped):
            return f"ERROR: Forbidden SQL keyword detected: {kw}"

    print(f"🔍 SQL → POST {base_url}/query")
    print(f"   {sql[:300]}")

    try:
        resp = httpx.post(
            f"{base_url}/query",
            headers=_auth_headers(),
            # ⚠️  Field name MUST be "query" per your OpenAPI QueryRequest schema
            json={"query": sql},
            timeout=30,
        )

        # Handle HTTP errors clearly
        if resp.status_code == 401:
            return "DATABASE CONNECTION FAILED: 401 Unauthorized — check your CLICKHOUSE_API_TOKEN."
        if resp.status_code == 403:
            return "DATABASE CONNECTION FAILED: 403 Forbidden — token may not have read permission."
        if resp.status_code == 422:
            # Validation error from the proxy — log detail for debugging
            detail = resp.text[:500]
            print(f"   422 detail: {detail}")
            return f"ERROR: Query was rejected by the proxy (422 Unprocessable Entity). Detail: {detail}"
        if resp.status_code != 200:
            return f"DATABASE ERROR: HTTP {resp.status_code} — {resp.text[:400]}"

        # Parse the response
        payload = resp.json()
        print(f"   Response type: {type(payload).__name__}, keys: {list(payload.keys()) if isinstance(payload, dict) else 'list'}")

        # Handle different response shapes the proxy might return
        if isinstance(payload, list):
            # Direct list of row dicts
            rows = payload
        elif isinstance(payload, dict):
            # Try common wrapper keys
            rows = (
                payload.get("data")
                or payload.get("rows")
                or payload.get("result")
                or payload.get("results")
                or None
            )
            if rows is None:
                # Unknown shape — return raw so Claude can interpret
                return json.dumps(payload, indent=2, default=str)[:3000]
        else:
            return f"Unexpected response type: {type(payload)}"

        if not rows:
            return "Query returned 0 rows."

        # Build a readable text table
        if isinstance(rows[0], dict):
            cols   = list(rows[0].keys())
            header = " | ".join(cols)
            sep    = "-" * min(len(header), 140)
            lines  = [header, sep]
            for row in rows[:100]:
                lines.append(" | ".join(str(row.get(c, "NULL")) for c in cols))
        else:
            # List of lists
            lines = [" | ".join(str(v) for v in rows[0]), "-" * 80]
            for row in rows[1:101]:
                lines.append(" | ".join(str(v) for v in row))

        if len(rows) > 100:
            lines.append(f"... ({len(rows) - 100} more rows not shown)")

        result = "\n".join(lines)
        print(f"   ✅ {len(rows)} rows returned. Preview: {result[:200]}")
        return result

    except httpx.ConnectError as e:
        return (
            f"DATABASE CONNECTION FAILED: Could not reach {base_url}. "
            f"Check that the URL is correct and the service is running.\nDetail: {e}"
        )
    except httpx.TimeoutException:
        return "DATABASE CONNECTION FAILED: Query timed out after 30 seconds."
    except Exception as exc:
        traceback.print_exc()
        return f"DATABASE CONNECTION FAILED: {type(exc).__name__}: {exc}"


# =============================================================================
# Helper: introspect schema via the proxy's /schema endpoint
# =============================================================================

def get_table_schema(database: str, table: str) -> dict:
    """Fetch column definitions from GET /schema/{database}/{table}."""
    base_url = _base_url()
    token    = _token()
    if not base_url or not token:
        return {"error": "API URL or token not configured"}
    try:
        resp = httpx.get(
            f"{base_url}/schema/{database}/{table}",
            headers=_auth_headers(),
            timeout=15,
        )
        if resp.status_code == 200:
            return resp.json()
        return {"error": f"HTTP {resp.status_code}: {resp.text[:300]}"}
    except Exception as e:
        return {"error": str(e)}


# =============================================================================
# System Prompt
# =============================================================================

_SYSTEM_PROMPT = """
You are DIUD (Decision Intelligence Using Data) — a conversational data assistant.
You have LIVE access to a ClickHouse database via the query_clickhouse tool.

CORE RULES:
- NEVER say you lack database access. You always have it via the tool.
- NEVER fabricate numbers. Query the DB for every metric question.
- NEVER run destructive SQL (INSERT / UPDATE / DELETE / DROP / ALTER / TRUNCATE).
- If a query fails or returns an error, report the exact error message to the user.
- Answer in clean markdown: use tables for data, bold for key numbers.
- Be concise but complete.
- When generating export content, use ## section headers so the PDF/PPTX renderer can parse them.

=================================================================
DATABASE ACCESS
=================================================================

Tool: query_clickhouse
Use it for any question about deals, pipeline, metrics, win/loss, or any data.
If it returns DATABASE CONNECTION FAILED, relay that message and tell the user
to open /debug/db to diagnose connectivity.

=================================================================
TABLE SCHEMA — deals (only table in scope for now)
=================================================================

TABLE: deals
One row per deal / sales opportunity.

COLUMNS:
  id              INTEGER   — unique deal identifier
  name            STRING    — deal / opportunity name
  stage           STRING    — current pipeline stage (see STAGE LIST below)
  owner           STRING    — AE / deal owner name
  amount          FLOAT     — deal value in USD
  region          STRING    — geographic region
  source          STRING    — lead source
  industry        STRING    — customer industry vertical
  close_date      DATE      — expected or actual close date
  created_date    DATE      — date deal was created
  status          STRING    — 'open', 'won', 'lost'
  lost_reason     STRING    — reason for loss (NULL if not lost)
  won_reason      STRING    — reason for win (NULL if not won)
  competitor      STRING    — competitor involved (NULL if none)
  notes           STRING    — freeform notes

⚠️  Replace these columns with your REAL deals table column names before deploying.
    You can fetch the actual schema using GET /schema/{database}/{table} on the proxy.

=================================================================
STAGE LIST
=================================================================
  'Prospecting'
  'Qualification'
  'Discovery'
  'Proposal'
  'Negotiation'
  'Closed Won'
  'Closed Lost'

=================================================================
QUERY RULES
=================================================================
1. SELECT / WITH only — no INSERT, UPDATE, DELETE, DROP, ALTER, TRUNCATE
2. LIMIT row-level queries to 100 rows maximum
3. Use COUNT(DISTINCT id) for unique deal counts
4. For USD totals: ROUND(SUM(amount)/1e6, 1) gives $M figures
5. Case-insensitive matching: use LOWER(col) = LOWER('value')  or ILIKE
6. NULLs in display: COALESCE(col, 'Unknown')

=================================================================
BUSINESS DEFINITIONS
=================================================================
"Active pipeline"  → status = 'open'
"Won deals"        → status = 'won'   OR stage = 'Closed Won'
"Lost deals"       → status = 'lost'  OR stage = 'Closed Lost'
"Win rate"         → won / (won + lost) * 100
"Pipeline value"   → SUM(amount) WHERE status = 'open'

=================================================================
SAMPLE QUERIES
=================================================================

-- Deals by stage:
SELECT stage, COUNT(DISTINCT id) AS deals, ROUND(SUM(amount)/1e6,2) AS value_m
FROM deals
GROUP BY stage ORDER BY value_m DESC

-- Top 10 open deals:
SELECT name, owner, stage, amount, close_date
FROM deals WHERE status = 'open'
ORDER BY amount DESC LIMIT 10

-- Win rate by region:
SELECT
  region,
  COUNT(DISTINCT CASE WHEN status='won'  THEN id END) AS won,
  COUNT(DISTINCT CASE WHEN status='lost' THEN id END) AS lost,
  ROUND(
    COUNT(DISTINCT CASE WHEN status='won' THEN id END) * 100.0
    / NULLIF(COUNT(DISTINCT CASE WHEN status IN ('won','lost') THEN id END), 0)
  , 1) AS win_rate_pct
FROM deals
GROUP BY region ORDER BY won DESC

-- Pipeline by industry:
SELECT COALESCE(industry,'Unknown') AS industry,
       COUNT(DISTINCT id) AS deals,
       ROUND(SUM(amount)/1e6,1) AS pipeline_m
FROM deals WHERE status = 'open'
GROUP BY industry ORDER BY pipeline_m DESC

=================================================================
"""

_QUERY_TOOL = {
    "name": "query_clickhouse",
    "description": (
        "Execute a SELECT query against the ClickHouse deals table. "
        "Use for any question about deals, pipeline value, win/loss rates, regions, "
        "industries, owners, stages, or any data metric. "
        "Always follow SELECT-only rule and LIMIT row queries to 100. "
        "Relay DATABASE CONNECTION FAILED or ERROR: messages directly to the user."
    ),
    "input_schema": {
        "type": "object",
        "properties": {
            "sql": {
                "type": "string",
                "description": "A valid SELECT or WITH query against the deals table.",
            }
        },
        "required": ["sql"],
    },
}


# =============================================================================
# Pydantic models
# =============================================================================

class ChatMessage(BaseModel):
    role: Literal["user", "assistant"]
    content: str

class ChatRequest(BaseModel):
    message: str
    history: List[ChatMessage] = []

class ExportRequest(BaseModel):
    format: Literal["pdf", "pptx"]
    conversation: List[ChatMessage] = []
    title: str = "Deals Intelligence Report"


# =============================================================================
# Claude tool loop
# =============================================================================

def _extract_text(content_blocks) -> str:
    return "\n".join(
        b.text for b in content_blocks if hasattr(b, "text") and b.text
    ).strip()


def _call_claude(messages: list, max_tokens: int = 2048) -> str:
    """
    Run Claude with the query_clickhouse tool.
    Executes up to 5 tool-use rounds, then returns the final text reply.
    """
    response = _ai_client.messages.create(
        model=_CLAUDE_MODEL,
        system=_SYSTEM_PROMPT,
        messages=messages,
        tools=[_QUERY_TOOL],
        temperature=0,
        max_tokens=max_tokens,
    )

    for round_num in range(5):
        if response.stop_reason != "tool_use":
            break

        tool_block = next((b for b in response.content if b.type == "tool_use"), None)
        if not tool_block:
            break

        sql = tool_block.input.get("sql", "")
        print(f"  🔄 Tool round {round_num + 1} | SQL: {sql[:120]}...")

        query_result = run_clickhouse_query(sql)

        is_error = any(query_result.startswith(p) for p in [
            "DATABASE CONNECTION FAILED",
            "ERROR:",
            "DATABASE ERROR:",
        ])

        messages = messages + [
            {"role": "assistant", "content": response.content},
            {
                "role": "user",
                "content": [
                    {
                        "type": "tool_result",
                        "tool_use_id": tool_block.id,
                        "content": query_result,
                        "is_error": is_error,
                    }
                ],
            },
        ]

        response = _ai_client.messages.create(
            model=_CLAUDE_MODEL,
            system=_SYSTEM_PROMPT,
            messages=messages,
            tools=[_QUERY_TOOL],
            temperature=0,
            max_tokens=max_tokens,
        )

    reply = _extract_text(response.content)

    if not reply:
        reply = (
            "⚠️ No response was generated.\n\n"
            "This usually means the ClickHouse connection failed before Claude could answer. "
            "Open **/debug/db** to diagnose the connection, then verify "
            "`CLICKHOUSE_API_URL` and `CLICKHOUSE_API_TOKEN` in `.env` and restart."
        )

    return reply


# =============================================================================
# Routes
# =============================================================================

@app.get("/", response_class=HTMLResponse)
def root():
    with open("chat.html", "r") as f:
        return HTMLResponse(content=f.read())


@app.get("/debug/db")
def debug_db():
    """
    Connectivity diagnostic. Open in browser: http://localhost:8000/debug/db
    Tests every relevant endpoint on your ClickHouse HTTP proxy.
    """
    base_url = _base_url()
    token    = _token()

    config = {
        "CLICKHOUSE_API_URL":   base_url or "❌ NOT SET",
        "CLICKHOUSE_API_TOKEN": f"✅ set ({len(token)} chars)" if token else "❌ NOT SET",
    }

    if not base_url or not token:
        return {
            "status": "MISCONFIGURED",
            "config": config,
            "fix": (
                "Add these to your .env file and restart:\n"
                "  CLICKHOUSE_API_URL=https://clickhouse-api-j55l.onrender.com\n"
                "  CLICKHOUSE_API_TOKEN=your_token_here"
            ),
        }

    results = {}

    # Test 1: health check (unauthenticated GET /)
    try:
        r = httpx.get(base_url, timeout=10)
        results["GET /"] = {"status": r.status_code, "body": r.text[:200]}
    except Exception as e:
        results["GET /"] = {"error": str(e)}

    # Test 2: list databases (authenticated)
    try:
        r = httpx.get(f"{base_url}/databases", headers=_auth_headers(), timeout=10)
        results["GET /databases"] = {"status": r.status_code, "body": r.text[:300]}
    except Exception as e:
        results["GET /databases"] = {"error": str(e)}

    # Test 3: actual query — SELECT 1
    ping_result = run_clickhouse_query("SELECT 1 AS ping")
    query_ok = not any(ping_result.startswith(p) for p in [
        "DATABASE CONNECTION FAILED", "ERROR:", "DATABASE ERROR:"
    ])
    results["POST /query (SELECT 1)"] = {"result": ping_result, "ok": query_ok}

    return {
        "status":  "OK" if query_ok else "FAILED",
        "config":  config,
        "tests":   results,
        "recommendation": (
            "✅ Database connected. Chat is ready."
            if query_ok else
            "❌ Query test failed. Check the 'POST /query' result above for the exact error.\n"
            "Common causes:\n"
            "  • Wrong token → 401 Unauthorized\n"
            "  • Wrong URL   → ConnectError or 404\n"
            "  • Proxy down  → ConnectError or 503"
        ),
    }


@app.post("/chat")
def chat(payload: ChatRequest):
    messages = [{"role": m.role, "content": m.content} for m in payload.history]
    messages.append({"role": "user", "content": payload.message})

    print(f"💬 [chat] {payload.message[:100]}")
    try:
        reply = _call_claude(messages)
    except Exception as exc:
        traceback.print_exc()
        raise HTTPException(status_code=502, detail=f"Claude error: {exc}")

    print(f"✅ [chat] {len(reply)} chars")
    return {"reply": reply}


@app.post("/export/pdf")
def export_pdf(payload: ExportRequest):
    conv_text = "\n\n".join(
        f"{'User' if m.role == 'user' else 'Assistant'}: {m.content}"
        for m in payload.conversation
    )
    prompt = (
        f"Based on this conversation, write a structured deals intelligence report "
        f"titled '{payload.title}'. Use these exact ## section headers:\n"
        "## Executive Summary\n## Pipeline Overview\n## Key Metrics\n"
        "## Regional Breakdown\n## Win / Loss Analysis\n## Recommendations\n\n"
        "Query the database for any missing numbers. Be data-driven.\n\n"
        f"CONVERSATION:\n{conv_text}"
    )
    try:
        report_text = _call_claude([{"role": "user", "content": prompt}], max_tokens=3000)
    except Exception as exc:
        raise HTTPException(status_code=502, detail=str(exc))

    pdf_bytes = _build_pdf(payload.title, report_text)
    filename  = re.sub(r"[^\w\-]", "_", payload.title) + ".pdf"
    return StreamingResponse(
        io.BytesIO(pdf_bytes), media_type="application/pdf",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


@app.post("/export/pptx")
def export_pptx(payload: ExportRequest):
    conv_text = "\n\n".join(
        f"{'User' if m.role == 'user' else 'Assistant'}: {m.content}"
        for m in payload.conversation
    )
    prompt = (
        f"Based on this conversation, write slide content for a presentation titled '{payload.title}'.\n"
        "Output each slide as:\nSLIDE: <Title>\nBULLETS:\n- bullet 1\n- bullet 2\n\n"
        "Include: Title/Overview, Pipeline Health, Key Metrics, Regional Breakdown, "
        "Win/Loss Analysis, Recommendations.\n"
        "Query the database for any missing data. Keep bullets crisp.\n\n"
        f"CONVERSATION:\n{conv_text}"
    )
    try:
        slide_text = _call_claude([{"role": "user", "content": prompt}], max_tokens=3000)
    except Exception as exc:
        raise HTTPException(status_code=502, detail=str(exc))

    pptx_bytes = _build_pptx(payload.title, slide_text)
    filename   = re.sub(r"[^\w\-]", "_", payload.title) + ".pptx"
    return StreamingResponse(
        io.BytesIO(pptx_bytes),
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


# =============================================================================
# PDF Builder
# =============================================================================

_C_NAVY  = colors.HexColor("#0D1B3E")
_C_BLUE  = colors.HexColor("#1565C0")
_C_WHITE = colors.white
_C_BG    = colors.HexColor("#F7F9FC")
_C_TXT   = colors.HexColor("#1E293B")
_C_DIM   = colors.HexColor("#94A3B8")

_SECTION_COLORS = {
    "executive": colors.HexColor("#0D1B3E"),
    "pipeline":  colors.HexColor("#1565C0"),
    "metric":    colors.HexColor("#004D40"),
    "regional":  colors.HexColor("#BF360C"),
    "win":       colors.HexColor("#B71C1C"),
    "loss":      colors.HexColor("#B71C1C"),
    "recommend": colors.HexColor("#1B5E20"),
}

PW, PH = A4
_ML = _MR = 0.6 * inch
_MT = 0.45 * inch
_MB = 0.40 * inch
_HDR_H = 44
_FTR_H = 20
_CW = PW - _ML - _MR


def _strip_md(t: str) -> str:
    t = re.sub(r'\*{1,3}(.*?)\*{1,3}', r'\1', t)
    t = re.sub(r'#{1,6}\s*', '', t)
    t = re.sub(r'^[\-\*•]\s*', '', t, flags=re.M)
    t = re.sub(r'`(.*?)`', r'\1', t)
    return t.strip()


def _pdf_styles():
    return {
        "Cover_Title": ParagraphStyle("Cover_Title", fontSize=26, leading=32,
                                       textColor=_C_WHITE, fontName="Helvetica-Bold", spaceAfter=8),
        "Cover_Sub":   ParagraphStyle("Cover_Sub",   fontSize=13, leading=18,
                                       textColor=colors.HexColor("#B0BEC5"), fontName="Helvetica"),
        "Section_H":   ParagraphStyle("Section_H",   fontSize=11, leading=15,
                                       textColor=_C_WHITE, fontName="Helvetica-Bold"),
        "Body":        ParagraphStyle("Body",   fontSize=9,  leading=14,
                                       textColor=_C_TXT, fontName="Helvetica", spaceAfter=4),
        "Bullet":      ParagraphStyle("Bullet", fontSize=9,  leading=14,
                                       textColor=_C_TXT, fontName="Helvetica",
                                       leftIndent=12, firstLineIndent=-8, spaceAfter=3),
        "H2":          ParagraphStyle("H2",  fontSize=11, leading=15, textColor=_C_NAVY,
                                       fontName="Helvetica-Bold", spaceBefore=10, spaceAfter=4),
        "H3":          ParagraphStyle("H3",  fontSize=9,  leading=13, textColor=_C_BLUE,
                                       fontName="Helvetica-Bold", spaceBefore=6, spaceAfter=2),
    }


def _parse_sections(text: str):
    parts = re.split(r'^##\s+', text, flags=re.MULTILINE)
    return [
        (lines[0].strip(), lines[1].strip() if len(lines) > 1 else "")
        for part in parts if part.strip()
        for lines in [part.strip().split("\n", 1)]
    ]


def _build_pdf(title: str, report_text: str) -> bytes:
    buf     = io.BytesIO()
    styles  = _pdf_styles()
    sections = _parse_sections(report_text)

    def _on_page(canvas, doc):
        canvas.saveState()
        canvas.setFillColor(_C_NAVY)
        canvas.rect(0, PH - _HDR_H - _MT, PW, _HDR_H + _MT, fill=1, stroke=0)
        canvas.setFillColor(_C_WHITE)
        canvas.setFont("Helvetica-Bold", 10)
        canvas.drawString(_ML, PH - _MT - 28, title)
        canvas.setFillColor(_C_BG)
        canvas.rect(0, 0, PW, _FTR_H + _MB, fill=1, stroke=0)
        canvas.setFillColor(_C_DIM)
        canvas.setFont("Helvetica", 7)
        footer = f"DIUD Report  |  AI-Generated  |  CONFIDENTIAL  |  {date.today().strftime('%B %Y')}"
        canvas.drawCentredString(PW / 2, _MB + 5, footer)
        canvas.drawRightString(PW - _MR, _MB + 5, f"Page {canvas.getPageNumber()}")
        canvas.restoreState()

    frame    = Frame(_ML, _MB + _FTR_H, _CW, PH - _HDR_H - _MT - _MB - _FTR_H, id="main")
    template = PageTemplate(id="main", frames=[frame], onPage=_on_page)
    doc = BaseDocTemplate(buf, pagesize=A4, leftMargin=_ML, rightMargin=_MR,
                          topMargin=_MT + _HDR_H, bottomMargin=_MB + _FTR_H)
    doc.addPageTemplates([template])

    story = [
        Spacer(1, 1.0 * inch),
        Paragraph(title, styles["Cover_Title"]),
        Paragraph(f"Generated {date.today().strftime('%B %d, %Y')}", styles["Cover_Sub"]),
        PageBreak(),
    ]

    for sec_title, sec_body in sections:
        color_key = next((k for k in _SECTION_COLORS if k in sec_title.lower()), None)
        bar_color = _SECTION_COLORS.get(color_key, _C_BLUE)

        story.append(Table(
            [[Paragraph(sec_title.upper(), styles["Section_H"])]],
            colWidths=[_CW],
            style=TableStyle([
                ("BACKGROUND",    (0, 0), (-1, -1), bar_color),
                ("TOPPADDING",    (0, 0), (-1, -1), 8),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
                ("LEFTPADDING",   (0, 0), (-1, -1), 10),
            ])
        ))
        story.append(Spacer(1, 6))

        for line in sec_body.split("\n"):
            line = line.strip()
            if not line:
                story.append(Spacer(1, 3))
            elif line.startswith("### "):
                story.append(Paragraph(_strip_md(line), styles["H3"]))
            elif line.startswith("## "):
                story.append(Paragraph(_strip_md(line), styles["H2"]))
            elif line.startswith(("- ", "* ", "• ")):
                story.append(Paragraph("• " + _strip_md(line[2:]), styles["Bullet"]))
            else:
                story.append(Paragraph(_strip_md(line), styles["Body"]))

        story.extend([Spacer(1, 12), PageBreak()])

    doc.build(story)
    return buf.getvalue()


# =============================================================================
# PPTX Builder
# =============================================================================

_C_NAVY_P  = RGBColor(0x0D, 0x1B, 0x3E)
_C_DNAV_P  = RGBColor(0x0A, 0x11, 0x28)
_C_BLUE_P  = RGBColor(0x1E, 0x88, 0xE5)
_C_WHITE_P = RGBColor(0xFF, 0xFF, 0xFF)
_C_LTBG_P  = RGBColor(0xF5, 0xF7, 0xFA)
_C_TXT_P   = RGBColor(0x1A, 0x1A, 0x2E)
_C_DIM_P   = RGBColor(0x88, 0x99, 0xAA)

_SLIDE_ACCENT = {
    "overview":  RGBColor(0x1E, 0x88, 0xE5),
    "pipeline":  RGBColor(0x00, 0x89, 0x7B),
    "metric":    RGBColor(0x2E, 0x7D, 0x32),
    "regional":  RGBColor(0xBF, 0x36, 0x0C),
    "win":       RGBColor(0x2E, 0x7D, 0x32),
    "loss":      RGBColor(0xC6, 0x28, 0x28),
    "recommend": RGBColor(0x1B, 0x5E, 0x20),
}

SLIDE_W = Inches(13.33)
SLIDE_H = Inches(7.5)


def _pptx_bg(slide, color):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color


def _pptx_rect(slide, l, t, w, h, color):
    shp = slide.shapes.add_shape(1, Inches(l), Inches(t), Inches(w), Inches(h))
    shp.fill.solid()
    shp.fill.fore_color.rgb = color
    shp.line.fill.background()
    return shp


def _pptx_txt(slide, text, l, t, w, h, bold=False, size=18, color=None, align=PP_ALIGN.LEFT):
    txb = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    txb.word_wrap = True
    tf = txb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size  = Pt(size)
    run.font.bold  = bold
    run.font.color.rgb = color or _C_TXT_P
    return txb


def _parse_slides(text: str):
    slides, cur_title, cur_bullets = [], None, []
    for line in text.split("\n"):
        line = line.rstrip()
        if line.startswith("SLIDE:"):
            if cur_title is not None:
                slides.append((cur_title, cur_bullets))
            cur_title, cur_bullets = line[6:].strip(), []
        elif line.startswith("- ") and cur_title:
            cur_bullets.append(line[2:].strip())
    if cur_title is not None:
        slides.append((cur_title, cur_bullets))
    return slides


def _build_pptx(title: str, slide_text: str) -> bytes:
    slides_data = _parse_slides(slide_text) or [(title, [slide_text[:400]])]

    prs = Presentation()
    prs.slide_width  = SLIDE_W
    prs.slide_height = SLIDE_H

    footer_text = f"DIUD  |  AI-Generated  |  CONFIDENTIAL  |  {date.today().strftime('%B %Y')}"
    blank = prs.slide_layouts[6]

    def _footer(slide):
        _pptx_rect(slide, 0, 7.1, 13.33, 0.4, _C_DNAV_P)
        _pptx_txt(slide, footer_text, 0.3, 7.12, 12, 0.35,
                  size=7, color=_C_DIM_P, align=PP_ALIGN.CENTER)

    def _accent(t_lower):
        for k, c in _SLIDE_ACCENT.items():
            if k in t_lower:
                return c
        return _C_BLUE_P

    # Cover slide
    cover = prs.slides.add_slide(blank)
    _pptx_bg(cover, _C_NAVY_P)
    _pptx_rect(cover, 0, 3.2, 13.33, 0.06, _C_BLUE_P)
    _pptx_txt(cover, title, 0.8, 1.6, 11.5, 1.4, bold=True, size=34, color=_C_WHITE_P)
    _pptx_txt(cover, "Deals Intelligence Report", 0.8, 3.0, 8, 0.6,
              size=15, color=RGBColor(0xB0, 0xBE, 0xC5))
    _pptx_txt(cover, f"Generated: {date.today().strftime('%B %d, %Y')}", 0.8, 3.6, 6, 0.45,
              size=12, color=RGBColor(0x78, 0x90, 0x9C))

    # Content slides
    for i, (s_title, bullets) in enumerate(slides_data):
        slide  = prs.slides.add_slide(blank)
        accent = _accent(s_title.lower())
        _pptx_bg(slide, _C_LTBG_P)
        _pptx_rect(slide, 0, 0, 13.33, 0.9, accent)
        _pptx_txt(slide, s_title.upper(), 0.35, 0.1, 12.5, 0.7,
                  bold=True, size=18, color=_C_WHITE_P)
        _pptx_txt(slide, str(i + 1), 12.5, 0.12, 0.6, 0.6,
                  size=11, color=_C_WHITE_P, align=PP_ALIGN.RIGHT)
        _pptx_rect(slide, 0.3, 1.0, 12.73, 5.9, _C_WHITE_P)

        if bullets:
            txb = slide.shapes.add_textbox(Inches(0.5), Inches(1.1), Inches(12.3), Inches(5.6))
            txb.word_wrap = True
            tf = txb.text_frame
            tf.word_wrap = True
            for j, bullet in enumerate(bullets[:12]):
                p = tf.add_paragraph() if j > 0 else tf.paragraphs[0]
                p.space_before = Pt(4)
                dot = p.add_run()
                dot.text = "●  "
                dot.font.size = Pt(8)
                dot.font.color.rgb = accent
                run = p.add_run()
                run.text = bullet
                run.font.size = Pt(12)
                run.font.color.rgb = _C_TXT_P
        else:
            _pptx_txt(slide, "No data available.", 0.5, 1.2, 12, 0.5,
                      size=11, color=_C_DIM_P)

        _footer(slide)

    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()
