import io
import json
import os
import re
import traceback
from datetime import date
from typing import List, Literal, Optional

import httpx
import anthropic
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, HTMLResponse, StreamingResponse
from pydantic import BaseModel
from dotenv import load_dotenv

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import (
    BaseDocTemplate, Frame, PageTemplate,
    PageBreak, Paragraph, Spacer, Table, TableStyle,
)

# =============================================================================
# Load ENV
# =============================================================================

load_dotenv()

# =============================================================================
# FastAPI App
# =============================================================================

app = FastAPI(title="DIUD", description="Decision Intelligence Using Data", version="3.0.0")

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

_ai_client    = anthropic.Anthropic(api_key=os.getenv("ANTHROPIC_API_KEY"))
_CLAUDE_MODEL = "claude-sonnet-4-5"

# =============================================================================
# ClickHouse HTTP proxy — base helpers
# =============================================================================

def _base_url() -> str:
    return (os.getenv("CLICKHOUSE_API_URL") or "").rstrip("/")

def _token() -> str:
    return os.getenv("CLICKHOUSE_API_TOKEN") or ""

def _auth_headers() -> dict:
    return {
        "Authorization": f"Bearer {_token()}",
        "Content-Type":  "application/json",
    }

FORBIDDEN_KEYWORDS = ["INSERT", "UPDATE", "DELETE", "DROP", "ALTER", "TRUNCATE", "CREATE"]

# =============================================================================
# Schema discovery — called once at startup, result stored in _LIVE_SCHEMA
# =============================================================================

_LIVE_SCHEMA: dict = {}
_SCHEMA_BLOCK: str = "Schema not yet loaded."


def _proxy_get(path: str) -> dict | list | None:
    """GET {base_url}{path} with auth. Returns parsed JSON or None on error."""
    base_url = _base_url()
    token    = _token()
    if not base_url or not token:
        return None
    try:
        r = httpx.get(f"{base_url}{path}", headers=_auth_headers(), timeout=20)
        if r.status_code == 200:
            return r.json()
        print(f"   ⚠️  GET {path} → HTTP {r.status_code}: {r.text[:200]}")
        return None
    except Exception as e:
        print(f"   ⚠️  GET {path} → {e}")
        return None


def discover_schema() -> str:
    global _LIVE_SCHEMA, _SCHEMA_BLOCK

    print("🔎 Discovering schema from ClickHouse proxy…")

    databases_raw = _proxy_get("/databases")
    if not databases_raw:
        msg = "⚠️  Could not fetch databases — check CLICKHOUSE_API_URL / CLICKHOUSE_API_TOKEN."
        print(msg)
        _SCHEMA_BLOCK = msg
        return msg

    if isinstance(databases_raw, list) and databases_raw:
        if isinstance(databases_raw[0], str):
            databases = databases_raw
        elif isinstance(databases_raw[0], dict):
            databases = [
                d.get("name") or d.get("database") or list(d.values())[0]
                for d in databases_raw
            ]
        else:
            databases = [str(d) for d in databases_raw]
    elif isinstance(databases_raw, dict):
        databases = (
            databases_raw.get("data") or
            databases_raw.get("databases") or
            list(databases_raw.values())[0]
            if databases_raw else []
        )
    else:
        databases = []

    SKIP_DBS = {"system", "information_schema", "INFORMATION_SCHEMA"}
    databases = [d for d in databases if d not in SKIP_DBS]
    print(f"   Databases found: {databases}")

    schema_lines = []
    schema_dict  = {}

    for db in databases:
        tables_raw = _proxy_get(f"/tables/{db}")
        if not tables_raw:
            continue

        if isinstance(tables_raw, list) and tables_raw:
            if isinstance(tables_raw[0], str):
                tables = tables_raw
            elif isinstance(tables_raw[0], dict):
                tables = [
                    t.get("name") or t.get("table") or list(t.values())[0]
                    for t in tables_raw
                ]
            else:
                tables = [str(t) for t in tables_raw]
        elif isinstance(tables_raw, dict):
            tables = (
                tables_raw.get("data") or
                tables_raw.get("tables") or
                list(tables_raw.values())[0]
                if tables_raw else []
            )
        else:
            tables = []

        print(f"   {db}: tables = {tables}")

        for tbl in tables:
            schema_raw = _proxy_get(f"/schema/{db}/{tbl}")
            if not schema_raw:
                schema_lines.append(f"\nTABLE: {db}.{tbl}\n  (schema unavailable)")
                continue

            if isinstance(schema_raw, list):
                cols = schema_raw
            elif isinstance(schema_raw, dict):
                cols = (
                    schema_raw.get("columns") or
                    schema_raw.get("data") or
                    schema_raw.get("schema") or
                    [schema_raw]
                )
            else:
                cols = []

            schema_dict[f"{db}.{tbl}"] = cols

            col_lines = []
            for col in cols:
                if isinstance(col, dict):
                    col_name = (
                        col.get("name") or col.get("column_name") or
                        col.get("Field") or list(col.keys())[0]
                    )
                    col_type = (
                        col.get("type") or col.get("data_type") or
                        col.get("Type") or ""
                    )
                    col_comment = col.get("comment") or col.get("Comment") or ""
                    comment_str = f"  — {col_comment}" if col_comment else ""
                    col_lines.append(f"  {col_name:<35} {col_type}{comment_str}")
                else:
                    col_lines.append(f"  {col}")

            schema_lines.append(f"\nTABLE: {db}.{tbl}")
            schema_lines.extend(col_lines)

    _LIVE_SCHEMA  = schema_dict
    _SCHEMA_BLOCK = "\n".join(schema_lines) if schema_lines else "No tables found."

    print(f"✅ Schema loaded: {list(schema_dict.keys())}")
    return _SCHEMA_BLOCK


# =============================================================================
# Build system prompt dynamically with live schema
# =============================================================================

def _build_system_prompt() -> str:
    table_refs = list(_LIVE_SCHEMA.keys())
    primary_table = table_refs[0] if table_refs else "your_database.deals"

    compact_lines = []
    for table_key, cols in _LIVE_SCHEMA.items():
        col_parts = []
        for col in cols:
            if isinstance(col, dict):
                name = (
                    col.get("name") or col.get("column_name") or
                    col.get("Field") or list(col.keys())[0]
                )
                typ = (
                    col.get("type") or col.get("data_type") or
                    col.get("Type") or ""
                )
                col_parts.append(f"{name}:{typ}")
            else:
                col_parts.append(str(col))
        compact_lines.append(f"{table_key}({', '.join(col_parts)})")

    schema = "\n".join(compact_lines) or "Schema not yet loaded — try restarting the server."

    if len(schema) > 20000:
        schema = schema[:20000] + "\n[schema truncated — too many tables/columns]"

    return f"""
You are DIUD (Decision Intelligence Using Data) — a conversational data assistant.
You have LIVE access to a ClickHouse database via the query_clickhouse tool.

=================================================================
GREETING RULE — HIGHEST PRIORITY
=================================================================
If the user's message is ONLY a greeting (e.g. "hi", "hey", "hello",
"hey there", "good morning", or any short salutation with no question),
respond with EXACTLY this and nothing else:

"Hey, I'm DIUD, your data intelligence agent to help you analyse
the live ClickHouse or Web data. How may I help you?"

Do NOT add bullet points, capability lists, examples, or any other
content. This rule overrides all other response formatting rules.
 
=================================================================
CLICKHOUSE DIRECT ACCESS
=================================================================
 
You have a tool called query_clickhouse.
Use it for any question about pipeline deals, AEs, regions, industries,
stages, win/loss, competitors, conversions, or any metric not already
in the conversation context.
 
If the tool returns a DATABASE CONNECTION FAILED message, relay it to
the user clearly and suggest they check /debug/clickhouse.

=================================================================
DUPLICATE RECORD EXCLUSION — ALWAYS APPLY
=================================================================
ALL queries must exclude duplicate records. Apply these rules
on EVERY table used:

1. hs_analytics tables (deals, owners, contacts, companies)
   ALWAYS use the FINAL keyword — it dedduplicates ReplacingMergeTree:
     FROM hs_analytics.deals FINAL
     FROM hs_analytics.owners FINAL
     FROM hs_analytics.contacts FINAL
     FROM hs_analytics.companies FINAL

2. Aggregations — always use countDistinct(), never count():
     countDistinct(deal_id)    — unique deals
     countDistinct(contact_id) — unique contacts/MQLs
     countDistinct(owner_id)   — unique owners
   NEVER use count(*) or count(deal_id) for business metrics.

3. Association / helper tables (no FINAL needed, but deduplicate
   the join result):
     LEFT JOIN (
       SELECT DISTINCT contact_id, deal_id
       FROM kore_ai_hubspot.gs_DealContactAssociation
     ) z ON c.contact_id = z.contact_id
   This prevents one contact mapping to the same deal multiple
   times and inflating MQL conversion counts.

4. Targets table — always GROUP BY + SUM, never raw select:
     SELECT region, mql_source, SUM(mql_target) AS mql_target
     FROM kore_ai_hubspot.gs_marketing_targets
     GROUP BY region, mql_source
   Multiple rows per combination exist by design — raw counts
   will always be wrong.

CHECKLIST before every query:
  ☐ FINAL on every hs_analytics table
  ☐ countDistinct() for every unique-count metric
  ☐ DISTINCT inside gs_DealContactAssociation subquery
  ☐ SUM + GROUP BY on gs_marketing_targets
=================================================================
 
=================================================================
TABLES
=================================================================
 
── TABLE 1: hs_analytics.deals ─────────────────────────────────
Primary table. One row per deal.
ALWAYS use FINAL keyword: FROM hs_analytics.deals FINAL
 
KEY COLUMNS:
  deal_id                    STRING  — unique deal identifier
  deal_name                  STRING  — name of the deal
  deal_owner                 STRING  — owner ID (join to hs_analytics.owners on o.id)
  deal_stage                 STRING  — current stage (see STAGE LIST below)
  deal_type                  STRING  — deal type (NULL = 'Not Assigned')
  pipeline                   STRING  — always filter: pipeline = 'default'
  amount                     FLOAT   — deal value in USD
  region                     STRING  — raw values (see REGION MAP below)
  deal_source_rollup         STRING  — raw source (see SOURCE MAP below)
  20_snapshot_deal_source_rollup  STRING — source at time of 20% qualification
  ai_for_x                   STRING  — AI use case category
  kore_primary_industry      STRING  — raw industry (see INDUSTRY MAP below)
  account_priority_level     STRING  — 'P1','P2'...'P10' (raw, not grouped)
  hubspot_team               STRING  — team ID (join to kore_ai_hubspot.gs_Teams)
 
  -- DATE COLUMNS (stored as strings; always cast to DATE before comparison)
  create_date                STRING  — deal creation date
  close_date                 STRING  — expected/actual close date
  became_5_deal_date         STRING  — entered 5% IQM Held
  became_10_deal_date        STRING  — entered 10% Discovery
  became_20_deal_date        STRING  — entered 20% Solution
  became_30_deal_date        STRING  — entered 30% Proof
  became_40_deal_date        STRING  — entered 40% Proposal
  became_60_deal_date        STRING  — entered 60% Price Negotiation
  became_75_deal_date        STRING  — entered 75% Contract Review
  last_contacted             STRING  — last contact date
 
── TABLE 2: hs_analytics.owners ─────────────────────────────────
ALWAYS use FINAL: FROM hs_analytics.owners FINAL
 
  id           STRING  — owner ID (join key to deals.deal_owner)
  firstName    STRING  — first name
  lastName     STRING  — last name
  email        STRING  — owner email
 
── TABLE 3: hs_analytics.companies ──────────────────────────────
ALWAYS use FINAL: FROM hs_analytics.companies FINAL
 
  company_id   STRING  — unique company ID
  name         STRING  — company name
  domain       STRING  — website domain
  industry     STRING  — company industry
  country      STRING  — company country
  city         STRING  — company city
 
── TABLE 4: hs_analytics.contacts ───────────────────────────────
ALWAYS use FINAL: FROM hs_analytics.contacts FINAL
 
  contact_id        STRING   — unique contact identifier
  email             STRING   — contact email
  first_name        STRING   — first name
  last_name         STRING   — last name
  company_name      STRING   — associated company name
  company_priority  STRING   — 'P1'–'P10'
  region            STRING   — raw region
  original_source   STRING   — raw HubSpot source
  lead_status       STRING   — exclude 'Bad Data' in most queries
  lifecycle_stage   STRING   — current lifecycle stage
  date_entered_marketing_qualified_lead_lifecycle_stage_pipeline
                    STRING   — date contact became MQL
 
── TABLE 5: kore_ai_hubspot.gs_DealContactAssociation ───────────
  contact_id  STRING  — links to hs_analytics.contacts.contact_id
  deal_id     STRING  — links to hs_analytics.deals.deal_id
 
── TABLE 6: kore_ai_hubspot.gs_marketing_targets ────────────────
  fy               STRING   — fiscal year
  quarter          STRING   — 'Q1','Q2','Q3','Q4'
  month            STRING   — abbreviated month
  region           STRING   — display region
  original_source  STRING   — display source
  mql_target       FLOAT32  — MQL count target
 
── TABLE 7: kore_ai_hubspot.gs_deal_ids_hs ──────────────────────
  deal_id_hs  STRING   — valid deal IDs whitelist

=================================================================
MANDATORY BASE FILTERS (apply to EVERY deals query)
=================================================================
Always include ALL three in every query on hs_analytics.deals:
 
  WHERE pipeline = 'default'
  AND CASE WHEN deal_type IS NULL THEN 'Not Assigned' ELSE deal_type END
      NOT IN ('Partner-Led SMB')
  AND toInt64(deal_id) IN (
      SELECT DISTINCT toInt64(deal_id_hs)
      FROM kore_ai_hubspot.gs_deal_ids_hs
  )

=================================================================
FISCAL YEAR CALCULATION
=================================================================
Kore.ai fiscal year runs April → March.
FY27 = Apr 2026 – Mar 2027   (result = 2027)
FY26 = Apr 2025 – Mar 2026   (result = 2026)

Active pipeline FY27: close_date >= '2026-04-01' AND close_date <= '2027-03-31'

=================================================================
ACTIVE PIPELINE DEFINITION (FY27)
=================================================================
  WHERE deal_stage IN ('20% - Solution','30% - Proof','40% - Proposal',
                       '60% - Price Negotiation','75% - Contract Review')
  AND toDate(LEFT(coalesce(close_date,'1900-01-01'),10)) >= '2026-04-01'
  AND toDate(LEFT(coalesce(close_date,'1900-01-01'),10)) <= '2027-03-31'

=================================================================
QUERY RULES
=================================================================
1. SELECT / WITH only — never INSERT, UPDATE, DELETE, DROP, ALTER, TRUNCATE
2. Always use FINAL on all hs_analytics tables
3. Always apply the 3 mandatory base filters on deals
4. For AGGREGATION/SUMMARY queries (counts, totals, breakdowns): no row limit needed.
5. For DEAL LIST queries (individual deal rows): do NOT add a LIMIT unless the user
   explicitly asks for "top N" or "first N". Return ALL matching rows so the user
   can export the complete dataset. The system will handle displaying them safely.
6. Use countDistinct(deal_id) for unique deal counts
7. Use round(sum(amount)/1e6, 1) for $M dollar amounts
8. Dates stored as strings — always cast: toDate(LEFT(coalesce(col,'1900-01-01'),10))
9. Default fiscal year context is FY27 unless user specifies otherwise

CORE RULES:
- NEVER say you lack database access. You always have it via the tool.
- NEVER fabricate numbers. Query the DB for every metric question.
- NEVER run destructive SQL.
- Answer in clean markdown: use tables for data, bold for key numbers.
- Be concise but complete.
- When generating export content, use ## section headers.
- When returning a deal list, always tell the user the TOTAL count found (e.g.
  "Found 256 deals matching your filters") even if the chat preview is condensed.

"""


# Initialise with a placeholder — will be replaced on startup
_SYSTEM_PROMPT = _build_system_prompt()


# =============================================================================
# ClickHouse query runner
# =============================================================================

def run_clickhouse_query(sql: str) -> str:
    base_url = _base_url()
    token    = _token()

    if not base_url:
        return (
            "DATABASE CONNECTION FAILED: CLICKHOUSE_API_URL is not set. "
            "Add it to .env and restart."
        )
    if not token:
        return (
            "DATABASE CONNECTION FAILED: CLICKHOUSE_API_TOKEN is not set. "
            "Add it to .env and restart."
        )

    stripped = sql.strip().upper()
    if not (stripped.startswith("SELECT") or stripped.startswith("WITH")):
        return "ERROR: Only SELECT/WITH queries are permitted."
    for kw in FORBIDDEN_KEYWORDS:
        if re.search(rf'\b{kw}\b', stripped):
            return f"ERROR: Forbidden keyword: {kw}"

    print(f"🔍 SQL → POST {base_url}/query\n   {sql[:300]}")

    try:
        resp = httpx.post(
            f"{base_url}/query",
            headers=_auth_headers(),
            json={"query": sql},
            timeout=30,
        )

        if resp.status_code == 401:
            return "DATABASE CONNECTION FAILED: 401 Unauthorized — check CLICKHOUSE_API_TOKEN."
        if resp.status_code == 403:
            return "DATABASE CONNECTION FAILED: 403 Forbidden."
        if resp.status_code == 422:
            detail = resp.text[:600]
            return f"ERROR: Proxy rejected the query (422). Detail: {detail}"
        if resp.status_code == 500:
            detail = resp.text[:600]
            return (
                f"DATABASE ERROR: HTTP 500 Internal Server Error from the proxy.\n"
                f"Detail: {detail}"
            )
        if resp.status_code != 200:
            return f"DATABASE ERROR: HTTP {resp.status_code} — {resp.text[:400]}"

        payload = resp.json()

        if isinstance(payload, list):
            rows = payload
        elif isinstance(payload, dict):
            rows = (
                payload.get("data") or payload.get("rows") or
                payload.get("result") or payload.get("results") or None
            )
            if rows is None:
                return json.dumps(payload, indent=2, default=str)[:3000]
        else:
            return f"Unexpected response type: {type(payload)}"

        if not rows:
            return "Query returned 0 rows."

        total_rows = len(rows)
        # Cap chat-display at 100 rows to keep the response readable,
        # but embed the FULL dataset as a compact JSON block so the
        # export pipeline can access every row.
        CHAT_DISPLAY_LIMIT = 100

        if isinstance(rows[0], dict):
            cols   = list(rows[0].keys())
            header = " | ".join(cols)
            display_rows = rows[:CHAT_DISPLAY_LIMIT]
            lines  = [header, "-" * min(len(header), 140)]
            for row in display_rows:
                lines.append(" | ".join(str(row.get(c, "NULL")) for c in cols))
        else:
            cols = None
            display_rows = rows[1:CHAT_DISPLAY_LIMIT + 1]
            lines = [" | ".join(str(v) for v in rows[0]), "-" * 80]
            for row in display_rows:
                lines.append(" | ".join(str(v) for v in row))

        if total_rows > CHAT_DISPLAY_LIMIT:
            lines.append(
                f"\n📊 **Showing {CHAT_DISPLAY_LIMIT} of {total_rows} rows** in this preview. "
                f"The full {total_rows} rows are available for export (PDF/CSV)."
            )

        # Embed the full dataset as a structured block for export use.
        # This block is parsed by the export pipeline; Claude ignores it.
        import json as _json
        full_data = {
            "total_rows": total_rows,
            "columns": cols if cols else [],
            "rows": rows  # all rows, not capped
        }
        full_block = (
            "\n\n<!-- FULL_DATASET_JSON\n"
            + _json.dumps(full_data, default=str)
            + "\nEND_FULL_DATASET_JSON -->"
        )

        result = "\n".join(lines) + full_block
        return result

    except httpx.ConnectError as e:
        return f"DATABASE CONNECTION FAILED: Could not reach {base_url}. Detail: {e}"
    except httpx.TimeoutException:
        return "DATABASE CONNECTION FAILED: Query timed out after 30 seconds."
    except Exception as exc:
        traceback.print_exc()
        return f"DATABASE CONNECTION FAILED: {type(exc).__name__}: {exc}"


# =============================================================================
# Startup event
# =============================================================================

@app.on_event("startup")
async def on_startup():
    global _SYSTEM_PROMPT
    discover_schema()
    _SYSTEM_PROMPT = _build_system_prompt()
    print("🚀 System prompt built with live schema.")


# =============================================================================
# Claude tool definition
# =============================================================================

_QUERY_TOOL = {
    "name": "query_clickhouse",
    "description": (
        "Execute a SELECT query against ClickHouse. "
        "Use for any question about deals, pipeline, win/loss, regions, owners, stages, metrics. "
        "ALWAYS use fully-qualified table names (database.table). "
        "If the result starts with DATABASE CONNECTION FAILED or ERROR: relay it to the user."
    ),
    "input_schema": {
        "type": "object",
        "properties": {
            "sql": {
                "type": "string",
                "description": (
                    "A valid ClickHouse SELECT or WITH query. "
                    "Use the exact database.table name from the schema. "
                    "For deal LIST queries, do NOT add a LIMIT clause unless the user asked for 'top N'. "
                    "For aggregation/summary queries (GROUP BY, counts, totals) no limit is needed. "
                    "Always return all rows matching the user's filters."
                ),
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

class ExportPreviewRequest(BaseModel):
    """Request to generate a preview of the export document from conversation."""
    conversation: List[ChatMessage] = []
    title: str = "Pipeline Intelligence Report"
    export_type: Literal["pdf", "pptx"] = "pdf"
    detail_level: Literal["summary", "detailed"] = "detailed"
    sections_to_include: Optional[List[str]] = None  # None = include all

class ExportDownloadRequest(BaseModel):
    """Request to download the final document."""
    format: Literal["pdf", "pptx"]
    content: str          # The pre-generated markdown content from preview
    title: str = "Pipeline Intelligence Report"

class ExportCSVRequest(BaseModel):
    """Request to export raw deal data as CSV, extracted from conversation history."""
    conversation: List[ChatMessage] = []
    title: str = "deals-export"


# =============================================================================
# Claude tool loop
# =============================================================================

def _extract_text(content_blocks) -> str:
    return "\n".join(
        b.text for b in content_blocks if hasattr(b, "text") and b.text
    ).strip()


def _call_claude(messages: list, max_tokens: int = 2048) -> str:
    """Run Claude with query_clickhouse tool. Up to 5 tool rounds."""

    safe_messages = []
    for m in messages:
        content = m.get("content", "")
        if isinstance(content, list):
            text_parts = []
            for b in content:
                if isinstance(b, dict) and b.get("type") == "text":
                    text_parts.append(b.get("text", ""))
                elif hasattr(b, "type") and b.type == "text" and hasattr(b, "text"):
                    text_parts.append(b.text)
            text = "\n".join(t for t in text_parts if t).strip()
            if text:
                safe_messages.append({"role": m["role"], "content": text})
        else:
            safe_messages.append({"role": m["role"], "content": content})

    response = _ai_client.messages.create(
        model=_CLAUDE_MODEL,
        system=_SYSTEM_PROMPT,
        messages=safe_messages,
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

        sql          = tool_block.input.get("sql", "")
        query_result = run_clickhouse_query(sql)
        is_error     = any(query_result.startswith(p) for p in [
            "DATABASE CONNECTION FAILED", "ERROR:", "DATABASE ERROR:"
        ])

        safe_messages = safe_messages + [
            {"role": "assistant", "content": response.content},
            {
                "role": "user",
                "content": [{
                    "type":        "tool_result",
                    "tool_use_id": tool_block.id,
                    "content":     query_result,
                    "is_error":    is_error,
                }],
            },
        ]

        response = _ai_client.messages.create(
            model=_CLAUDE_MODEL,
            system=_SYSTEM_PROMPT,
            messages=safe_messages,
            tools=[_QUERY_TOOL],
            temperature=0,
            max_tokens=max_tokens,
        )

    reply = _extract_text(response.content)
    if not reply:
        reply = (
            "⚠️ No response was generated.\n\n"
            "This usually means the ClickHouse query failed. "
            "Open **/debug/db** to diagnose."
        )
    return reply


# =============================================================================
# Export content generation — AI-driven document structuring
# =============================================================================

def _strip_dataset_block(text: str) -> str:
    """Remove the hidden FULL_DATASET_JSON comment block from a message string."""
    return re.sub(
        r'\n*<!--\s*FULL_DATASET_JSON.*?END_FULL_DATASET_JSON\s*-->',
        '',
        text,
        flags=re.DOTALL
    ).strip()


def _dataset_to_markdown_table(dataset: dict, max_rows: int = 5000) -> str:
    """
    Render the full dataset as a GitHub-flavoured markdown table.
    max_rows is a safety cap to avoid unbounded memory use in the PDF builder.
    """
    rows    = dataset.get("rows", [])
    columns = dataset.get("columns", [])
    total   = dataset.get("total_rows", len(rows))

    if not rows:
        return "_No data returned._"

    # Determine column names
    if not columns and rows and isinstance(rows[0], dict):
        columns = list(rows[0].keys())

    # Header
    header = "| " + " | ".join(str(c) for c in columns) + " |"
    sep    = "| " + " | ".join("---" for _ in columns) + " |"
    lines  = [header, sep]

    for row in rows[:max_rows]:
        if isinstance(row, dict):
            cells = [str(row.get(c, "")) for c in columns]
        else:
            cells = [str(v) for v in row]
        # Escape pipe chars inside cells
        cells = [c.replace("|", "\\|") for c in cells]
        lines.append("| " + " | ".join(cells) + " |")

    if total > max_rows:
        lines.append(f"\n_Showing {max_rows:,} of {total:,} rows._")
    else:
        lines.append(f"\n_Total: {total:,} rows._")

    return "\n".join(lines)


def _generate_export_content(
    conversation: List[ChatMessage],
    title: str,
    export_type: str,
    detail_level: str,
    sections_to_include: Optional[List[str]] = None,
) -> str:
    """
    Use Claude to intelligently structure the conversation into a
    well-formatted document. Returns markdown text.

    When the conversation contains a full deal-list dataset (embedded via
    FULL_DATASET_JSON), the complete rows are injected as a markdown table
    in place of the truncated chat preview, so every row appears in the export.
    """

    # Extract the full dataset once (used both to inject and to strip the block)
    dataset = _extract_full_dataset(conversation)

    # Build clean conversation text:
    # - Strip the hidden JSON block from every assistant message
    # - Replace the truncated preview notice with the full markdown table
    conv_parts = []
    for m in conversation:
        role    = 'USER' if m.role == 'user' else 'DIUD AGENT'
        content = _strip_dataset_block(m.content)

        # If this assistant message had the truncated preview notice,
        # append the full dataset table so Claude sees all rows
        if (
            m.role == 'assistant'
            and dataset
            and 'Showing' in m.content
            and 'rows' in m.content.lower()
        ):
            full_table = _dataset_to_markdown_table(dataset)
            # Remove the "Showing X of Y rows" notice and replace with full table
            content = re.sub(
                r'📊 \*\*Showing \d+ of \d+ rows\*\*.*',
                '',
                content,
                flags=re.DOTALL
            ).strip()
            content = content + "\n\n**Full Deal List:**\n\n" + full_table

        conv_parts.append(f"{role}: {content}")

    conv_text = "\n\n".join(conv_parts)

    sections_hint = ""
    if sections_to_include:
        sections_hint = f"\nOnly include these sections: {', '.join(sections_to_include)}"

    detail_hint = (
        "Include all data, tables, and detailed analysis from the conversation. "
        "The full deal list table above must appear verbatim in the document — do NOT summarise or truncate it."
        if detail_level == "detailed"
        else "Provide a high-level executive summary with key metrics and insights only."
    )

    format_hint = (
        "Format for a PowerPoint presentation: use SLIDE: <title> for each slide, followed by bullet points."
        if export_type == "pptx"
        else "Format as a professional PDF report with ## section headers, tables, and narrative prose."
    )

    prompt = f"""You are preparing a professional export document from a data intelligence conversation.

CONVERSATION:
{conv_text}

TASK:
Create a well-structured {export_type.upper()} document titled "{title}".

{format_hint}
{detail_hint}
{sections_hint}

REQUIREMENTS:
- Extract all key metrics, data tables, and insights from the conversation
- If a "Full Deal List" table is present in the conversation, include it completely in a dedicated ## Deal List section — preserve every row, do not truncate
- Organise logically with clear sections
- Preserve all numerical data accurately
- Add an executive summary at the start
- Include a "Key Recommendations" section at the end if insights warrant it
- Make it suitable for a business audience
- Today's date: {date.today().strftime('%B %d, %Y')}

Generate the document content now:"""

    messages = [{"role": "user", "content": prompt}]

    response = _ai_client.messages.create(
        model=_CLAUDE_MODEL,
        system="You are a professional business report writer. Generate clean, well-structured document content.",
        messages=messages,
        temperature=0,
        max_tokens=8192,  # increased to handle large deal tables
    )



    return _extract_text(response.content)


# =============================================================================
# Routes
# =============================================================================

@app.get("/", response_class=HTMLResponse)
def root():
    with open("chat.html", "r") as f:
        return HTMLResponse(content=f.read())


@app.get("/logo.png")
def serve_logo():
    return FileResponse("logo.png", media_type="image/png")


@app.get("/debug/db")
def debug_db():
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
            "fix": "Add CLICKHOUSE_API_URL and CLICKHOUSE_API_TOKEN to .env, then restart.",
        }

    tests = {}

    try:
        r = httpx.get(base_url, timeout=10)
        tests["GET /"] = {"status": r.status_code, "body": r.text[:200]}
    except Exception as e:
        tests["GET /"] = {"error": str(e)}

    try:
        r = httpx.get(f"{base_url}/databases", headers=_auth_headers(), timeout=10)
        tests["GET /databases"] = {"status": r.status_code, "body": r.text[:400]}
    except Exception as e:
        tests["GET /databases"] = {"error": str(e)}

    ping = run_clickhouse_query("SELECT 1 AS ping")
    query_ok = not any(ping.startswith(p) for p in [
        "DATABASE CONNECTION FAILED", "ERROR:", "DATABASE ERROR:"
    ])
    tests["POST /query (SELECT 1)"] = {"result": ping, "ok": query_ok}

    return {
        "status":            "OK" if query_ok else "FAILED",
        "config":            config,
        "discovered_tables": list(_LIVE_SCHEMA.keys()),
        "tests":             tests,
        "recommendation": (
            "✅ DB connected. Schema loaded. Chat is ready."
            if query_ok else
            "❌ Query test failed. See 'POST /query' result for the exact error."
        ),
    }


@app.post("/refresh-schema")
def refresh_schema():
    global _SYSTEM_PROMPT
    schema = discover_schema()
    _SYSTEM_PROMPT = _build_system_prompt()
    return {
        "status":  "refreshed",
        "tables":  list(_LIVE_SCHEMA.keys()),
        "schema":  schema[:2000],
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


# =============================================================================
# NEW: Export Preview — generates document content for side panel
# =============================================================================

@app.post("/export/preview")
async def export_preview(req: ExportPreviewRequest):
    """
    Generate a preview of the document from conversation context.

    When the conversation contains a full deal-list dataset, the complete
    rows are rendered directly into the markdown content (no row cap).
    Claude writes the narrative/summary; the table is injected verbatim.
    """
    if not req.conversation:
        raise HTTPException(status_code=400, detail="No conversation to export.")

    print(f"📄 [export/preview] type={req.export_type} detail={req.detail_level} msgs={len(req.conversation)}")

    dataset       = _extract_full_dataset(req.conversation)
    full_row_count = dataset.get("total_rows", 0) if dataset else 0

    try:
        # Step 1 — Claude generates the narrative document (summary, insights, etc.)
        # The full table is intentionally kept out of the Claude prompt to avoid
        # token waste; we inject it ourselves in Step 2.
        ai_content = _generate_export_content(
            conversation=req.conversation,
            title=req.title,
            export_type=req.export_type,
            detail_level=req.detail_level,
            sections_to_include=req.sections_to_include,
        )

        # Step 2 — If a full deal dataset exists, inject the complete table.
        # Replace any placeholder/truncated table Claude may have written
        # with the authoritative full-row version built from the raw data.
        if dataset and full_row_count > 0:
            full_table_md = _dataset_to_markdown_table(dataset)
            total         = dataset.get("total_rows", 0)
            section_header = f"\n\n## Deal List ({total:,} deals)\n\n"

            # If Claude already wrote a Deal List section, replace its table body
            if "## Deal List" in ai_content or "## deal list" in ai_content.lower():
                # Remove everything from the Deal List header to the next ## or end
                ai_content = re.sub(
                    r'##\s*Deal List.*?(?=\n##|\Z)',
                    section_header + full_table_md + "\n\n",
                    ai_content,
                    flags=re.DOTALL | re.IGNORECASE
                )
            else:
                # Append the full table as a new section at the end
                ai_content = ai_content.rstrip() + section_header + full_table_md

        content = ai_content

    except Exception as exc:
        traceback.print_exc()
        raise HTTPException(status_code=502, detail=f"Content generation error: {exc}")

    return {
        "content": content,
        "title": req.title,
        "export_type": req.export_type,
        "word_count": len(content.split()),
        "generated_at": date.today().isoformat(),
        "full_row_count": full_row_count,
    }


# =============================================================================
# NEW: Export Download — builds and streams the actual file
# =============================================================================

@app.post("/export/download")
async def export_download(req: ExportDownloadRequest):
    """
    Build the final file (PDF or PPTX) from pre-generated content
    and stream it back as a download.
    """
    print(f"⬇️  [export/download] format={req.format} title={req.title}")

    try:
        if req.format == "pdf":
            file_bytes = _build_pdf(req.title, req.content)
            return StreamingResponse(
                io.BytesIO(file_bytes),
                media_type="application/pdf",
                headers={
                    "Content-Disposition": f'attachment; filename="{_safe_filename(req.title)}.pdf"'
                },
            )
        else:
            file_bytes = _build_pptx(req.title, req.content)
            return StreamingResponse(
                io.BytesIO(file_bytes),
                media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                headers={
                    "Content-Disposition": f'attachment; filename="{_safe_filename(req.title)}.pptx"'
                },
            )
    except Exception as exc:
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=f"File generation error: {exc}")


def _safe_filename(title: str) -> str:
    """Convert title to safe filename."""
    return re.sub(r'[^\w\-]', '_', title)[:50]


# =============================================================================
# Full-dataset extraction — parses embedded JSON blocks from tool results
# =============================================================================

def _extract_full_dataset(conversation: List[ChatMessage]) -> dict | None:
    """
    Scan assistant messages for the embedded FULL_DATASET_JSON comment block
    injected by run_clickhouse_query. Returns the most recent dataset found
    (the last query result in the conversation), or None if not present.

    The block format is:
        <!-- FULL_DATASET_JSON
        {"total_rows": N, "columns": [...], "rows": [...]}
        END_FULL_DATASET_JSON -->
    """
    import json as _json
    import re as _re

    pattern = _re.compile(
        r'<!--\s*FULL_DATASET_JSON\s*(.*?)\s*END_FULL_DATASET_JSON\s*-->',
        _re.DOTALL
    )

    last_dataset = None
    for msg in conversation:
        if msg.role != "assistant":
            continue
        matches = pattern.findall(msg.content)
        if matches:
            # Take the last match in this message (most recent query)
            try:
                last_dataset = _json.loads(matches[-1].strip())
            except Exception:
                pass  # malformed block — skip

    return last_dataset


def _dataset_to_csv(dataset: dict) -> str:
    """
    Convert the extracted dataset dict to a UTF-8 CSV string.
    Handles both list-of-dicts and list-of-lists row formats.
    """
    import csv
    import io as _io

    buf = _io.StringIO()
    rows    = dataset.get("rows", [])
    columns = dataset.get("columns", [])

    if not rows:
        return "No data available.\n"

    # Determine column names
    if columns:
        fieldnames = columns
    elif rows and isinstance(rows[0], dict):
        fieldnames = list(rows[0].keys())
    else:
        fieldnames = [f"col_{i}" for i in range(len(rows[0]) if rows else 0)]

    writer = csv.DictWriter(buf, fieldnames=fieldnames, extrasaction="ignore",
                            lineterminator="\n")
    writer.writeheader()

    for row in rows:
        if isinstance(row, dict):
            writer.writerow(row)
        elif isinstance(row, (list, tuple)):
            writer.writerow(dict(zip(fieldnames, row)))

    return buf.getvalue()


# =============================================================================
# NEW: CSV Export — streams the full raw dataset from the last query
# =============================================================================

@app.post("/export/csv")
async def export_csv(req: ExportCSVRequest):
    """
    Extract the full (un-truncated) dataset embedded in the conversation by
    run_clickhouse_query, and stream it as a downloadable CSV file.
    All rows are exported — not just the 100-row chat preview.
    """
    if not req.conversation:
        raise HTTPException(status_code=400, detail="No conversation to export.")

    dataset = _extract_full_dataset(req.conversation)

    if dataset is None:
        raise HTTPException(
            status_code=404,
            detail=(
                "No query result found in this conversation. "
                "Ask a deal-list question first (e.g. 'list all active pipeline deals'), "
                "then export."
            ),
        )

    total = dataset.get("total_rows", len(dataset.get("rows", [])))
    print(f"📊 [export/csv] exporting {total} rows")

    csv_text = _dataset_to_csv(dataset)

    return StreamingResponse(
        iter([csv_text.encode("utf-8")]),
        media_type="text/csv",
        headers={
            "Content-Disposition": f'attachment; filename="{_safe_filename(req.title)}.csv"',
            "X-Total-Rows": str(total),
        },
    )




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
    "summary":   colors.HexColor("#0D1B3E"),
    "analysis":  colors.HexColor("#1565C0"),
    "overview":  colors.HexColor("#004D40"),
    "insight":   colors.HexColor("#1B5E20"),
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
    # Escape XML-reserved characters for ReportLab
    t = t.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
    return t.strip()


def _pdf_styles():
    return {
        "Cover_Title": ParagraphStyle("Cover_Title", fontSize=26, leading=32,
                                       textColor=_C_WHITE, fontName="Helvetica-Bold", spaceAfter=8),
        "Cover_Sub":   ParagraphStyle("Cover_Sub",   fontSize=13, leading=18,
                                       textColor=colors.HexColor("#B0BEC5"), fontName="Helvetica"),
        "Section_H":   ParagraphStyle("Section_H",   fontSize=11, leading=15,
                                       textColor=_C_WHITE, fontName="Helvetica-Bold"),
        "Body":   ParagraphStyle("Body",   fontSize=9, leading=14, textColor=_C_TXT,
                                  fontName="Helvetica", spaceAfter=4),
        "Bullet": ParagraphStyle("Bullet", fontSize=9, leading=14, textColor=_C_TXT,
                                  fontName="Helvetica", leftIndent=12, firstLineIndent=-8, spaceAfter=3),
        "H2": ParagraphStyle("H2", fontSize=11, leading=15, textColor=_C_NAVY,
                               fontName="Helvetica-Bold", spaceBefore=10, spaceAfter=4),
        "H3": ParagraphStyle("H3", fontSize=9,  leading=13, textColor=_C_BLUE,
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
    buf      = io.BytesIO()
    styles   = _pdf_styles()
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
        canvas.drawCentredString(
            PW / 2, _MB + 5,
            f"DIUD Report  |  AI-Generated  |  CONFIDENTIAL  |  {date.today().strftime('%B %Y')}"
        )
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

    if not sections:
        # Fallback: just dump the text
        for line in report_text.split("\n"):
            line = line.strip()
            if line:
                story.append(Paragraph(_strip_md(line), styles["Body"]))
        story.append(Spacer(1, 6))
    else:
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
    "summary":   RGBColor(0x1E, 0x88, 0xE5),
    "analysis":  RGBColor(0x00, 0x89, 0x7B),
    "insight":   RGBColor(0x1B, 0x5E, 0x20),
    "executive": RGBColor(0x0D, 0x1B, 0x3E),
}


def _pptx_bg(slide, color):
    f = slide.background.fill; f.solid(); f.fore_color.rgb = color

def _pptx_rect(slide, l, t, w, h, color):
    shp = slide.shapes.add_shape(1, Inches(l), Inches(t), Inches(w), Inches(h))
    shp.fill.solid(); shp.fill.fore_color.rgb = color; shp.line.fill.background()
    return shp

def _pptx_txt(slide, text, l, t, w, h, bold=False, size=18, color=None, align=PP_ALIGN.LEFT):
    txb = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    txb.word_wrap = True
    tf = txb.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]; p.alignment = align
    run = p.add_run(); run.text = text
    run.font.size = Pt(size); run.font.bold = bold
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
    prs.slide_width = Inches(13.33); prs.slide_height = Inches(7.5)
    footer_text = f"DIUD  |  AI-Generated  |  CONFIDENTIAL  |  {date.today().strftime('%B %Y')}"
    blank = prs.slide_layouts[6]

    def _footer(s):
        _pptx_rect(s, 0, 7.1, 13.33, 0.4, _C_DNAV_P)
        _pptx_txt(s, footer_text, 0.3, 7.12, 12, 0.35, size=7, color=_C_DIM_P, align=PP_ALIGN.CENTER)

    def _accent(t):
        for k, c in _SLIDE_ACCENT.items():
            if k in t: return c
        return _C_BLUE_P

    cover = prs.slides.add_slide(blank)
    _pptx_bg(cover, _C_NAVY_P)
    _pptx_rect(cover, 0, 3.2, 13.33, 0.06, _C_BLUE_P)
    _pptx_txt(cover, title, 0.8, 1.6, 11.5, 1.4, bold=True, size=34, color=_C_WHITE_P)
    _pptx_txt(cover, "Deals Intelligence Report", 0.8, 3.0, 8, 0.6,
              size=15, color=RGBColor(0xB0, 0xBE, 0xC5))
    _pptx_txt(cover, f"Generated: {date.today().strftime('%B %d, %Y')}", 0.8, 3.6, 6, 0.45,
              size=12, color=RGBColor(0x78, 0x90, 0x9C))

    for i, (s_title, bullets) in enumerate(slides_data):
        slide = prs.slides.add_slide(blank)
        ac = _accent(s_title.lower())
        _pptx_bg(slide, _C_LTBG_P)
        _pptx_rect(slide, 0, 0, 13.33, 0.9, ac)
        _pptx_txt(slide, s_title.upper(), 0.35, 0.1, 12.5, 0.7, bold=True, size=18, color=_C_WHITE_P)
        _pptx_txt(slide, str(i+1), 12.5, 0.12, 0.6, 0.6, size=11, color=_C_WHITE_P, align=PP_ALIGN.RIGHT)
        _pptx_rect(slide, 0.3, 1.0, 12.73, 5.9, _C_WHITE_P)
        if bullets:
            txb = slide.shapes.add_textbox(Inches(0.5), Inches(1.1), Inches(12.3), Inches(5.6))
            txb.word_wrap = True
            tf = txb.text_frame; tf.word_wrap = True
            for j, bullet in enumerate(bullets[:12]):
                p = tf.add_paragraph() if j > 0 else tf.paragraphs[0]
                p.space_before = Pt(4)
                dot = p.add_run(); dot.text = "●  "; dot.font.size = Pt(8); dot.font.color.rgb = ac
                run = p.add_run(); run.text = bullet; run.font.size = Pt(12); run.font.color.rgb = _C_TXT_P
        else:
            _pptx_txt(slide, "No data available.", 0.5, 1.2, 12, 0.5, size=11, color=_C_DIM_P)
        _footer(slide)

    buf = io.BytesIO(); prs.save(buf)
    return buf.getvalue()
