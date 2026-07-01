
import csv
import io
import json
import os
import re
import traceback
import uuid
from rules import (
    validate_sql_against_rules,
    validate_result_against_rules,
    get_intent,
    get_rulebook_entry,
    validate_summary_against_facts,
)
from datetime import date, datetime, timedelta
from typing import Dict, List, Literal, Optional, Callable
from charts import build_chart_html, reply_already_has_chart

import httpx
import anthropic
from fastapi import FastAPI, Header, HTTPException
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
app = FastAPI(title="DIUD", description="Decision Intelligence Using Data", version="4.0.0")

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
_CLAUDE_MODEL = "claude-sonnet-4-6"
ALLOWED_MODELS = {
    "sonnet": "claude-sonnet-4-6",
    "opus":   "claude-opus-4-6",
}

# FIX 1: Increased max_tokens across all call sites
# Chat responses: 4096 → handles long analytical responses without mid-sentence cut
# Export generation: 6144 → handles full reports with large deal tables
CHAT_MAX_TOKENS   = 6144
EXPORT_MAX_TOKENS = 6144

# =============================================================================
# SERVER-SIDE SESSION STORE
# =============================================================================
class QueryResult:
    def __init__(self, sql: str, columns: List[str], rows: List[dict],
                 total_rows: int, captured_at: str, filters_applied: str = ""):
        self.sql             = sql
        self.columns         = columns
        self.rows            = rows
        self.total_rows      = total_rows
        self.captured_at     = captured_at
        self.filters_applied = filters_applied


import threading

_SESSION_STORE:      Dict[str, QueryResult] = {}
_SESSION_TIMESTAMPS: Dict[str, datetime]    = {}

def _store_result(session_id: str, result: QueryResult):
    _SESSION_STORE[session_id]      = result
    _SESSION_TIMESTAMPS[session_id] = datetime.utcnow()
    
def _get_result(session_id: str) -> Optional[QueryResult]:
    """Retrieve stored query result for a session. Returns None if not found."""
    return _SESSION_STORE.get(session_id)

def _cleanup_sessions():
    cutoff  = datetime.utcnow() - timedelta(hours=4)
    expired = [sid for sid, ts in list(_SESSION_TIMESTAMPS.items()) if ts < cutoff]
    for sid in expired:
        _SESSION_STORE.pop(sid, None)
        _SESSION_TIMESTAMPS.pop(sid, None)
    if expired:
        print(f"🧹 Cleaned {len(expired)} expired sessions.")
    threading.Timer(3600, _cleanup_sessions).start()

_RULE_AUDIT_LOG: List[dict] = []   # most-recent-first, capped

def _log_rule_audit(session_id: Optional[str], sql: str, violations: List[str],
                     stage: str, user_message: str):
    entry = {
        "ts": datetime.utcnow().isoformat(),
        "session_id": session_id,
        "stage": stage,                # "pre_execute" | "post_execute"
        "violations": violations,
        "sql_preview": sql[:300],
        "user_message": user_message[:200],
    }
    _RULE_AUDIT_LOG.insert(0, entry)
    del _RULE_AUDIT_LOG[200:]          # cap log size
    if violations:
        print(f"[RULE-AUDIT][{stage}] session={session_id} violations={violations}")

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
# Schema discovery
# =============================================================================
_LIVE_SCHEMA: dict = {}
_SCHEMA_BLOCK: str = "Schema not yet loaded."


def _proxy_get(path: str) -> dict | list | None:
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
        print(msg); _SCHEMA_BLOCK = msg; return msg

    if isinstance(databases_raw, list) and databases_raw:
        if isinstance(databases_raw[0], str):
            databases = databases_raw
        elif isinstance(databases_raw[0], dict):
            databases = [d.get("name") or d.get("database") or list(d.values())[0] for d in databases_raw]
        else:
            databases = [str(d) for d in databases_raw]
    elif isinstance(databases_raw, dict):
        databases = databases_raw.get("data") or databases_raw.get("databases") or list(databases_raw.values())[0] if databases_raw else []
    else:
        databases = []

    SKIP_DBS = {"system", "information_schema", "INFORMATION_SCHEMA"}
    databases = [d for d in databases if d not in SKIP_DBS]
    print(f"   Databases found: {databases}")

    schema_lines, schema_dict = [], {}

    for db in databases:
        tables_raw = _proxy_get(f"/tables/{db}")
        if not tables_raw:
            continue
        if isinstance(tables_raw, list) and tables_raw:
            tables = tables_raw if isinstance(tables_raw[0], str) else [t.get("name") or t.get("table") or list(t.values())[0] for t in tables_raw]
        elif isinstance(tables_raw, dict):
            tables = tables_raw.get("data") or tables_raw.get("tables") or list(tables_raw.values())[0] if tables_raw else []
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
                cols = schema_raw.get("columns") or schema_raw.get("data") or schema_raw.get("schema") or [schema_raw]
            else:
                cols = []

            schema_dict[f"{db}.{tbl}"] = cols
            col_lines = []
            for col in cols:
                if isinstance(col, dict):
                    col_name    = col.get("name") or col.get("column_name") or col.get("Field") or list(col.keys())[0]
                    col_type    = col.get("type") or col.get("data_type") or col.get("Type") or ""
                    col_comment = col.get("comment") or col.get("Comment") or ""
                    col_lines.append(f"  {col_name:<35} {col_type}" + (f"  — {col_comment}" if col_comment else ""))
                else:
                    col_lines.append(f"  {col}")
            schema_lines.append(f"\nTABLE: {db}.{tbl}")
            schema_lines.extend(col_lines)

    _LIVE_SCHEMA  = schema_dict
    _SCHEMA_BLOCK = "\n".join(schema_lines) if schema_lines else "No tables found."
    print(f"✅ Schema loaded: {list(schema_dict.keys())}")
    return _SCHEMA_BLOCK


# =============================================================================
# System prompt
# =============================================================================
def _build_system_prompt() -> str:
    compact_lines = []
    for table_key, cols in _LIVE_SCHEMA.items():
        col_parts = []
        for col in cols:
            if isinstance(col, dict):
                name = col.get("name") or col.get("column_name") or col.get("Field") or list(col.keys())[0]
                typ  = col.get("type") or col.get("data_type") or col.get("Type") or ""
                col_parts.append(f"{name}:{typ}")
            else:
                col_parts.append(str(col))
        compact_lines.append(f"{table_key}({', '.join(col_parts)})")

    schema = "\n".join(compact_lines) or "Schema not yet loaded."
    if len(schema) > 20000:
        schema = schema[:20000] + "\n[schema truncated]"

    return f"""
You are DIUD (Decision Intelligence Using Data) — a conversational business intelligence assistant.
You convert natural language questions into ClickHouse SQL queries, retrieve live data, and deliver
executive-grade insights. You never fabricate numbers and never run destructive SQL.

RULE PRIORITY ORDER (highest → lowest):
  1. Greeting Rule (§1)
  2. Safety — SELECT/WITH only, no destructive SQL ever
  3. MANDATORY_BASE_FILTERS (§3)
  4. Tool Usage (§2)
  5. Business & SQL Rules (§4–§13)
  6. Formatting & Chart Rules (§14–§15)

═══════════════════════════════════════════════════════════════
§1  GREETING RULE — HIGHEST PRIORITY
═══════════════════════════════════════════════════════════════
If the user's message is ONLY a greeting (hi, hey, hello, good morning, etc.),
respond with EXACTLY this text and nothing else:

  "Hey, I'm DIUD, your data intelligence agent to help you analyse
  the live ClickHouse data. How may I help you?"

No bullet points, no extras. This rule overrides everything below.

═══════════════════════════════════════════════════════════════
§2  TOOL USAGE
═══════════════════════════════════════════════════════════════
You have LIVE access to a ClickHouse database via the query_clickhouse tool.
Use it for any question about pipeline deals, AEs, regions, industries, stages,
win/loss, competitors, conversions, or any metric not already confirmed in the
current conversation context.

NEVER use numbers from memory or a previous conversation turn.
Every metric must be queried live from the database.

DATABASE CONNECTION FAILURE:
If the tool returns "DATABASE CONNECTION FAILED", relay that error directly
to the user without attempting to answer from memory.

EXPORT INTENT RULE:
When the user asks to export, download, create a report/presentation, or
retrieve data in any format, output this marker on its own line:

  __EXPORT_INTENT__

Then on the next line write a brief confirmation such as:
  "Sure! Opening the export panel — select your format and detail level, then click Generate Preview."

This applies to ALL of these phrasings (and similar ones):
  - "export this conversation as a PDF report"
  - "export as PDF" / "export as report"
  - "create a presentation" / "create a deck"
  - "export the list" / "give me those N deals"
  - "download this as CSV" / "save this as a report"

⚠️  NEVER say exporting is "outside your capabilities".
⚠️  NEVER suggest Ctrl+P, screenshots, copy-paste, or contacting an admin.
⚠️  NEVER treat export requests as questions about your features.
The export panel is fully functional — always emit __EXPORT_INTENT__ for any export request.
Do NOT re-run the query. The export panel handles format selection and download.

═══════════════════════════════════════════════════════════════
§3  SCHEMA, DUPLICATE EXCLUSION, AND MANDATORY BASE FILTERS
═══════════════════════════════════════════════════════════════

── DUPLICATE RECORD EXCLUSION — ALWAYS APPLY ──────────────────
Rule                           | Apply when
-------------------------------|------------------------------------------
FINAL keyword                  | Every query against hs_analytics.* tables
countDistinct(), never count() | All deal/contact aggregations
DISTINCT in subquery           | Association table subqueries
GROUP BY + SUM                 | All target table queries

── TABLE 1: hs_analytics.deals (always use FINAL) ─────────────
PURPOSE: Core deals fact table.
Key columns:
  deal_id, deal_name, deal_owner, deal_stage, deal_type, pipeline,
  amount, region, deal_source_rollup, kore_primary_industry,
  account_priority_level, create_date, close_date,
  became_5_deal_date, became_10_deal_date, became_20_deal_date,
  became_30_deal_date, became_40_deal_date, became_60_deal_date,
  became_75_deal_date

── TABLE 2: hs_analytics.owners (always use FINAL) ────────────
PURPOSE: AE/owner master data.
Columns: id, firstName, lastName, email

── TABLE 3: hs_analytics.companies (always use FINAL) ─────────
PURPOSE: Company/account master data.
Columns: company_id, name, domain, industry, country, city

── TABLE 4: hs_analytics.contacts (always use FINAL) ──────────
PURPOSE: Contact/lead master data.
Columns:
  contact_id, email, first_name, last_name, company_name,
  company_priority, region, original_source, lead_status,
  lifecycle_stage,
  date_entered_marketing_qualified_lead_lifecycle_stage_pipeline

── TABLE 5: kore_ai_hubspot.gs_DealContactAssociation ─────────
PURPOSE: Many-to-many link between contacts and deals.
Columns: contact_id, deal_id
Usage: Always use DISTINCT in subqueries against this table.

── TABLE 6: kore_ai_hubspot.gs_marketing_targets ──────────────
PURPOSE: Marketing MQL and pipeline targets by source.
Columns: fy, quarter, month, region, original_source, mql_target

── TABLE 7: kore_ai_hubspot.gs_deal_ids_hs ────────────────────
PURPOSE: Allowlist of valid deal IDs — used in MANDATORY_BASE_FILTERS.
Columns: deal_id_hs

── MANDATORY BASE FILTERS (apply to EVERY deals query) ────────
Every SQL query against hs_analytics.deals MUST include ALL of these:

  pipeline = 'default'
  AND deal_stage <> 'Duplicate Record'
  AND CASE WHEN deal_type IS NULL THEN 'Not Assigned' ELSE deal_type END
      NOT IN ('Partner-Led SMB')
  AND toInt64(deal_id) IN (
      SELECT DISTINCT toInt64(deal_id_hs) FROM kore_ai_hubspot.gs_deal_ids_hs
  )
  AND always use FINAL on every hs_analytics.* table reference
  AND always use countDistinct(deal_id) — never count() or count(deal_id)

FULL DEAL STAGE ALLOWLIST (required in all pipegen/funnel queries):
  '10% - Discovery', '20% - Solution', '30% - Proof',
  '40% - Proposal', '60% - Price Negotiation', '75% - Contract Review',
  '90% - Deal Desk Review', 'Closed Won', 'Closed Lost',
  'Didn''t Qualify', 'Prospect Disengaged', 'Deal on Hold'

═══════════════════════════════════════════════════════════════
§4  TARGET TABLES — SCHEMA, TIERS, AND CASTING RULES
═══════════════════════════════════════════════════════════════

── TARGET TIER DEFAULT RULE — CRITICAL ────────────────────────
Three tiers exist: L2 (base/default), L1 (stretch), Committed.
DEFAULT: Always use L2 targets unless the user explicitly says
"L1", "stretch", or "committed". Never mix tiers in one query
unless the user explicitly asks for a comparison.

── NULLABLE STRING CASTING — MANDATORY ────────────────────────
ALL columns in every target table are Nullable(String) in ClickHouse.
ALWAYS cast before any arithmetic:
  CORRECT:   SUM(toFloat64OrZero(amount_target_20))
  INCORRECT: SUM(amount_target_20)   ← will cause silent null or type error

── COLUMN NAMING BY TIER — T1: gs_pipeline_quotas_v1 ──────────
Tier       | Column pattern        | Example
-----------|-----------------------|---------------------------
L2 DEFAULT | (no prefix/suffix)    | amount_target_20
L1         | suffix _l1            | amount_target_20_l1
Committed  | suffix _committed     | amount_target_20_committed

── COLUMN NAMING BY TIER — T2: gs_partner_targets_region_wise ─
Tier       | Column pattern        | Example
-----------|-----------------------|---------------------------
L2 DEFAULT | prefix l2_            | l2_amount_target_20
L1         | prefix l1_            | l1_amount_target_20
Committed  | prefix committed_     | committed_amount_target_20

── T1: kore_ai_hubspot.gs_pipeline_quotas_v1 ──────────────────
PURPOSE: Org-wide pipeline targets by region, source, funnel stage.
USE FOR: Pipeline attainment, EOP tracking, coverage ratio, gap-to-target.
COLUMNS (all Nullable(String) except id):
  id, fy, quarter, month, monthly_share, quarterly_share,
  region, regional_share, source, source_share

  L2 DEFAULT: amount_target_20, deals_target_20,
              amount_target_10, deals_target_10,
              amount_target_5,  deals_target_5

  L1 ONLY:    amount_target_20_l1, deals_target_20_l1,
              amount_target_10_l1, deals_target_10_l1,
              amount_target_5_l1,  deals_target_5_l1

  COMMITTED:  amount_target_20_committed, deals_target_20_committed,
              amount_target_10_committed, deals_target_10_committed,
              amount_target_5_committed,  deals_target_5_committed

── T2: kore_ai_hubspot.gs_partner_targets_region_wise ─────────
PURPOSE: Region-level partner pipeline targets by partner type.
USE FOR: Partner pipeline attainment by region, hyperscaler splits.
COLUMNS (all Nullable(String) except id):
  id, fy, quarter, month, region, regional_split,
  partner_team, partner_team_type, hyperscaler_type, amount_pk

  L2 DEFAULT: l2_amount_target_20, l2_deals_target_20,
              l2_amount_target_10, l2_deals_target_10,
              l2_amount_target_5,  l2_deals_target_5

  L1 ONLY:    l1_amount_target_20, l1_deals_target_20,
              l1_amount_target_10, l1_deals_target_10,
              l1_amount_target_5,  l1_deals_target_5

  COMMITTED:  committed_amount_target_20, committed_deals_target_20/10/5
  ⚠️  committed_amount_target_10 and committed_amount_target_5 do NOT exist.

  HYPERSCALER C1: msft_c1_targets_20, msft_c1_amount_target_20,
                  msft_c1_targets_10, msft_c1_targets_5,
                  aws_c1_targets_20,  aws_c1_amount_target_20,
                  aws_c1_targets_10,  aws_c1_targets_5

── T3: kore_ai_hubspot.gs_partner_targets_psd ─────────────────
PURPOSE: PSD (Partner Sales Director) level partner targets.
USE FOR: PSD quota attainment, individual PSD performance.
COLUMNS (all Nullable(String) except id):
  id, fy, quarter, month, region, partner_team,
  psd, hyperscaler_type, amount_primary_key

  COMMITTED ONLY (no L1/L2 in this table):
    committed_amount_target_20/10/5, committed_deals_target_20/10/5

  NOTE: For L1/L2 PSD-level targets, use T2 filtered by partner_team.

── T4: kore_ai_hubspot.gs_marketing_targets ───────────────────
PURPOSE: Marketing MQL and pipeline targets by source.
USE FOR: MQL attainment, marketing-sourced pipeline vs target.
COLUMNS (all Nullable(String) except id):
  id, fy, quarter, month, monthly_share, quarterly_share,
  region, regional_share, original_source, source_share

  L2 DEFAULT: amount_target_20, deals_target_20,
              amount_target_10, deals_target_10,
              amount_target_5,  deals_target_5, mql_target

  L1 ONLY:    l1_mql_target, l1_deals_target_20,
              l1_deals_target_10, l1_deals_target_5

  NOTE: No Committed tier and no L1 Amount columns in this table.

── T5: kore_ai_hubspot.gs_closed_won_quotas ───────────────────
PURPOSE: Closed Won revenue quotas by AE.
USE FOR: CW attainment %, AE-level quota tracking.
COLUMNS (all Nullable(String) except id):
  fy, quarter, month, region,
  ae                       — AE name; JOIN to hs_analytics.deals.deal_owner
  role, manager,
  assigned_amount_quota, assigned_deals_quota,
  annualized_amount_quota, annualized_deals_quota

  NOTE: Single quota tier only — no L1/L2/Committed split.
        Always cast: toFloat64OrZero(assigned_amount_quota)

── ATTAINMENT FORMULA ─────────────────────────────────────────
  attainment_pct = round(actual / nullIf(target, 0) * 100, 1)
  coverage_ratio = round(pipeline / nullIf(revenue_target, 0), 1)
  Always use nullIf(denominator, 0) to avoid divide-by-zero.

── TARGET QUERY RULES ─────────────────────────────────────────
1. NEVER join raw deal rows directly to target rows — use independent CTEs.
2. NEVER derive a quarterly target by dividing annual by 4.
3. ALWAYS filter targets to the exact quarter: WHERE fy='FY27' AND quarter='Q1'
4. Period grain must match: if actuals are Q1 FY27, filter target table to same.
5. For partner tables: filter partner_team_type IN ('Hyperscaler','GSI/SI','Reseller/BPO/TSD').
6. IF A TARGET QUERY RETURNS 0 ROWS: do not explain your plan in prose first.
   In the SAME tool-use round, immediately run a diagnostic query
     SELECT DISTINCT fy, quarter FROM <target_table>
   against the relevant target table, identify the correct fy string format
   (e.g. 'FY27' vs '27' vs '2027'), then re-run the original attainment
   query with the corrected filter — all before writing any response to
   the user. Never tell the user "let me check X" without having already
   called the tool to check X in that same turn.

═══════════════════════════════════════════════════════════════
§5  THREE QUERY PATTERNS — DECISION GUIDE
═══════════════════════════════════════════════════════════════
Before writing ANY SQL, determine which pattern applies:

PATTERN A — CUMULATIVE PIPEGEN / FUNNEL STAGE COUNTS
  Use when the user asks:
  • "how many deals reached [stage]" / "pipegen at 10%, 20%, 30%..."
  • "funnel breakdown", "stage counts", "pipeline funnel"
  • "deals created in Q1", "10% created", "20% created"
  • "conversion from X% to Y%", "funnel conversion rate"
  • "deals by region/source/industry at each stage"
  KEY: A deal is counted at stage N if it has EVER reached N or beyond.
       FY/quarter is ALWAYS anchored to became_10_deal_date (pipeline entry point),
       regardless of which stage is being counted. This is the cohort definition
       used in Looker. Stage counting uses cumulative OR chains, NOT cohort exclusions.
  See §6 for full SQL pattern.

PATTERN B — DEAL-LEVEL DETAIL / ACTIVE PIPELINE VIEW
  Use when the user asks:
  • "show me the deals", "list all deals", "deal details"
  • "active pipeline" as individual rows
  • "days in stage", "stalled deals", "deal health"
  • "BANT status", "last contacted", "AE view of deals"
  • "forecast", "management forecast", "ae_forecast"
  KEY: One row per deal. Primary filter is close_date.
       FY computed from became_10_deal_date (for "qualified_in")
       AND from close_date (for "closing_in").
  See §7 for full SQL pattern.

PATTERN C — ACTUALS vs TARGET / ATTAINMENT
  Use when the user asks:
  • "attainment", "vs target", "quota", "coverage ratio"
  • "are we on track", "gap to target", "EOP tracking"
  • "pipegen target", "10% target", "20% target attainment"
  • "which regions are below quota"
  KEY: Two independent CTEs — actuals from deals, targets from quota table.
       NEVER fan-out join deals directly to target rows.
       The became_<N>_deal_date anchor MATCHES the stage being targeted:
         - 10% targets → became_10_deal_date
         - 20% targets → became_20_deal_date
         - 5% targets  → became_5_deal_date
       Source mapping MERGES Executive Outreach + Investor (to match quota table).
  See §8 for full SQL pattern.

PATTERN D — GENERAL / AD-HOC QUERIES
  Use when the question does NOT match Pattern A, B, or C.
  Examples:
  • Win rate by source/region, average deal size, sales cycle length
  • BANT qualification rate, competitor mentions, AE rankings
  • Cross-table joins (contacts → deals via association table)
  • MQL-to-deal funnels, any analytical question answerable with SQL
  
  KEY RULES:
  - Apply MANDATORY_BASE_FILTERS (§3) when querying deals.
  - Apply MQL filters (§11) when querying contacts.
  - Use FINAL on all hs_analytics.* tables.
  - Use countDistinct() for all aggregations.
  - Start SQL with WITH or SELECT (no leading comments needed).
  - Use appropriate JOINs between tables as needed.
  - For association table joins, always use DISTINCT in subqueries.

PATTERN SELECTION QUICK REFERENCE:

"how many deals reached/passed through [stage]"   → A
"pipegen at [stage]", "[stage] created"            → A
"funnel breakdown", "stage counts", "conversion"   → A
"show me the deals", "list active pipeline"        → B
"stalled deals", "days in stage", "deal health"    → B
"BANT status", "last contacted", "AE deal list"    → B
"attainment", "vs target", "quota", "on track"     → C
"pipegen target", "EOP target", "coverage"         → C
"gap to target", "which regions below quota"       → C

AMBIGUOUS: "pipegen attainment by region"          → C (needs target table)
AMBIGUOUS: "how many 20% deals vs target"          → C
AMBIGUOUS: "active pipeline vs EOP"                → C

"win rate", "avg deal size", "sales cycle"             → D
"BANT rate", "competitor analysis", "AE ranking"       → D
"MQL to deal", "contact to deal funnel"                → D
Any question not matching A, B, or C                   → D

═══════════════════════════════════════════════════════════════
§6  PATTERN A — CUMULATIVE PIPEGEN / FUNNEL STAGE COUNTS
═══════════════════════════════════════════════════════════════

── CORE LOGIC ────────────────────────────────────────────────
FY/quarter is ALWAYS anchored to became_10_deal_date for ALL stages.
A deal is counted at stage N if it has EVER reached N or beyond (cumulative OR).
Stage counting conditions use OR chains — NOT cohort exclusions.

── COUNTING CONDITIONS BY STAGE ──────────────────────────────

10% (ever reached 10% or beyond):
  (became_10_deal_date != '1900-01-01'
   OR became_20_deal_date != '1900-01-01'
   OR became_30_deal_date != '1900-01-01'
   OR became_40_deal_date != '1900-01-01'
   OR became_60_deal_date != '1900-01-01'
   OR became_75_deal_date != '1900-01-01'
   OR deal_stage IN ('10% - Discovery','20% - Solution','30% - Proof',
                     '40% - Proposal','60% - Price Negotiation',
                     '75% - Contract Review','90% - Deal Desk Review','Closed Won'))

20% (ever reached 20% or beyond):
  (became_20_deal_date != '1900-01-01'
   OR became_30_deal_date != '1900-01-01'
   OR became_40_deal_date != '1900-01-01'
   OR became_60_deal_date != '1900-01-01'
   OR became_75_deal_date != '1900-01-01'
   OR deal_stage IN ('20% - Solution','30% - Proof','40% - Proposal',
                     '60% - Price Negotiation','75% - Contract Review',
                     '90% - Deal Desk Review','Closed Won'))

30% (ever reached 30% or beyond):
  (became_30_deal_date != '1900-01-01'
   OR became_40_deal_date != '1900-01-01'
   OR became_60_deal_date != '1900-01-01'
   OR became_75_deal_date != '1900-01-01'
   OR deal_stage IN ('30% - Proof','40% - Proposal','60% - Price Negotiation',
                     '75% - Contract Review','90% - Deal Desk Review','Closed Won'))

40% (ever reached 40% or beyond):
  (became_40_deal_date != '1900-01-01'
   OR became_60_deal_date != '1900-01-01'
   OR became_75_deal_date != '1900-01-01'
   OR deal_stage IN ('40% - Proposal','60% - Price Negotiation',
                     '75% - Contract Review','90% - Deal Desk Review','Closed Won'))

60% (ever reached 60% or beyond):
  (became_60_deal_date != '1900-01-01'
   OR became_75_deal_date != '1900-01-01'
   OR deal_stage IN ('60% - Price Negotiation','75% - Contract Review',
                     '90% - Deal Desk Review','Closed Won'))

75% (ever reached 75% or beyond):
  (became_75_deal_date != '1900-01-01'
   OR deal_stage IN ('75% - Contract Review','90% - Deal Desk Review','Closed Won'))

Closed Won:
  deal_stage IN ('90% - Deal Desk Review','Closed Won')

── SOURCE MAPPING FOR PATTERN A/B ───────────────────────────
Executive Outreach and Investor are kept SEPARATE in Pattern A and B:

  CASE
    WHEN deal_source_rollup IN ('Marketing','Customer Success','Executive Outreach',
         'AE Outbound','Investor','Inception','Hyperscaler') THEN deal_source_rollup
    WHEN deal_source_rollup = 'BDR Outbound' THEN 'BDR'
    WHEN deal_source_rollup = 'Partner'       THEN 'Partner - Non Hyperscaler'
    ELSE 'Other'
  END AS deal_source_rollup

── SINGLE-STAGE SIMPLIFICATION ───────────────────────────────
For a single-stage question, run only that stage's SELECT — no UNION ALL needed.

── CONVERSION RATE ───────────────────────────────────────────
  conversion_N_to_M % = count_at_M / count_at_N * 100
Run two separate counts using the conditions above, then divide.

── COMPLETE SQL TEMPLATE FOR PATTERN A ───────────────────────

WITH pipe_gen AS (
  SELECT
    toInt64(deal_id)  AS deal_id,
    CAST(LEFT(coalesce(create_date,'1900-01-01'),10) AS DATE)          AS create_date,
    CAST(LEFT(coalesce(close_date,'1900-01-01'),10) AS DATE)           AS close_date,
    CAST(LEFT(coalesce(became_10_deal_date,'1900-01-01'),10) AS DATE)  AS became_10_deal_date,
    CAST(LEFT(coalesce(became_20_deal_date,'1900-01-01'),10) AS DATE)  AS became_20_deal_date,
    CAST(LEFT(coalesce(became_30_deal_date,'1900-01-01'),10) AS DATE)  AS became_30_deal_date,
    CAST(LEFT(coalesce(became_40_deal_date,'1900-01-01'),10) AS DATE)  AS became_40_deal_date,
    CAST(LEFT(coalesce(became_60_deal_date,'1900-01-01'),10) AS DATE)  AS became_60_deal_date,
    CAST(LEFT(coalesce(became_75_deal_date,'1900-01-01'),10) AS DATE)  AS became_75_deal_date,
    deal_stage, deal_type,
    CASE WHEN region='japac' THEN 'JAPAC' WHEN region='Africa' THEN 'Middle East'
         WHEN region='india___sea' THEN 'ISEA' ELSE region END AS region,
    amount,
    CASE
      WHEN deal_source_rollup IN ('Marketing','Customer Success','Executive Outreach',
           'AE Outbound','Investor','Inception','Hyperscaler') THEN deal_source_rollup
      WHEN deal_source_rollup = 'BDR Outbound' THEN 'BDR'
      WHEN deal_source_rollup = 'Partner'       THEN 'Partner - Non Hyperscaler'
      ELSE 'Other'
    END AS deal_source_rollup,
    CASE WHEN account_priority_level IN ('P1','P2','P3','P4') THEN 'P1-P4'
         WHEN account_priority_level IN ('P5','P6','P7')      THEN 'P5-P7'
         WHEN account_priority_level IN ('P8','P9','P10')     THEN 'P8-P10'
         ELSE NULL END AS account_priority_level,
    ai_for_x,
    CASE WHEN kore_primary_industry IN ('Financial Services','Banking','Insurance')
              THEN 'Financial Services'
         WHEN kore_primary_industry IN ('Manufacturing Discreet','Manufacturing Process','CPG')
              THEN 'Manufacturing'
         WHEN kore_primary_industry IN ('Hi-Tech','Telecom / Media / Entertainment')
              THEN 'TMT'
         WHEN kore_primary_industry IN ('Business Services','Government','Energy & Utilities',
              'Education','Restaurants','null','Energy') THEN 'Other'
         ELSE kore_primary_industry END AS kore_primary_industry,
    CASE WHEN is_this_a_deal_with_inception='Yes' THEN 'Yes' ELSE 'No' END
         AS is_this_a_deal_with_inception
  FROM hs_analytics.deals d FINAL
  LEFT JOIN hs_analytics.owners o FINAL ON d.deal_owner = CAST(o.id AS VARCHAR)
  WHERE became_10_deal_date >= '2025-04-01'
    AND deal_stage <> 'Duplicate Record'
    AND deal_stage IN (
      '10% - Discovery','20% - Solution','30% - Proof','40% - Proposal',
      '60% - Price Negotiation','75% - Contract Review','90% - Deal Desk Review',
      'Closed Won','Closed Lost','Didn''t Qualify','Prospect Disengaged','Deal on Hold'
    )
    AND pipeline = 'default'
    AND toInt64(d.deal_id) IN (
      SELECT DISTINCT toInt64(deal_id_hs) FROM kore_ai_hubspot.gs_deal_ids_hs
    )
),
stage_base AS (
  -- FY/quarter ALWAYS from became_10_deal_date regardless of stage being counted
  SELECT DISTINCT *,
    toYear(became_10_deal_date) + if(toMonth(became_10_deal_date)>=4,1,0) AS create_fy,
    CASE WHEN toMonth(became_10_deal_date) IN (1,2,3)    THEN 'Q4'
         WHEN toMonth(became_10_deal_date) IN (4,5,6)    THEN 'Q1'
         WHEN toMonth(became_10_deal_date) IN (7,8,9)    THEN 'Q2'
         WHEN toMonth(became_10_deal_date) IN (10,11,12) THEN 'Q3'
    END AS create_quarter,
    LEFT(formatDateTime(became_10_deal_date,'%M'),3) AS create_month,
    toYear(close_date) + if(toMonth(close_date)>=4,1,0) AS close_fy,
    CASE WHEN toMonth(close_date) IN (1,2,3)    THEN 'Q4'
         WHEN toMonth(close_date) IN (4,5,6)    THEN 'Q1'
         WHEN toMonth(close_date) IN (7,8,9)    THEN 'Q2'
         WHEN toMonth(close_date) IN (10,11,12) THEN 'Q3'
    END AS close_quarter,
    LEFT(formatDateTime(close_date,'%M'),3) AS close_month,
    toYear(create_date) + if(toMonth(create_date)>=4,1,0) AS hs_create_fy,
    CASE WHEN toMonth(create_date) IN (1,2,3)    THEN 'Q4'
         WHEN toMonth(create_date) IN (4,5,6)    THEN 'Q1'
         WHEN toMonth(create_date) IN (7,8,9)    THEN 'Q2'
         WHEN toMonth(create_date) IN (10,11,12) THEN 'Q3'
    END AS hs_create_quarter,
    LEFT(formatDateTime(create_date,'%M'),3) AS hs_create_month
  FROM pipe_gen
  WHERE CASE WHEN deal_type IS NULL THEN 'Not Assigned' ELSE deal_type END
        NOT IN ('Partner-Led SMB')
)

-- 10% stage count
SELECT '10% - Discovery' AS stage,
       create_fy, create_quarter, create_month,
       deal_source_rollup, region, ai_for_x, account_priority_level,
       is_this_a_deal_with_inception,
       countDistinct(deal_id) AS deals
FROM stage_base
WHERE create_fy >= 2025
  AND (became_10_deal_date != '1900-01-01'
       OR became_20_deal_date != '1900-01-01'
       OR became_30_deal_date != '1900-01-01'
       OR became_40_deal_date != '1900-01-01'
       OR became_60_deal_date != '1900-01-01'
       OR became_75_deal_date != '1900-01-01'
       OR deal_stage IN ('10% - Discovery','20% - Solution','30% - Proof',
                         '40% - Proposal','60% - Price Negotiation',
                         '75% - Contract Review','90% - Deal Desk Review','Closed Won'))
GROUP BY 1,2,3,4,5,6,7,8,9

UNION ALL

-- 20% stage count
SELECT '20% - Solution' AS stage,
       create_fy, create_quarter, create_month,
       deal_source_rollup, region, ai_for_x, account_priority_level,
       is_this_a_deal_with_inception,
       countDistinct(deal_id) AS deals
FROM stage_base
WHERE create_fy >= 2025
  AND (became_20_deal_date != '1900-01-01'
       OR became_30_deal_date != '1900-01-01'
       OR became_40_deal_date != '1900-01-01'
       OR became_60_deal_date != '1900-01-01'
       OR became_75_deal_date != '1900-01-01'
       OR deal_stage IN ('20% - Solution','30% - Proof','40% - Proposal',
                         '60% - Price Negotiation','75% - Contract Review',
                         '90% - Deal Desk Review','Closed Won'))
GROUP BY 1,2,3,4,5,6,7,8,9

UNION ALL

-- 30% stage count
SELECT '30% - Proof' AS stage,
       create_fy, create_quarter, create_month,
       deal_source_rollup, region, ai_for_x, account_priority_level,
       is_this_a_deal_with_inception,
       countDistinct(deal_id) AS deals
FROM stage_base
WHERE create_fy >= 2025
  AND (became_30_deal_date != '1900-01-01'
       OR became_40_deal_date != '1900-01-01'
       OR became_60_deal_date != '1900-01-01'
       OR became_75_deal_date != '1900-01-01'
       OR deal_stage IN ('30% - Proof','40% - Proposal','60% - Price Negotiation',
                         '75% - Contract Review','90% - Deal Desk Review','Closed Won'))
GROUP BY 1,2,3,4,5,6,7,8,9

UNION ALL

-- 40% stage count
SELECT '40% - Proposal' AS stage,
       create_fy, create_quarter, create_month,
       deal_source_rollup, region, ai_for_x, account_priority_level,
       is_this_a_deal_with_inception,
       countDistinct(deal_id) AS deals
FROM stage_base
WHERE create_fy >= 2025
  AND (became_40_deal_date != '1900-01-01'
       OR became_60_deal_date != '1900-01-01'
       OR became_75_deal_date != '1900-01-01'
       OR deal_stage IN ('40% - Proposal','60% - Price Negotiation',
                         '75% - Contract Review','90% - Deal Desk Review','Closed Won'))
GROUP BY 1,2,3,4,5,6,7,8,9

UNION ALL

-- 60% stage count
SELECT '60% - Price Negotiation' AS stage,
       create_fy, create_quarter, create_month,
       deal_source_rollup, region, ai_for_x, account_priority_level,
       is_this_a_deal_with_inception,
       countDistinct(deal_id) AS deals
FROM stage_base
WHERE create_fy >= 2025
  AND (became_60_deal_date != '1900-01-01'
       OR became_75_deal_date != '1900-01-01'
       OR deal_stage IN ('60% - Price Negotiation','75% - Contract Review',
                         '90% - Deal Desk Review','Closed Won'))
GROUP BY 1,2,3,4,5,6,7,8,9

UNION ALL

-- 75% stage count
SELECT '75% - Contract Review' AS stage,
       create_fy, create_quarter, create_month,
       deal_source_rollup, region, ai_for_x, account_priority_level,
       is_this_a_deal_with_inception,
       countDistinct(deal_id) AS deals
FROM stage_base
WHERE create_fy >= 2025
  AND (became_75_deal_date != '1900-01-01'
       OR deal_stage IN ('75% - Contract Review','90% - Deal Desk Review','Closed Won'))
GROUP BY 1,2,3,4,5,6,7,8,9

UNION ALL

-- Closed Won count
SELECT 'Closed Won' AS stage,
       create_fy, create_quarter, create_month,
       deal_source_rollup, region, ai_for_x, account_priority_level,
       is_this_a_deal_with_inception,
       countDistinct(deal_id) AS deals
FROM stage_base
WHERE create_fy >= 2025
  AND deal_stage IN ('90% - Deal Desk Review','Closed Won')
GROUP BY 1,2,3,4,5,6,7,8,9

═══════════════════════════════════════════════════════════════
§7  PATTERN B — DEAL-LEVEL DETAIL / ACTIVE PIPELINE VIEW
═══════════════════════════════════════════════════════════════

── KEY LOGIC ─────────────────────────────────────────────────
• Primary filter is close_date (not became_10_deal_date)
• Two FY dimensions:
    create_fy  = toYear(became_10_deal_date) + if(toMonth(became_10_deal_date)>=4,1,0)
    close_fy   = toYear(close_date) + if(toMonth(close_date)>=4,1,0)
• "qualified_in" = concat(create_fy,' | ',create_quarter)  [from became_10_deal_date]
• "closing_in"   = concat(close_fy,' | ',close_quarter)    [from close_date]

── DAYS IN CURRENT STAGE ─────────────────────────────────────
  CASE
    WHEN deal_stage IN ('Prospect Disengaged','Closed Lost','Didn''t Qualify',
                        '90% - Deal Desk Review','Closed Won') THEN NULL
    WHEN deal_stage = '5% - IQM Held'  AND became_5_deal_date  <>'1900-01-01'
         THEN DATE_DIFF('Day', became_5_deal_date,  CURRENT_DATE())
    WHEN deal_stage = '10% - Discovery' AND became_10_deal_date <>'1900-01-01'
         THEN DATE_DIFF('Day', became_10_deal_date, CURRENT_DATE())
    WHEN deal_stage = '20% - Solution'  AND became_20_deal_date <>'1900-01-01'
         THEN DATE_DIFF('Day', became_20_deal_date, CURRENT_DATE())
    WHEN deal_stage = '30% - Proof'     AND became_30_deal_date <>'1900-01-01'
         THEN DATE_DIFF('Day', became_30_deal_date, CURRENT_DATE())
    WHEN deal_stage = '40% - Proposal'  AND became_40_deal_date <>'1900-01-01'
         THEN DATE_DIFF('Day', became_40_deal_date, CURRENT_DATE())
    WHEN deal_stage = '60% - Price Negotiation' AND became_60_deal_date <>'1900-01-01'
         THEN DATE_DIFF('Day', became_60_deal_date, CURRENT_DATE())
    WHEN deal_stage = '75% - Contract Review'   AND became_75_deal_date <>'1900-01-01'
         THEN DATE_DIFF('Day', became_75_deal_date, CURRENT_DATE())
    ELSE NULL
  END AS days_in_current_stage

── DEAL HEALTH BENCHMARKS ────────────────────────────────────
  Stage                   | Benchmark | Green    | Yellow   | Red
  ------------------------|-----------|----------|----------|--------
  1% - IQM Scheduled      |  7 days   | < 10     | < 14     | >= 14
  5% - IQM Held           | 21 days   | < 31     | < 42     | >= 42
  10% - Discovery         | 28 days   | < 42     | < 56     | >= 56
  20% - Solution          | 41 days   | < 61     | < 82     | >= 82
  30% - Proof             | 15 days   | < 22     | < 30     | >= 30
  40% - Proposal          | 29 days   | < 43     | < 58     | >= 58
  60% - Price Negotiation | 27 days   | < 40     | < 54     | >= 54
  75% - Contract Review   | 34 days   | < 51     | < 68     | >= 68

── STAGE CATEGORY ROLLUP ─────────────────────────────────────
  Active Pipeline   → deal_stage IN ('20% - Solution','30% - Proof','40% - Proposal',
                                     '60% - Price Negotiation','75% - Contract Review')
  Fallen Out        → deal_stage IN ('Prospect Disengaged','Closed Lost','Didn''t Qualify')
  Closed Won        → deal_stage IN ('90% - Deal Desk Review','Closed Won')
  Pre-Qualification → everything else

── STAGE RANK (for sorting) ──────────────────────────────────
  1=5%, 2=10%, 3=20%, 4=30%, 5=40%, 6=60%, 7=75%, 8=90%,
  9=CW, 10=Deal on Hold, 11=Prospect Disengaged,
  12=Closed Lost, 13=Didn't Qualify

── BANT ──────────────────────────────────────────────────────
  CASE WHEN is_there_a_confirmation_of_budget='Yes'
        AND who_is_the_decision_maker IS NOT NULL
        AND use_case IS NOT NULL
        AND what_is_the_estimated_timeline IS NOT NULL THEN 'Yes'
       ELSE 'No' END AS BANT

── COMPLETE SQL TEMPLATE FOR PATTERN B ───────────────────────

WITH active_pipe AS (
  SELECT DISTINCT
    toInt64(d.deal_id)  AS deal_id,
    deal_name,
    concat(o.firstName,' ',o.lastName) AS deal_owner_name,
    CAST(LEFT(coalesce(d.create_date,'1900-01-01'),10) AS DATE)         AS create_date,
    CAST(LEFT(coalesce(close_date,'1900-01-01'),10) AS DATE)            AS close_date,
    CAST(LEFT(coalesce(became_5_deal_date,'1900-01-01'),10) AS DATE)    AS became_5_deal_date,
    CAST(LEFT(coalesce(became_10_deal_date,'1900-01-01'),10) AS DATE)   AS became_10_deal_date,
    CAST(LEFT(coalesce(became_20_deal_date,'1900-01-01'),10) AS DATE)   AS became_20_deal_date,
    CAST(LEFT(coalesce(became_30_deal_date,'1900-01-01'),10) AS DATE)   AS became_30_deal_date,
    CAST(LEFT(coalesce(became_40_deal_date,'1900-01-01'),10) AS DATE)   AS became_40_deal_date,
    CAST(LEFT(coalesce(became_60_deal_date,'1900-01-01'),10) AS DATE)   AS became_60_deal_date,
    CAST(LEFT(coalesce(became_75_deal_date,'1900-01-01'),10) AS DATE)   AS became_75_deal_date,
    deal_stage,
    CASE WHEN region='japac' THEN 'JAPAC' WHEN region='Africa' THEN 'Middle East'
         WHEN region='india___sea' THEN 'ISEA' ELSE region END AS deal_region,
    amount, country, pipeline,
    CASE
      WHEN deal_source_rollup IN ('Marketing','Customer Success','Executive Outreach',
           'AE Outbound','Investor','Inception','Hyperscaler') THEN deal_source_rollup
      WHEN deal_source_rollup = 'BDR Outbound' THEN 'BDR'
      WHEN deal_source_rollup = 'Partner'       THEN 'Partner - Non Hyperscaler'
      ELSE 'Other'
    END AS deal_source_rollup,
    CASE WHEN account_priority_level IN ('P1','P2','P3','P4') THEN 'P1-P4'
         WHEN account_priority_level IN ('P5','P6','P7')      THEN 'P5-P7'
         WHEN account_priority_level IN ('P8','P9','P10')     THEN 'P8-P10'
         ELSE NULL END AS account_priority_level,
    ai_for_x, kore_primary_industry, deal_url,
    CASE WHEN deal_type IS NULL THEN 'Not Assigned' ELSE deal_type END AS deal_type,
    CASE WHEN is_this_a_deal_with_inception='Yes' THEN 'Yes' ELSE 'No' END
         AS is_this_a_deal_with_inception,
    CASE WHEN is_there_a_confirmation_of_budget='Yes'
          AND who_is_the_decision_maker IS NOT NULL
          AND use_case IS NOT NULL
          AND what_is_the_estimated_timeline IS NOT NULL THEN 'Yes'
         ELSE 'No' END AS BANT,
    CAST(LEFT(coalesce(last_contacted,'1900-01-01'),10) AS DATE) AS last_contacted,
    t.name AS Team,
    forecast_amount, forecast_probability, management_forecast, ae_forecast
  FROM hs_analytics.deals d FINAL
  LEFT JOIN hs_analytics.owners o FINAL ON d.deal_owner = CAST(o.id AS VARCHAR)
  LEFT JOIN kore_ai_hubspot.gs_Teams t ON d.hubspot_team = t.team_id
  WHERE close_date >= '2025-04-01'
    AND deal_stage IN (
      '1% - IQM Scheduled','5% - IQM Held','10% - Discovery','20% - Solution',
      '30% - Proof','40% - Proposal','60% - Price Negotiation','75% - Contract Review',
      '90% - Deal Desk Review','Closed Won','Closed Lost','Didn''t Qualify',
      'Prospect Disengaged','Deal on Hold'
    )
    AND pipeline = 'default'
    AND toInt64(d.deal_id) IN (
      SELECT DISTINCT toInt64(deal_id_hs) FROM kore_ai_hubspot.gs_deal_ids_hs
    )
),
deal_detail AS (
  SELECT *,
    toYear(became_10_deal_date) + if(toMonth(became_10_deal_date)>=4,1,0) AS create_fy,
    CASE WHEN toMonth(became_10_deal_date) IN (1,2,3) THEN 'Q4'
         WHEN toMonth(became_10_deal_date) IN (4,5,6) THEN 'Q1'
         WHEN toMonth(became_10_deal_date) IN (7,8,9) THEN 'Q2'
         WHEN toMonth(became_10_deal_date) IN (10,11,12) THEN 'Q3' END AS create_quarter,
    toYear(close_date) + if(toMonth(close_date)>=4,1,0) AS close_fy,
    CASE WHEN toMonth(close_date) IN (1,2,3) THEN 'Q4'
         WHEN toMonth(close_date) IN (4,5,6) THEN 'Q1'
         WHEN toMonth(close_date) IN (7,8,9) THEN 'Q2'
         WHEN toMonth(close_date) IN (10,11,12) THEN 'Q3' END AS close_quarter,
    DATE_DIFF('Day',became_10_deal_date, became_20_deal_date) AS days_10_to_20,
    DATE_DIFF('Day',became_20_deal_date, became_30_deal_date) AS days_20_to_30,
    DATE_DIFF('Day',became_30_deal_date, became_40_deal_date) AS days_30_to_40,
    DATE_DIFF('Day',became_40_deal_date, became_60_deal_date) AS days_40_to_60,
    DATE_DIFF('Day',became_60_deal_date, became_75_deal_date) AS days_60_to_75,
    CASE
      WHEN deal_stage IN ('Prospect Disengaged','Closed Lost','Didn''t Qualify',
                          '90% - Deal Desk Review','Closed Won') THEN NULL
      WHEN deal_stage='5% - IQM Held'  AND became_5_deal_date<>'1900-01-01'
           THEN DATE_DIFF('Day',became_5_deal_date,CURRENT_DATE())
      WHEN deal_stage='10% - Discovery' AND became_10_deal_date<>'1900-01-01'
           THEN DATE_DIFF('Day',became_10_deal_date,CURRENT_DATE())
      WHEN deal_stage='20% - Solution'  AND became_20_deal_date<>'1900-01-01'
           THEN DATE_DIFF('Day',became_20_deal_date,CURRENT_DATE())
      WHEN deal_stage='30% - Proof'     AND became_30_deal_date<>'1900-01-01'
           THEN DATE_DIFF('Day',became_30_deal_date,CURRENT_DATE())
      WHEN deal_stage='40% - Proposal'  AND became_40_deal_date<>'1900-01-01'
           THEN DATE_DIFF('Day',became_40_deal_date,CURRENT_DATE())
      WHEN deal_stage='60% - Price Negotiation' AND became_60_deal_date<>'1900-01-01'
           THEN DATE_DIFF('Day',became_60_deal_date,CURRENT_DATE())
      WHEN deal_stage='75% - Contract Review'   AND became_75_deal_date<>'1900-01-01'
           THEN DATE_DIFF('Day',became_75_deal_date,CURRENT_DATE())
      ELSE NULL
    END AS days_in_current_stage
  FROM active_pipe
  WHERE CASE WHEN deal_type IS NULL THEN 'Not Assigned' ELSE deal_type END
        NOT IN ('Partner-Led SMB')
)
SELECT
  deal_id, deal_name, deal_owner_name, create_date, close_date,
  became_5_deal_date, became_10_deal_date, became_20_deal_date,
  became_30_deal_date, became_40_deal_date, became_60_deal_date, became_75_deal_date,
  deal_region AS region, amount, pipeline, deal_source_rollup, deal_url,
  account_priority_level, ai_for_x, kore_primary_industry,
  create_fy, create_quarter, close_fy, close_quarter,
  concat(create_fy,' | ',create_quarter) AS qualified_in,
  concat(close_fy,' | ',close_quarter)   AS closing_in,
  CASE
    WHEN deal_stage IN ('20% - Solution','30% - Proof','40% - Proposal',
                        '60% - Price Negotiation','75% - Contract Review') THEN 'Active Pipeline'
    WHEN deal_stage IN ('Prospect Disengaged','Closed Lost','Didn''t Qualify') THEN 'Fallen Out'
    WHEN deal_stage IN ('90% - Deal Desk Review','Closed Won') THEN 'Closed Won'
    ELSE 'Pre-Qualification'
  END AS stage_category,
  CASE WHEN (days_10_to_20 <0 OR days_10_to_20 >10000) THEN 0 ELSE days_10_to_20 END AS days_10_to_20,
  CASE WHEN (days_20_to_30 <0 OR days_20_to_30 >10000) THEN 0 ELSE days_20_to_30 END AS days_20_to_30,
  CASE WHEN (days_30_to_40 <0 OR days_30_to_40 >10000) THEN 0 ELSE days_30_to_40 END AS days_30_to_40,
  CASE WHEN (days_40_to_60 <0 OR days_40_to_60 >10000) THEN 0 ELSE days_40_to_60 END AS days_40_to_60,
  CASE WHEN (days_60_to_75 <0 OR days_60_to_75 >10000) THEN 0 ELSE days_60_to_75 END AS days_60_to_75,
  CASE WHEN deal_stage='5% - IQM Held'  THEN 1
       WHEN deal_stage='10% - Discovery' THEN 2
       WHEN deal_stage='20% - Solution'  THEN 3
       WHEN deal_stage='30% - Proof'     THEN 4
       WHEN deal_stage='40% - Proposal'  THEN 5
       WHEN deal_stage='60% - Price Negotiation' THEN 6
       WHEN deal_stage='75% - Contract Review'   THEN 7
       WHEN deal_stage='90% - Deal Desk Review'  THEN 8
       WHEN deal_stage='Closed Won'        THEN 9
       WHEN deal_stage IN ('Deal on Hold','Deal on Hold (Sales Pipeline)') THEN 10
       WHEN deal_stage='Prospect Disengaged' THEN 11
       WHEN deal_stage='Closed Lost'          THEN 12
       WHEN deal_stage='Didn''t Qualify'      THEN 13
       ELSE 0 END AS deal_stage_rank,
  deal_stage, days_in_current_stage,
  CASE
    WHEN deal_stage='10% - Discovery' AND days_in_current_stage<42  THEN 'Green'
    WHEN deal_stage='10% - Discovery' AND days_in_current_stage<56  THEN 'Yellow'
    WHEN deal_stage='10% - Discovery' AND days_in_current_stage>=56 THEN 'Red'
    WHEN deal_stage='20% - Solution'  AND days_in_current_stage<61  THEN 'Green'
    WHEN deal_stage='20% - Solution'  AND days_in_current_stage<82  THEN 'Yellow'
    WHEN deal_stage='20% - Solution'  AND days_in_current_stage>=82 THEN 'Red'
    WHEN deal_stage='30% - Proof'     AND days_in_current_stage<22  THEN 'Green'
    WHEN deal_stage='30% - Proof'     AND days_in_current_stage<30  THEN 'Yellow'
    WHEN deal_stage='30% - Proof'     AND days_in_current_stage>=30 THEN 'Red'
    WHEN deal_stage='40% - Proposal'  AND days_in_current_stage<43  THEN 'Green'
    WHEN deal_stage='40% - Proposal'  AND days_in_current_stage<58  THEN 'Yellow'
    WHEN deal_stage='40% - Proposal'  AND days_in_current_stage>=58 THEN 'Red'
    WHEN deal_stage='60% - Price Negotiation' AND days_in_current_stage<40  THEN 'Green'
    WHEN deal_stage='60% - Price Negotiation' AND days_in_current_stage<54  THEN 'Yellow'
    WHEN deal_stage='60% - Price Negotiation' AND days_in_current_stage>=54 THEN 'Red'
    WHEN deal_stage='75% - Contract Review'   AND days_in_current_stage<51  THEN 'Green'
    WHEN deal_stage='75% - Contract Review'   AND days_in_current_stage<68  THEN 'Yellow'
    WHEN deal_stage='75% - Contract Review'   AND days_in_current_stage>=68 THEN 'Red'
    ELSE NULL
  END AS deal_health_colour,
  CASE
    WHEN deal_stage='1% - IQM Scheduled' THEN 7
    WHEN deal_stage='5% - IQM Held'       THEN 21
    WHEN deal_stage='10% - Discovery'     THEN 28
    WHEN deal_stage='20% - Solution'      THEN 41
    WHEN deal_stage='30% - Proof'         THEN 15
    WHEN deal_stage='40% - Proposal'      THEN 29
    WHEN deal_stage='60% - Price Negotiation' THEN 27
    WHEN deal_stage='75% - Contract Review'   THEN 34
    ELSE NULL
  END AS avg_days_in_stage_benchmark,
  BANT, last_contacted,
  DATE_DIFF('Day', last_contacted, CURRENT_DATE()) AS days_since_last_contact,
  Team, is_this_a_deal_with_inception,
  forecast_amount, forecast_probability, management_forecast, ae_forecast
FROM deal_detail
-- Caller adds: WHERE close_fy = 2027 AND stage_category = 'Active Pipeline' etc.

═══════════════════════════════════════════════════════════════
§8  PATTERN C — ACTUALS vs TARGET / ATTAINMENT
═══════════════════════════════════════════════════════════════

── SOURCE MAPPING FOR PATTERN C (DIFFERENT FROM A/B) ─────────
Executive Outreach AND Investor are MERGED in Pattern C to match
the target table bucket 'Executive Outreach':

  CASE
    WHEN deal_source_rollup IN ('Executive Outreach','Investor')
         THEN 'Executive Outreach'
    WHEN deal_source_rollup IN ('Marketing','Customer Success',
         'AE Outbound','Inception','Hyperscaler') THEN deal_source_rollup
    WHEN deal_source_rollup = 'BDR Outbound' THEN 'BDR'
    WHEN deal_source_rollup = 'Partner'       THEN 'Partner - Non Hyperscaler'
    ELSE 'Other'
  END AS deal_source_rollup

Target table source column mapping:
  source = 'Hyperscalers'                     → 'Hyperscaler'
  source IN ('Executive Outreach','Investor')  → 'Executive Outreach'
  source = 'Partner - Excluding Hyperscalers' → 'Partner - Non Hyperscaler'
  everything else                             → as-is

── COMPLETE SQL TEMPLATE FOR PATTERN C (10% example) ─────────

WITH actuals AS (
  SELECT
    toYear(became_10_deal_date) + if(toMonth(became_10_deal_date)>=4,1,0) AS create_fy,
    CASE WHEN toMonth(became_10_deal_date) IN (1,2,3)    THEN 'Q4'
         WHEN toMonth(became_10_deal_date) IN (4,5,6)    THEN 'Q1'
         WHEN toMonth(became_10_deal_date) IN (7,8,9)    THEN 'Q2'
         WHEN toMonth(became_10_deal_date) IN (10,11,12) THEN 'Q3'
    END AS create_quarter,
    LEFT(formatDateTime(became_10_deal_date,'%M'),3) AS create_month,
    CASE WHEN region='japac' THEN 'JAPAC' WHEN region='Africa' THEN 'Middle East'
         WHEN region='india___sea' THEN 'ISEA' ELSE region END AS region,
    CASE
      WHEN deal_source_rollup IN ('Executive Outreach','Investor') THEN 'Executive Outreach'
      WHEN deal_source_rollup IN ('Marketing','Customer Success',
           'AE Outbound','Inception','Hyperscaler') THEN deal_source_rollup
      WHEN deal_source_rollup = 'BDR Outbound' THEN 'BDR'
      WHEN deal_source_rollup = 'Partner'       THEN 'Partner - Non Hyperscaler'
      ELSE 'Other'
    END AS deal_source_rollup,
    SUM(amount)            AS actual_amount,
    countDistinct(deal_id) AS actual_deals
  FROM hs_analytics.deals FINAL
  WHERE pipeline = 'default'
    AND deal_stage <> 'Duplicate Record'
    AND CASE WHEN deal_type IS NULL THEN 'Not Assigned' ELSE deal_type END
        NOT IN ('Partner-Led SMB')
    AND toInt64(deal_id) IN (
      SELECT DISTINCT toInt64(deal_id_hs) FROM kore_ai_hubspot.gs_deal_ids_hs
    )
    AND became_10_deal_date >= '2025-04-01'
    AND (became_10_deal_date != '1900-01-01'
         OR became_20_deal_date != '1900-01-01'
         OR became_30_deal_date != '1900-01-01'
         OR became_40_deal_date != '1900-01-01'
         OR became_60_deal_date != '1900-01-01'
         OR became_75_deal_date != '1900-01-01'
         OR deal_stage IN ('10% - Discovery','20% - Solution','30% - Proof',
                           '40% - Proposal','60% - Price Negotiation',
                           '75% - Contract Review','90% - Deal Desk Review','Closed Won'))
  GROUP BY 1,2,3,4,5
),
targets AS (
  SELECT
    CAST(fy AS INT) AS create_fy,
    quarter         AS create_quarter,
    month           AS create_month,
    region,
    CASE WHEN source = 'Hyperscalers'                        THEN 'Hyperscaler'
         WHEN source IN ('Executive Outreach','Investor')    THEN 'Executive Outreach'
         WHEN source = 'Partner - Excluding Hyperscalers'    THEN 'Partner - Non Hyperscaler'
         ELSE source END AS deal_source_rollup,
    SUM(toFloat64OrZero(amount_target_10))           AS amount_target,
    SUM(toFloat32OrZero(deals_target_10))            AS deals_target,
    SUM(toFloat64OrZero(amount_target_10_l1))        AS amount_target_l1,
    SUM(toFloat32OrZero(deals_target_10_l1))         AS deals_target_l1,
    SUM(toFloat64OrZero(amount_target_10_committed)) AS amount_target_committed,
    SUM(toFloat32OrZero(deals_target_10_committed))  AS deals_target_committed
  FROM kore_ai_hubspot.gs_pipeline_quotas_v1
  GROUP BY 1,2,3,4,5
)
SELECT
  t.create_fy, t.create_quarter, t.create_month,
  t.region, t.deal_source_rollup,
  coalesce(a.actual_amount, 0) AS actual_amount,
  coalesce(a.actual_deals,  0) AS actual_deals,
  t.amount_target, t.deals_target,
  round(coalesce(a.actual_deals, 0)  / nullIf(t.deals_target,  0) * 100, 1) AS deals_attainment_pct,
  round(coalesce(a.actual_amount, 0) / nullIf(t.amount_target, 0) * 100, 1) AS amount_attainment_pct
FROM targets t
LEFT JOIN actuals a
  ON  t.create_fy          = a.create_fy
  AND t.create_quarter     = a.create_quarter
  AND t.create_month       = a.create_month
  AND t.deal_source_rollup = a.deal_source_rollup
  AND t.region             = a.region
-- WHERE t.create_fy = 2027 AND t.create_quarter = 'Q1'

NOTE — FOR 20% ATTAINMENT:
  Replace became_10_deal_date → became_20_deal_date in the actuals CTE.
  Replace amount_target_10 → amount_target_20, deals_target_10 → deals_target_20.

NOTE — FOR PARTNER ATTAINMENT:
  Use kore_ai_hubspot.gs_partner_targets_region_wise with l2_ prefix columns.
  Filter by partner_team_type IN ('Hyperscaler','GSI/SI','Reseller/BPO/TSD') as needed.
  committed_amount_target_10 and _5 do NOT exist in this table.

NOTE — FOR BDR ATTAINMENT:
  Filter actuals WHERE deal_source_rollup = 'BDR' (after source mapping).
  Use the same Pattern C template with that additional source filter.

═══════════════════════════════════════════════════════════════
§9  FISCAL YEAR AND DATE RULES
═══════════════════════════════════════════════════════════════

Kore.ai fiscal year starts April 1:
  FY27 = Apr 2026 – Mar 2027  (current default)
  FY26 = Apr 2025 – Mar 2026
  FY25 = Apr 2024 – Mar 2025

  FY formula: toYear(date) + if(toMonth(date) >= 4, 1, 0)
  Q1 = Apr/May/Jun | Q2 = Jul/Aug/Sep | Q3 = Oct/Nov/Dec | Q4 = Jan/Feb/Mar

Target table stores: fy as 'FY27', quarter as 'Q1', month as 'Apr' etc.
When joining actuals to targets:
  CAST(fy AS INT) must equal the numeric FY (2027, not 'FY27').

Date column standard — ALWAYS:
  CAST(LEFT(coalesce(column_name, '1900-01-01'), 10) AS DATE)

Sentinel '1900-01-01' = date was never set (NULL equivalent).
Check: column != '1900-01-01'

═══════════════════════════════════════════════════════════════
§10  DIMENSION MAPPINGS (consistent across all patterns)
═══════════════════════════════════════════════════════════════

REGION (apply in SELECT, not WHERE):
  'japac'       → 'JAPAC'
  'Africa'      → 'Middle East'
  'india___sea' → 'ISEA'

INDUSTRY:
  ('Financial Services','Banking','Insurance')             → 'Financial Services'
  ('Manufacturing Discreet','Manufacturing Process','CPG') → 'Manufacturing'
  ('Hi-Tech','Telecom / Media / Entertainment')           → 'TMT'
  ('Business Services','Government','Energy & Utilities',
   'Education','Restaurants','null','Energy') or NULL     → 'Other'

ACCOUNT PRIORITY:
  ('P1','P2','P3','P4') → 'P1-P4'
  ('P5','P6','P7')      → 'P5-P7'
  ('P8','P9','P10')     → 'P8-P10'
  NULL → NULL (Pattern A) or 'No Priority' (Pattern B/C detail)

INCEPTION DEALS:
  Only exclude if user explicitly asks. Default: include all deals.

═══════════════════════════════════════════════════════════════
§11  MQL RULES
═══════════════════════════════════════════════════════════════
MQL actuals from hs_analytics.contacts FINAL require ALL THREE:
  1. lifecycle_stage = 'marketingqualifiedlead'
     AND date_entered_marketing_qualified_lead_lifecycle_stage_pipeline IS NOT NULL
  2. company_priority IN ('P1','P2','P3','P4','P5','P6','P7')
  3. lead_status != 'Bad Data'

NEVER omit any of the 3 MQL filters — missing any one inflates counts.

MQL ACTUALS PATTERN:
  SELECT
    region,
    original_source,
    toYYYYMM(toDate(LEFT(date_entered_marketing_qualified_lead_lifecycle_stage_pipeline,10))) AS ym,
    countDistinct(contact_id) AS mql_count
  FROM hs_analytics.contacts FINAL
  WHERE lifecycle_stage = 'marketingqualifiedlead'
    AND date_entered_marketing_qualified_lead_lifecycle_stage_pipeline IS NOT NULL
    AND company_priority IN ('P1','P2','P3','P4','P5','P6','P7')
    AND lead_status != 'Bad Data'
    AND <date range filter on date_entered_... column>
  GROUP BY region, original_source, ym

MQL TARGET PATTERN — always filter by exact quarter, NEVER divide annual by 4:
  SELECT
    region,
    original_source,
    SUM(toFloat64OrZero(mql_target)) AS mql_tgt
  FROM kore_ai_hubspot.gs_marketing_targets
  WHERE fy = 'FY27' AND quarter = 'Q1'
  GROUP BY region, original_source

FILTER CONFIRMATION for MQL queries — include in Filters Applied block:
  - Company Priority: P1–P7 (excludes unranked contacts)
  - Lead Status: excludes 'Bad Data'
  - MQL date range: <the date range used>

═══════════════════════════════════════════════════════════════
§12  DASHBOARD DEFINITIONS
═══════════════════════════════════════════════════════════════
When a user asks about a specific dashboard, apply the correct logic below.
If unclear which dashboard, ask the user.

── DASHBOARD 1: EOP (End-of-Period) ───────────────────────────
PURPOSE: Tracks pipeline health and attainment against EOP targets.
KEY METRICS: EOP Pipeline Value, EOP Target, EOP Attainment %, stage-wise and region-wise EOP.
FILTERS: MANDATORY_BASE_FILTERS + close_date within current quarter end window
         + deal_stage IN active stages (20%–75%)

── DASHBOARD 2: EXEC KPI ──────────────────────────────────────
PURPOSE: Senior leadership view of pipeline performance.
KEY METRICS: Total Active Pipeline ($M), Closed Won ($M), CW Attainment %,
             Win Rate %, Pipeline Coverage, New Logo Count.

── DASHBOARD 3: CS (Customer Success) ─────────────────────────
PURPOSE: Tracks renewals, upsells, expansions, CS team performance.

── DASHBOARD 4: GLOBAL PIPELINE GOVERNANCE ────────────────────
PURPOSE: Executive governance view across all regions, sources, partner types.

── DASHBOARD 5: GLOBAL PIPEGEN ────────────────────────────────
PURPOSE: Org-wide pipeline generation performance and attainment.
KEY METRICS: 5%, 10%, 20% Pipeline Amount and Deal Count, Actual vs Target,
             Attainment %, Pipeline Trend, Funnel Conversion, Region/Source/Industry.
FILTERS: MANDATORY_BASE_FILTERS + stage/date filters + gs_pipeline_quotas_v1 targets.

── DASHBOARD 6: PARTNERSHIP ───────────────────────────────────
PURPOSE: Partner-generated pipeline, target attainment, partner ecosystem performance.
KEY METRICS: Partner 5%/10%/20% Pipeline, Partner Attainment %, PSD Performance,
             Hyperscaler Performance, GSI/SI, Reseller/BPO/TSD, Partner Funnel Conversion.
FILTERS: MANDATORY_BASE_FILTERS + Partner/Hyperscaler sources + partner target tables.

── DASHBOARD 7: MARKETING ─────────────────────────────────────
PURPOSE: MQL actuals vs target, marketing pipeline, marketing attainment.
KEY METRICS: MQL Actual vs Target, Marketing-Sourced Pipeline, Deal Count,
             Source-wise Performance, Region-wise Performance, MQL Conversion Rate.
FILTERS: Mandatory MQL filters (§11) + gs_marketing_targets + selected fiscal period.

── DASHBOARD 8: AE FOCUS ──────────────────────────────────────
PURPOSE: AE pipeline, quota attainment, and sales performance.
KEY METRICS: Active Pipeline, CW ARR, CW Deal Count, Actual vs Quota,
             Attainment %, Win Rate, Pipeline Coverage, Avg Deal Size, Sales Cycle.
FILTERS: MANDATORY_BASE_FILTERS + gs_closed_won_quotas + AE filters.

── DASHBOARD 9: BDR FOCUS ─────────────────────────────────────
PURPOSE: BDR pipeline generation and conversion performance.
KEY METRICS: Meetings Created, Opportunities Generated, 5%/10%/20% PipeGen,
             Stage Conversion, BDR Performance, Region-wise, Target Attainment.
FILTERS: MANDATORY_BASE_FILTERS + BDR ownership filters + selected fiscal period.

═══════════════════════════════════════════════════════════════
§13  SQL GENERATION GUARDRAILS
═══════════════════════════════════════════════════════════════
1.  SELECT / WITH only — no INSERT, UPDATE, DELETE, DROP, ALTER.
2.  FINAL on every hs_analytics.* table.
3.  All MANDATORY BASE FILTERS (§3) on every deals query.
4.  countDistinct(deal_id) — never count() or count(deal_id).
5.  No LIMIT unless user says "top N" or "first N".
6.  All target table numeric columns: SUM(toFloat64OrZero(col)).
7.  Division: always wrap denominator with nullIf(denom, 0).
8.  Date columns: CAST(LEFT(coalesce(col,'1900-01-01'),10) AS DATE).
9.  NEVER compute a quarterly target by dividing annual by 4.
10. State total row count in every answer.

⚠️  IMPORTANT FOR RULES.PY VALIDATOR:
Pattern A uses became_N_deal_date in OR conditions (cumulative counting),
NOT as a cohort anchor with stage exclusions. The rules.py cohort checks
were designed for true cohort funnel queries, not Pattern A.
For Pattern A SQL, add a comment at the top of the query:
  -- Pattern A: cumulative stage counting, not cohort funnel

═══════════════════════════════════════════════════════════════
§14  OUTPUT FORMATTING
═══════════════════════════════════════════════════════════════
- Clean markdown. Tables for data. Bold key numbers.
- Never fabricate numbers. Never run destructive SQL.
- Format dollar amounts: round(sum(amount)/1e6, 1) as $M.
- Always complete full response — never truncate mid-table.

After every DB-backed answer, append:

---
**Filters Applied:**
- Pattern used: [A — Cumulative Funnel / B — Deal Detail / C — Attainment]
- FY anchor column: [which became_N_deal_date or close_date was used]
- [list all active filters: FY, quarter, region, source, stage, etc.]

Please verify these filters match your expectation.
---

═══════════════════════════════════════════════════════════════
§15  VISUAL / CHART GENERATION
═══════════════════════════════════════════════════════════════
Charts are generated automatically by the system after your response.
Do NOT write Chart.js, SVG, or any ```html chart code yourself.

═══════════════════════════════════════════════════════════════
LIVE DATABASE SCHEMA (auto-injected below)
═══════════════════════════════════════════════════════════════
{schema}
"""

_SYSTEM_PROMPT = _build_system_prompt()

# =============================================================================
# ClickHouse query runner
# =============================================================================
def run_clickhouse_query(sql: str, session_id: Optional[str] = None) -> str:
    base_url = _base_url()
    token    = _token()

    if not base_url:
        return "DATABASE CONNECTION FAILED: CLICKHOUSE_API_URL is not set."
    if not token:
        return "DATABASE CONNECTION FAILED: CLICKHOUSE_API_TOKEN is not set."
        
    cleaned = re.sub(r'^\s*--[^\n]*\n', '', sql.strip(), flags=re.MULTILINE).strip()
    stripped = cleaned.upper()
    if not (stripped.startswith("SELECT") or stripped.startswith("WITH")):
        return "ERROR: Only SELECT/WITH queries are permitted."
    for kw in FORBIDDEN_KEYWORDS:
        if re.search(rf'\b{kw}\b', stripped):
            return f"ERROR: Forbidden keyword: {kw}"
    

    print(f"🔍 SQL (session={session_id}) → {sql[:200]}")

    try:
        resp = httpx.post(
            f"{base_url}/query",
            headers=_auth_headers(),
            json={"query": sql},
            timeout=60,
        )

        if resp.status_code == 401:
            return "DATABASE CONNECTION FAILED: 401 Unauthorized."
        if resp.status_code == 403:
            return "DATABASE CONNECTION FAILED: 403 Forbidden."
        if resp.status_code == 422:
            return f"ERROR: Proxy rejected query (422): {resp.text[:400]}"
        if resp.status_code == 500:
            return f"DATABASE ERROR: HTTP 500: {resp.text[:400]}"
        if resp.status_code != 200:
            return f"DATABASE ERROR: HTTP {resp.status_code} — {resp.text[:300]}"

        payload = resp.json()

        if isinstance(payload, list):
            rows = payload
        elif isinstance(payload, dict):
            rows = payload.get("data") or payload.get("rows") or payload.get("result") or payload.get("results")
            api_columns = payload.get("columns") or payload.get("meta") or payload.get("column_names")
            if rows is None:
                return json.dumps(payload, indent=2, default=str)[:3000]
        else:
            return f"Unexpected response type: {type(payload)}"

        if not rows:
            return "Query returned 0 rows."

        if isinstance(rows[0], dict):
            columns = list(rows[0].keys())
            norm_rows = rows
        else:
            if api_columns and len(api_columns) == len(rows[0]):
                columns = [c["name"] if isinstance(c, dict) else c for c in api_columns]
            else:
                columns = [f"col_{i}" for i in range(len(rows[0]))]
            norm_rows = [dict(zip(columns, r)) for r in rows]

        total_rows = len(norm_rows)

        if session_id:
            existing = _SESSION_STORE.get(session_id)
            # Always store the most recent query result for chart/export purposes.
            # Use a SEPARATE key (chat_session_id + ":latest") for the most recent
            # result specifically, while preserving the old "largest result wins"
            # behavior under the original session_id for export/CSV purposes.
            _store_result(f"{session_id}:latest", QueryResult(
                sql             = sql,
                columns         = columns,
                rows            = norm_rows,
                total_rows      = total_rows,
                captured_at     = datetime.utcnow().isoformat() + "Z",
                filters_applied = _extract_filters_from_sql(sql),
            ))
            # Only overwrite the EXPORT-facing session result if this is larger.
            # This prevents a trivial follow-up query from replacing a large
            # deal-list result meant for CSV/PDF export with a tiny one.
            if not existing or total_rows > existing.total_rows:
                _store_result(session_id, QueryResult(
                    sql             = sql,
                    columns         = columns,
                    rows            = norm_rows,
                    total_rows      = total_rows,
                    captured_at     = datetime.utcnow().isoformat() + "Z",
                    filters_applied = _extract_filters_from_sql(sql),
                ))
            else:
                print(f"   ⏭️  Skipping EXPORT session store update — {total_rows} row(s) won't overwrite existing {existing.total_rows} rows.")

        CHAT_DISPLAY_LIMIT = 100
        header = " | ".join(columns)
        lines  = [header, "-" * min(len(header), 140)]
        for row in norm_rows[:CHAT_DISPLAY_LIMIT]:
            lines.append(" | ".join(str(row.get(c, "")) for c in columns))

        if total_rows > CHAT_DISPLAY_LIMIT:
            lines.append(
                f"\n📊 **Showing {CHAT_DISPLAY_LIMIT} of {total_rows} rows.** "
                f"Say **\"export these deals as CSV\"** or **\"export as PDF\"** "
                f"to download all {total_rows} records."
            )

        result = "\n".join(lines)
        print(f"   ✅ {total_rows} rows returned. Session store updated.")
        return result

    except httpx.ConnectError as e:
        return f"DATABASE CONNECTION FAILED: Could not reach {base_url}. {e}"
    except httpx.TimeoutException:
        return "DATABASE CONNECTION FAILED: Query timed out after 60 seconds."
    except Exception as exc:
        traceback.print_exc()
        return f"DATABASE CONNECTION FAILED: {type(exc).__name__}: {exc}"


def _extract_filters_from_sql(sql: str) -> str:
    sql_upper = sql.upper()
    filters = []
    if "PIPELINE = 'DEFAULT'" in sql_upper:
        filters.append("Pipeline: default")
    
    # Dynamically detect any cohort stage (5, 10, 20, 30, 40, 60, 75)
    cohort_stages = sorted(set(re.findall(r'BECAME_(\d+)_DEAL_DATE', sql_upper)), key=int)
    for stage in cohort_stages:
        filters.append(f"Cohort: {stage}% qualified deals")
    
    if "CLOSE_DATE" in sql_upper and "2026-04-01" in sql:
        filters.append("FY27 active pipeline")
    if "DEAL_STAGE" in sql_upper:
        m = re.search(r"deal_stage\s+IN\s*\(([^)]+)\)", sql, re.IGNORECASE)
        if m:
            filters.append(f"Stage filter: {m.group(1)[:60]}")
    if "REGION" in sql_upper:
        m = re.search(r"region\s*=\s*'([^']+)'", sql, re.IGNORECASE)
        if m:
            filters.append(f"Region: {m.group(1)}")
    if "COMPANY_PRIORITY" in sql_upper:
        filters.append("Company Priority: P1–P7")
    if "LEAD_STATUS" in sql_upper and "BAD DATA" in sql_upper:
        filters.append("Lead Status: excludes 'Bad Data'")
    if "LIFECYCLE_STAGE" in sql_upper and "MARKETINGQUALIFIEDLEAD" in sql_upper:
        filters.append("Lifecycle: MQL only")
    return "; ".join(filters) if filters else "Standard base filters applied"

def _is_numericish(v) -> bool:
    try:
        float(v)
        return True
    except (TypeError, ValueError):
        return False


def _to_float(v) -> float:
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0.0


def _compute_deterministic_headline(stored: Optional["QueryResult"]) -> Optional[str]:
    """For small aggregate results, compute totals/averages in Python so Claude
    never has to do the arithmetic itself when writing the summary."""
    if not stored or not stored.rows or len(stored.rows) > 20:
        return None

    numeric_cols = [
        c for c in stored.columns
        if any(_is_numericish(r.get(c)) for r in stored.rows)
    ]
    if not numeric_cols:
        return None

    parts = []
    for col in numeric_cols:
        total = sum(_to_float(r.get(col)) for r in stored.rows)
        parts.append(f"{col}_total={total:,.2f}")
        if len(stored.rows) > 1:
            avg = total / len(stored.rows)
            parts.append(f"{col}_avg={avg:,.2f}")
    return " | ".join(parts)


_METRIC_FNS: Dict[str, Callable[[List[dict]], float]] = {
    "win_rate": lambda rows: (
        sum(1 for r in rows if r.get("deal_stage") == "Closed Won")
        / max(sum(1 for r in rows if r.get("deal_stage") in ("Closed Won", "Closed Lost")), 1)
        * 100
    ),
    "active_pipeline_total": lambda rows: sum(
        _to_float(r.get("amount")) for r in rows
        if r.get("deal_stage") in (
            "20% - Solution", "30% - Proof", "40% - Proposal",
            "60% - Price Negotiation", "75% - Contract Review",
        )
    ) / 1_000_000,
    "avg_deal_size": lambda rows: (
        sum(_to_float(r.get("amount")) for r in rows) / max(len(rows), 1)
    ),
    "closed_won_total": lambda rows: sum(
        _to_float(r.get("amount")) for r in rows if r.get("deal_stage") == "Closed Won"
    ) / 1_000_000,
    "mql_count": lambda rows: float(len(rows)),
    "attainment_pct": lambda rows: (
        sum(_to_float(r.get("achieved_m")) for r in rows)
        / max(sum(_to_float(r.get("target_m")) for r in rows), 0.0001) * 100
    ),
}


def compute_verified_metric(metric_name: str, session_id: Optional[str]) -> str:
    stored = _get_result(session_id) if session_id else None
    if not stored or not stored.rows:
        return "ERROR: No stored query result for this session — run a query first."
    fn = _METRIC_FNS.get(metric_name)
    if not fn:
        return f"ERROR: Unknown metric '{metric_name}'."
    try:
        value = fn(stored.rows)
        return f"VERIFIED {metric_name} = {value:,.2f}"
    except Exception as e:
        return f"ERROR computing {metric_name}: {e}"
        
        
        
def _parse_rows_for_validation(query_result: str, session_id: Optional[str]) -> List[dict]:
    """
    run_clickhouse_query() returns a human-readable pipe-delimited text table,
    not JSON, so we can't parse query_result directly. Instead, pull the
    structured rows it already stored in _SESSION_STORE for this session.
    """
    if not session_id:
        return []
    stored = _SESSION_STORE.get(session_id)
    if not stored:
        return []
    return stored.rows


# =============================================================================
# Startup
# =============================================================================
@app.on_event("startup")
async def on_startup():
    global _SYSTEM_PROMPT
    try:
        discover_schema()
        _SYSTEM_PROMPT = _build_system_prompt()
    except Exception as e:
        print(f"⚠️  Schema discovery failed: {e}")
    _cleanup_sessions()
    print("🚀 DIUD v4 started — session-store export enabled.")


# =============================================================================
# Claude tool definition
# =============================================================================
_QUERY_TOOL = {
    "name": "query_clickhouse",
    "description": (
        "Execute a SELECT query against ClickHouse. "
        "Use for deal pipeline, AE performance, win/loss, regions, stages, MQL metrics. "
        "ALWAYS use fully-qualified table names. "
        "Relay DATABASE CONNECTION FAILED errors directly to the user."
    ),
    "input_schema": {
        "type": "object",
        "properties": {
            "sql": {
                "type": "string",
                "description": (
                    "Valid ClickHouse SELECT or WITH query. "
                    "For deal LIST queries: NO LIMIT unless user asks for 'top N'. "
                    "Return all matching rows — the system displays them safely."
                ),
            }
        },
        "required": ["sql"],
    },
}

_RULEBOOK_TOOL = {
    "name": "lookup_business_rule",
    "description": (
        "Look up the exact business rule for a specific metric/topic BEFORE "
        "writing SQL for it. ALWAYS call this before generating SQL for MQL, "
        "active pipeline, cohort funnels, attainment/targets, closed won, "
        "partner targets, or dashboard-specific questions."
    ),
    "input_schema": {
        "type": "object",
        "properties": {
            "topic": {
                "type": "string",
                "enum": [
                    "mql", "active_pipeline", "cohort_funnel", "attainment",
                    "closed_won", "partner_targets", "dashboard_definitions",
                ],
            }
        },
        "required": ["topic"],
    },
}

_COMPUTE_METRIC_TOOL = {
    "name": "compute_verified_metric",
    "description": (
        "Compute a verified metric (win_rate, active_pipeline_total, "
        "avg_deal_size, closed_won_total, mql_count, attainment_pct) from the "
        "current session's query results. ALWAYS use this instead of computing "
        "percentages or sums yourself from the raw row dump."
    ),
    "input_schema": {
        "type": "object",
        "properties": {
            "metric_name": {
                "type": "string",
                "enum": [
                    "win_rate", "active_pipeline_total", "avg_deal_size",
                    "closed_won_total", "mql_count", "attainment_pct",
                ],
            },
        },
        "required": ["metric_name"],
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
    session_id: Optional[str] = None
    model: str = "sonnet"

class ExportPreviewRequest(BaseModel):
    conversation: List[ChatMessage] = []
    title: str = "Pipeline Intelligence Report"
    export_type: Literal["pdf", "pptx"] = "pdf"
    detail_level: Literal["summary", "detailed"] = "detailed"
    session_id: Optional[str] = None

class ExportDownloadRequest(BaseModel):
    format: Literal["pdf", "pptx", "csv"]
    content: Optional[str] = None
    title: str = "Pipeline Intelligence Report"
    session_id: Optional[str] = None


# =============================================================================
# Claude tool loop
# FIX 1: max_tokens raised to CHAT_MAX_TOKENS (4096) for chat
# =============================================================================
def _extract_text(content_blocks) -> str:
    return "\n".join(
        b.text for b in content_blocks if hasattr(b, "text") and b.text
    ).strip()


def _call_claude(messages: list, max_tokens: int = CHAT_MAX_TOKENS,
                 session_id: Optional[str] = None, model: str = "sonnet") -> str:

    selected_model = ALLOWED_MODELS.get(model, ALLOWED_MODELS["sonnet"])

    # FIX: original_user_message was referenced below but never defined —
    # this caused "NameError: name 'original_user_message' is not defined"
    # on every request that triggered the tool loop.
    original_user_message = next(
        (m.get("content", "") for m in reversed(messages)
         if m.get("role") == "user" and isinstance(m.get("content", ""), str)),
        ""
    )

    safe_messages = []
    for m in messages:
        content = m.get("content", "")
        if isinstance(content, list):
            text_parts = [
                b.get("text", "") if isinstance(b, dict) else (b.text if hasattr(b, "text") else "")
                for b in content if (isinstance(b, dict) and b.get("type") == "text")
                   or (hasattr(b, "type") and b.type == "text")
            ]
            text = "\n".join(t for t in text_parts if t).strip()
            if text:
                safe_messages.append({"role": m["role"], "content": text})
        else:
            safe_messages.append({"role": m["role"], "content": content})

    response = _ai_client.messages.create(
        model=selected_model,
        system=_SYSTEM_PROMPT,
        messages=safe_messages,
        tools=[_QUERY_TOOL, _RULEBOOK_TOOL, _COMPUTE_METRIC_TOOL],
        temperature=0,
        max_tokens=max_tokens,
    )

    MAX_ROUNDS = 8
    last_error = None

    for round_num in range(MAX_ROUNDS):
        if response.stop_reason != "tool_use":
            break

        tool_blocks = [b for b in response.content if b.type == "tool_use"]
        if not tool_blocks:
            break

        tool_result_blocks = []
        for tool_block in tool_blocks:

            # --- rulebook lookup branch ----------------------------
            if tool_block.name == "lookup_business_rule":
                topic = tool_block.input.get("topic", "")
                tool_result_blocks.append({
                    "type":        "tool_result",
                    "tool_use_id": tool_block.id,
                    "content":     get_rulebook_entry(topic),
                    "is_error":    False,
                })
                continue

            # --- deterministic metric branch ------------------------
            if tool_block.name == "compute_verified_metric":
                metric_name = tool_block.input.get("metric_name", "")
                tool_result_blocks.append({
                    "type":        "tool_result",
                    "tool_use_id": tool_block.id,
                    "content":     compute_verified_metric(metric_name, session_id),
                    "is_error":    False,
                })
                continue

            sql = tool_block.input.get("sql", "")

            # --- pre-execution rule gate -------------------------------
            sql_violations = validate_sql_against_rules(sql, original_user_message)
            _log_rule_audit(session_id, sql, sql_violations, "pre_execute", original_user_message)

            if sql_violations:
                cohort_violations = [v for v in sql_violations if "cohort" in v.lower() or "8b" in v.lower()]

                if cohort_violations:
                    query_result = (
                        "RULE VIOLATION — cohort funnel SQL rejected (§8b). NOT executed.\n\n"
                        "Violations:\n"
                        + "\n".join(f"- {v}" for v in sql_violations)
                        + """

            ⚠️ REQUIRED: Rewrite from scratch using this exact skeleton:

            WITH cohort AS (
              SELECT deal_id, deal_stage, amount
              FROM hs_analytics.deals FINAL
              WHERE became_<N>_deal_date IS NOT NULL
                AND deal_stage NOT IN (
                    '1% - Prospect', '<stages before N%>'
                )
                AND pipeline = 'default'
                AND CASE WHEN deal_type IS NULL THEN 'Not Assigned' ELSE deal_type END
                    NOT IN ('Partner-Led SMB')
                AND toInt64(deal_id) IN (
                    SELECT DISTINCT toInt64(deal_id_hs) FROM kore_ai_hubspot.gs_deal_ids_hs
                )
                AND toDate(LEFT(coalesce(became_<N>_deal_date, '1900-01-01'), 10))
                    BETWEEN '<start_date>' AND '<end_date>'
            )
            SELECT
                deal_stage,
                countDistinct(deal_id)       AS deal_count,
                round(SUM(amount) / 1e6, 1) AS pipeline_m
            FROM cohort
            GROUP BY deal_stage
            ORDER BY deal_count DESC

            Fill in <N> and the stage exclusion list from §8b, then resubmit.
            """
                    )
                else:
                    query_result = (
                        "RULE VIOLATION — query rejected, NOT executed against the database:\n"
                        + "\n".join(f"- {v}" for v in sql_violations)
                        + "\n\nRewrite the SQL to satisfy these rules, then call the tool again."
                    )
                is_error = True

            else:
                query_result = run_clickhouse_query(sql, session_id=session_id)
                is_error = any(query_result.startswith(p) for p in [
                    "DATABASE CONNECTION FAILED", "ERROR:", "DATABASE ERROR:"
                ])

                # --- post-execution result-level rule check ------------
                if not is_error:
                    parsed_rows = _parse_rows_for_validation(query_result, session_id)
                    result_violations = validate_result_against_rules(
                        parsed_rows, original_user_message, sql
                    )
                    _log_rule_audit(session_id, sql, result_violations, "post_execute", original_user_message)
                    if result_violations:
                        query_result += (
                            "\n\n⚠️ RESULT RULE VIOLATION:\n"
                            + "\n".join(f"- {v}" for v in result_violations)
                            + "\nThis result is likely wrong — re-derive using the single-CTE "
                              "cohort pattern from §8b and call the tool again."
                        )
                        is_error = True
                    else:
                        # --- deterministic headline injection ------
                        headline = _compute_deterministic_headline(_get_result(session_id))
                        if headline:
                            query_result += (
                                f"\n\n[VERIFIED TOTALS — use these exact figures, "
                                f"do not recompute]: {headline}"
                            )

            if is_error:
                last_error = query_result

            tool_result_blocks.append({
                "type":        "tool_result",
                "tool_use_id": tool_block.id,
                "content":     query_result,
                "is_error":    is_error,
            })

        safe_messages = safe_messages + [
            {"role": "assistant", "content": response.content},
            {"role": "user",      "content": tool_result_blocks},
        ]
        is_last_round = (round_num == MAX_ROUNDS - 1)
        response = _ai_client.messages.create(
            model=selected_model,
            system=_SYSTEM_PROMPT,
            messages=safe_messages,
            tools=[] if is_last_round else [_QUERY_TOOL, _RULEBOOK_TOOL, _COMPUTE_METRIC_TOOL],
            temperature=0,
            max_tokens=max_tokens,
        )

    reply = _extract_text(response.content)

    # ── INCOMPLETE-TURN GUARD ────────────────────────────────────────
    # If the final round was cut off by max_tokens (not a natural
    # tool_use/end_turn stop), "reply" may just be Claude narrating its
    # NEXT planned step ("let me check...") without executing it. Force
    # one continuation so the investigation actually finishes instead of
    # serving unfinished reasoning as the answer.
    if response.stop_reason == "max_tokens":
        continue_messages = safe_messages + [
            {"role": "assistant", "content": reply},
            {"role": "user", "content": (
                "Continue. You were cut off mid-response. If you said you "
                "were going to check or fetch something, actually call the "
                "appropriate tool now and complete the investigation. Do "
                "not repeat or re-explain your plan — execute it, then give "
                "the final answer."
            )},
        ]
        try:
            continue_response = _ai_client.messages.create(
                model=selected_model,
                system=_SYSTEM_PROMPT,
                messages=continue_messages,
                tools=[_QUERY_TOOL, _RULEBOOK_TOOL, _COMPUTE_METRIC_TOOL],
                temperature=0,
                max_tokens=max_tokens,
            )
            for _ in range(3):
                if continue_response.stop_reason != "tool_use":
                    break
                tool_blocks = [b for b in continue_response.content if b.type == "tool_use"]
                if not tool_blocks:
                    break
                tool_result_blocks = []
                for tb in tool_blocks:
                    if tb.name == "lookup_business_rule":
                        tool_result_blocks.append({
                            "type": "tool_result", "tool_use_id": tb.id,
                            "content": get_rulebook_entry(tb.input.get("topic", "")),
                            "is_error": False,
                        })
                        continue
                    if tb.name == "compute_verified_metric":
                        tool_result_blocks.append({
                            "type": "tool_result", "tool_use_id": tb.id,
                            "content": compute_verified_metric(tb.input.get("metric_name", ""), session_id),
                            "is_error": False,
                        })
                        continue
                    sql = tb.input.get("sql", "")
                    violations = validate_sql_against_rules(sql, original_user_message)
                    _log_rule_audit(session_id, sql, violations, "pre_execute", original_user_message)
                    if violations:
                        result = "RULE VIOLATION — not executed:\n" + "\n".join(violations)
                        is_err = True
                    else:
                        result = run_clickhouse_query(sql, session_id=session_id)
                        is_err = any(result.startswith(p) for p in
                            ["DATABASE CONNECTION FAILED", "ERROR:", "DATABASE ERROR:"])
                    tool_result_blocks.append({
                        "type": "tool_result", "tool_use_id": tb.id,
                        "content": result, "is_error": is_err,
                    })
                continue_messages = continue_messages + [
                    {"role": "assistant", "content": continue_response.content},
                    {"role": "user", "content": tool_result_blocks},
                ]
                continue_response = _ai_client.messages.create(
                    model=selected_model, system=_SYSTEM_PROMPT,
                    messages=continue_messages,
                    tools=[_QUERY_TOOL, _RULEBOOK_TOOL, _COMPUTE_METRIC_TOOL],
                    temperature=0, max_tokens=max_tokens,
                )
            completed = _extract_text(continue_response.content)
            if completed:
                reply = completed
                safe_messages = continue_messages  # keep fact-binding verifier in sync
        except Exception as e:
            print(f"⚠️ Continuation after max_tokens cutoff failed: {e}")
            # fall through with the original (incomplete) reply rather than erroring out


    # ── FACT-BINDING VERIFIER ───────────────────────────────────────
    if reply and session_id:
        stored = _get_result(f"{session_id}:latest") or _get_result(session_id)
        if stored and stored.rows:
            violations = validate_summary_against_facts(reply, stored.rows)
            if violations:
                _log_rule_audit(session_id, "summary_check", violations, "post_summary", original_user_message)
                print(f"⚠️ Unverified numbers in summary: {violations}")
                retry_messages = safe_messages + [
                    {"role": "assistant", "content": reply},
                    {"role": "user", "content": (
                        "Rewrite your last response as a clean, standalone answer "
                        "using ONLY numbers that literally appear in the returned "
                        "rows, the [VERIFIED TOTALS] block, or a "
                        "compute_verified_metric result. IMPORTANT: This rewrite is "
                        "the ONLY version the user will ever see — they have not "
                        "seen your previous draft. Do NOT reference 'your previous "
                        "response', do NOT apologize, do NOT say 'you're right' — "
                        "just give the answer directly, as if this is the first "
                        "and only time you're answering."
                    )},
                ]
                try:
                    retry_response = _ai_client.messages.create(
                        model=selected_model,
                        system=_SYSTEM_PROMPT,
                        messages=retry_messages,
                        tools=[],
                        temperature=0,
                        max_tokens=max_tokens,
                    )
                    reply = _extract_text(retry_response.content) or reply
                except Exception as e:
                    print(f"⚠️ Fact-binding regeneration failed: {e}")
                    # keep original reply rather than losing the response entirely
# ── DETERMINISTIC CHART INJECTION ────────────────────────────────
    _INVESTIGATION_MARKERS = ("let me check", "let me look", "let me verify",
                               "i'll check", "i'll look", "checking what",
                               "let me fetch", "let me re-run", "let me query")
    reply_looks_unfinished = (
        any(m in reply.lower() for m in _INVESTIGATION_MARKERS)
        and "filters applied" not in reply.lower()
    )
    if reply and session_id and not reply_already_has_chart(reply) and not reply_looks_unfinished:
        stored = _get_result(f"{session_id}:latest") or _get_result(session_id)
        if stored and stored.rows:
            chart_html = build_chart_html(
                stored.columns, stored.rows, stored.filters_applied
            )
            if chart_html:
                reply = reply.rstrip() + "\n\n" + chart_html

    if not reply:
        if last_error:
            reply = (
                "⚠️ I couldn't complete this query. The last database error was:\n\n"
                f"`{last_error[:400]}`\n\n"
                "Could you rephrase the question, or check **/debug/db** if this persists?"
            )
        else:
            reply = (
                "⚠️ I wasn't able to finish answering this in time — it may need "
                "a more specific or simpler question. Could you try rephrasing it "
                "(e.g. break it into smaller asks)?"
            )
    return reply

# =============================================================================
# Routes — chat
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
        return {"status": "MISCONFIGURED", "config": config}

    tests = {}
    try:
        r = httpx.get(base_url, timeout=10)
        tests["GET /"] = {"status": r.status_code}
    except Exception as e:
        tests["GET /"] = {"error": str(e)}

    ping = run_clickhouse_query("SELECT 1 AS ping")
    query_ok = not any(ping.startswith(p) for p in ["DATABASE CONNECTION FAILED", "ERROR:", "DATABASE ERROR:"])
    tests["SELECT 1"] = {"ok": query_ok, "result": ping[:100]}

    return {
        "status":            "OK" if query_ok else "FAILED",
        "config":            config,
        "discovered_tables": list(_LIVE_SCHEMA.keys()),
        "active_sessions":   len(_SESSION_STORE),
        "tests":             tests,
    }


@app.post("/refresh-schema")
def refresh_schema():
    global _SYSTEM_PROMPT
    schema = discover_schema()
    _SYSTEM_PROMPT = _build_system_prompt()
    return {"status": "refreshed", "tables": list(_LIVE_SCHEMA.keys())}


@app.post("/chat")
def chat(payload: ChatRequest):
    messages = [{"role": m.role, "content": m.content} for m in payload.history]
    messages.append({"role": "user", "content": payload.message})
    print(f"💬 [chat] session={payload.session_id} msg={payload.message[:80]}")

    try:
        # FIX 1: pass CHAT_MAX_TOKENS explicitly
        reply = _call_claude(messages, max_tokens=CHAT_MAX_TOKENS,
                             session_id=payload.session_id, model=payload.model)
    except Exception as exc:
        traceback.print_exc()
        raise HTTPException(status_code=502, detail=f"Claude error: {exc}")

    has_dataset = payload.session_id is not None and payload.session_id in _SESSION_STORE
    stored = _SESSION_STORE.get(payload.session_id) if payload.session_id else None

    return {
        "reply":        reply,
        "has_dataset":  has_dataset,
        "dataset_rows": stored.total_rows if stored else 0,
        "export_intent": "__EXPORT_INTENT__" in reply,
    }


# =============================================================================
# Retry
# =============================================================================
class RetryRequest(BaseModel):
    history:    List[ChatMessage] = []
    session_id: Optional[str] = None
    model:      str = "sonnet"


@app.post("/chat/retry")
def chat_retry(payload: RetryRequest):
    if not payload.history:
        raise HTTPException(status_code=400, detail="history must not be empty for retry.")

    clean_history = list(payload.history)
    while clean_history and clean_history[-1].role == "assistant":
        clean_history.pop()

    if not clean_history or clean_history[-1].role != "user":
        raise HTTPException(
            status_code=400,
            detail="No user message found to retry. history must end with a user turn."
        )

    last_user_msg = clean_history[-1].content
    prior_history = clean_history[:-1]

    print(f"🔄 [retry] session={payload.session_id} retrying: {last_user_msg[:80]}")

    messages = [{"role": m.role, "content": m.content} for m in prior_history]
    messages.append({"role": "user", "content": last_user_msg})

    try:
        reply = _call_claude(messages, max_tokens=CHAT_MAX_TOKENS, session_id=payload.session_id)
    except Exception as exc:
        traceback.print_exc()
        raise HTTPException(status_code=502, detail=f"Claude error on retry: {exc}")

    has_dataset = payload.session_id is not None and payload.session_id in _SESSION_STORE
    stored      = _SESSION_STORE.get(payload.session_id) if payload.session_id else None

    return {
        "reply":         reply,
        "has_dataset":   has_dataset,
        "dataset_rows":  stored.total_rows if stored else 0,
        "export_intent": "__EXPORT_INTENT__" in reply,
        "retried":       True,
    }


# =============================================================================
# Session info
# =============================================================================
@app.get("/session/{session_id}/dataset-info")
def session_dataset_info(session_id: str):
    result = _get_result(session_id)
    if not result:
        return {"has_dataset": False}
    return {
        "has_dataset":     True,
        "total_rows":      result.total_rows,
        "columns":         result.columns,
        "captured_at":     result.captured_at,
        "filters_applied": result.filters_applied,
        "sql_preview":     result.sql[:300],
    }


# =============================================================================
# Export preview
# FIX 1: uses EXPORT_MAX_TOKENS (6144) for long reports
# =============================================================================
@app.post("/export/preview")
async def export_preview(req: ExportPreviewRequest):
    if not req.conversation:
        raise HTTPException(status_code=400, detail="No conversation to export.")

    print(f"📄 [export/preview] session={req.session_id} type={req.export_type}")

    stored = _get_result(req.session_id) if req.session_id else None

    try:
        ai_content = _generate_export_content(
            conversation   = req.conversation,
            title          = req.title,
            export_type    = req.export_type,
            detail_level   = req.detail_level,
            stored_dataset = stored,
        )
    except Exception as exc:
        traceback.print_exc()
        raise HTTPException(status_code=502, detail=f"Content generation error: {exc}")

    return {
        "content":       ai_content,
        "title":         req.title,
        "export_type":   req.export_type,
        "word_count":    len(ai_content.split()),
        "generated_at":  date.today().isoformat(),
        "total_rows":    stored.total_rows if stored else 0,
        "filters":       stored.filters_applied if stored else "",
    }


# =============================================================================
# Export download
# =============================================================================
@app.post("/export/download")
async def export_download(req: ExportDownloadRequest):
    print(f"⬇️  [export/download] format={req.format} session={req.session_id}")

    if req.format == "csv":
        stored = _get_result(req.session_id) if req.session_id else None
        if not stored:
            raise HTTPException(
                status_code=404,
                detail=(
                    "No query result found for this session. "
                    "Ask a deal-list question first, then export."
                ),
            )
        csv_bytes = _build_csv(stored)
        return StreamingResponse(
            io.BytesIO(csv_bytes),
            media_type="text/csv",
            headers={
                "Content-Disposition": f'attachment; filename="{_safe_filename(req.title)}.csv"',
                "X-Total-Rows": str(stored.total_rows),
            },
        )

    if not req.content:
        raise HTTPException(status_code=400, detail="content is required for PDF/PPTX export.")

    try:
        if req.format == "pdf":
            file_bytes = _build_pdf(req.title, req.content)
            media_type = "application/pdf"
            ext        = "pdf"
        else:
            file_bytes = _build_pptx(req.title, req.content)
            media_type = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
            ext        = "pptx"

        return StreamingResponse(
            io.BytesIO(file_bytes),
            media_type=media_type,
            headers={"Content-Disposition": f'attachment; filename="{_safe_filename(req.title)}.{ext}"'},
        )
    except Exception as exc:
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=f"File generation error: {exc}")


# =============================================================================
# Helpers
# =============================================================================
def _safe_filename(title: str) -> str:
    return re.sub(r'[^\w\-]', '_', title)[:60]


def _strip_md(t: str) -> str:
    t = re.sub(r'\*{1,3}(.*?)\*{1,3}', r'\1', t)
    t = re.sub(r'#{1,6}\s*', '', t)
    t = re.sub(r'^[\-\*•]\s*', '', t, flags=re.M)
    t = re.sub(r'`(.*?)`', r'\1', t)
    t = t.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
    return t.strip()


# =============================================================================
# CSV builder
# =============================================================================
def _build_csv(stored: QueryResult) -> bytes:
    buf = io.StringIO()
    buf.write(f"# Title: {stored.sql[:80]}\n")
    buf.write(f"# Generated: {date.today().isoformat()}\n")
    buf.write(f"# Total Records: {stored.total_rows}\n")
    buf.write(f"# Filters: {stored.filters_applied}\n")
    buf.write(f"# Captured at: {stored.captured_at}\n")
    buf.write("#\n")

    writer = csv.DictWriter(
        buf,
        fieldnames     = stored.columns,
        extrasaction   = "ignore",
        lineterminator = "\n",
    )
    writer.writeheader()
    for row in stored.rows:
        writer.writerow(row)

    return buf.getvalue().encode("utf-8")


# =============================================================================
# Export content generation
# FIX 1: EXPORT_MAX_TOKENS used here for long document generation
# =============================================================================

def _render_html_to_png(html_content: str, width: int = 800) -> Optional[bytes]:
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        print("⚠️  Playwright not installed.")
        return None

    full_html = f"""<!DOCTYPE html><html><head>
<meta charset="UTF-8">
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap" rel="stylesheet">
<style>
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{ background: #ffffff; font-family: 'Inter', Arial, sans-serif; }}
</style>
</head><body>{html_content}
</body></html>"""

    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(
                args=[
                    "--no-sandbox",
                    "--disable-setuid-sandbox",
                    "--disable-dev-shm-usage",
                    "--disable-gpu",
                    "--single-process",          # ← critical for Render
                    "--no-zygote",               # ← critical for Render
                ]
            )
            page = browser.new_page(viewport={"width": width, "height": 800})
            page.set_content(full_html, wait_until="networkidle", timeout=20000)
            page.wait_for_timeout(2000)
            height = page.evaluate("document.body.scrollHeight")
            page.set_viewport_size({"width": width, "height": max(height, 100)})
            png_bytes = page.screenshot(full_page=True)
            browser.close()
            return png_bytes
    except Exception as e:
        print(f"⚠️  Playwright screenshot failed: {e}")
        return None


def _extract_html_blocks(text: str):
    """
    Split a message into alternating text/html segments.
    Returns list of ("text"|"html", content) tuples.
    """
    segments = []
    last_end = 0
    for m in re.finditer(r'```html\s*\n(.*?)```', text, flags=re.DOTALL | re.IGNORECASE):
        segments.append(("text", text[last_end:m.start()]))
        segments.append(("html", m.group(1).strip()))
        last_end = m.end()
    segments.append(("text", text[last_end:]))
    return segments


def _generate_export_content(
    conversation:    List[ChatMessage],
    title:           str,
    export_type:     str,
    detail_level:    str = "detailed",
    stored_dataset:  Optional[QueryResult] = None,
) -> str:

    def _clean_text_segment(text: str) -> str:
        text = re.sub(r'```[\w]*\n.*?```', '', text, flags=re.DOTALL)
        text = re.sub(r'^\s*--\s*$', '', text, flags=re.MULTILINE)
        text = re.sub(r'^\s*`+\s*$', '', text, flags=re.MULTILINE)
        text = re.sub(r'\n{3,}', '\n\n', text)
        text = text.replace('■', '').replace('🟦', '')
        return text.strip()

    # ── VERBATIM MODE ────────────────────────────────────────────────────
    if detail_level == "detailed":
        if export_type == "pptx":
            lines = [f"# {title}", f"*Exported: {date.today().strftime('%B %d, %Y')}*", ""]
            for i, m in enumerate(conversation):
                role = "User" if m.role == "user" else "DIUD Agent"
                segments = _extract_html_blocks(m.content)
                lines.append(f"SLIDE: {role} — Turn {i + 1}")
                for seg_type, seg_content in segments:
                    if seg_type == "html":
                        lines.append("- [Interactive chart — see live DIUD interface]")
                    else:
                        cleaned = _clean_text_segment(seg_content)
                        for part in cleaned.split("\n"):
                            part = part.strip()
                            if part:
                                lines.append(f"- {part}")
            return "\n".join(lines)
        else:
            # PDF: encode segments as JSON with sentinel so _build_pdf can embed chart images
            pdf_segments = []
            for i, m in enumerate(conversation):
                role = "User" if m.role == "user" else "DIUD Agent"
                pdf_segments.append({"type": "turn_header", "turn": i + 1, "role": role})
                for seg_type, seg_content in _extract_html_blocks(m.content):
                    if seg_type == "html":
                        pdf_segments.append({"type": "html_chart", "html": seg_content})
                    else:
                        cleaned = _clean_text_segment(seg_content)
                        if cleaned:
                            pdf_segments.append({"type": "text", "content": cleaned})
            return "__VERBATIM_SEGMENTS__" + json.dumps(pdf_segments)

    # ── SUMMARY MODE ─────────────────────────────────────────────────────
    selected_model = ALLOWED_MODELS["sonnet"]
    conv_text = "\n\n".join(
        f"{'USER' if m.role == 'user' else 'DIUD AGENT'}: {m.content}"
        for m in conversation
    )
    format_hint = (
        "Format as a PowerPoint: use SLIDE: <title> for each slide, then bullet points."
        if export_type == "pptx"
        else "Format as a professional PDF report: ## section headers, narrative prose, tables."
    )
    dataset_hint = ""
    if stored_dataset:
        dataset_hint = (
            f"\n\nDATASET CONTEXT: The query returned {stored_dataset.total_rows} records "
            f"with columns: {', '.join(stored_dataset.columns[:12])}. "
            f"Filters: {stored_dataset.filters_applied}."
        )
    prompt = f"""You are preparing a professional {export_type.upper()} summary report.

CONVERSATION:
{conv_text}
{dataset_hint}

TASK: Create a concise executive summary titled "{title}"

{format_hint}

REQUIREMENTS:
- Executive summary at the start with key numbers only
- Logical sections: summary, key metrics, insights, recommendations
- Bold key numbers; clean professional tone
- Today: {date.today().strftime('%B %d, %Y')}
- Generate the COMPLETE document — do not truncate

Generate the summary report now:"""

    response = _ai_client.messages.create(
        model=selected_model,
        system="You are a professional document formatter. Follow the instructions exactly. Never add content not requested.",
        messages=[{"role": "user", "content": prompt}],
        temperature=0,
        max_tokens=EXPORT_MAX_TOKENS,
    )
    ai_text = _extract_text(response.content)

    if stored_dataset and stored_dataset.total_rows > 0:
        table_md  = _rows_to_markdown_table(stored_dataset)
        meta_line = (
            f"**Total records:** {stored_dataset.total_rows:,} | "
            f"**Filters:** {stored_dataset.filters_applied} | "
            f"**Exported:** {date.today().strftime('%B %d, %Y')}"
        )
        full_section = f"\n\n## Data Export ({stored_dataset.total_rows:,} records)\n\n{meta_line}\n\n{table_md}"
        if "[DEAL_TABLE_PLACEHOLDER]" in ai_text:
            ai_text = ai_text.replace("[DEAL_TABLE_PLACEHOLDER]", full_section)
        else:
            ai_text = ai_text.rstrip() + full_section

    return ai_text

def _rows_to_markdown_table(stored: QueryResult) -> str:
    if not stored.rows:
        return "_No data._"

    cols   = stored.columns
    header = "| " + " | ".join(cols) + " |"
    sep    = "| " + " | ".join("---" for _ in cols) + " |"
    lines  = [header, sep]

    for row in stored.rows:
        cells = [str(row.get(c, "")).replace("|", "\\|") for c in cols]
        lines.append("| " + " | ".join(cells) + " |")

    return "\n".join(lines)


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
    "deal":      colors.HexColor("#1565C0"),
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
_CW    = PW - _ML - _MR


def _pdf_styles():
    return {
        "Cover_Title": ParagraphStyle("Cover_Title", fontSize=26, leading=32,
            textColor=_C_WHITE, fontName="Helvetica-Bold", spaceAfter=8),
        "Cover_Sub": ParagraphStyle("Cover_Sub", fontSize=13, leading=18,
            textColor=colors.HexColor("#B0BEC5"), fontName="Helvetica"),
        "Section_H": ParagraphStyle("Section_H", fontSize=11, leading=15,
            textColor=_C_WHITE, fontName="Helvetica-Bold"),
        "Body": ParagraphStyle("Body", fontSize=9, leading=14, textColor=_C_TXT,
            fontName="Helvetica", spaceAfter=4),
        "Bullet": ParagraphStyle("Bullet", fontSize=9, leading=14, textColor=_C_TXT,
            fontName="Helvetica", leftIndent=12, firstLineIndent=-8, spaceAfter=3),
        "H2": ParagraphStyle("H2", fontSize=11, leading=15, textColor=_C_NAVY,
            fontName="Helvetica-Bold", spaceBefore=10, spaceAfter=4),
        "H3": ParagraphStyle("H3", fontSize=9, leading=13, textColor=_C_BLUE,
            fontName="Helvetica-Bold", spaceBefore=6, spaceAfter=2),
        "TH": ParagraphStyle("TH", fontSize=7, leading=9, textColor=_C_WHITE,
            fontName="Helvetica-Bold"),
        "TD": ParagraphStyle("TD", fontSize=7, leading=9, textColor=_C_TXT,
            fontName="Helvetica"),
    }


def _parse_sections(text: str):
    parts = re.split(r'^##\s+', text, flags=re.MULTILINE)
    return [
        (lines[0].strip(), lines[1].strip() if len(lines) > 1 else "")
        for part in parts if part.strip()
        for lines in [part.strip().split("\n", 1)]
    ]


def _build_pdf(title: str, report_text: str) -> bytes:
    from reportlab.platypus import Image as RLImage
    buf    = io.BytesIO()
    styles = _pdf_styles()

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

    cover_story = [
        Spacer(1, 1.0 * inch),
        Paragraph(title, styles["Cover_Title"]),
        Paragraph(f"Generated {date.today().strftime('%B %d, %Y')}", styles["Cover_Sub"]),
        PageBreak(),
    ]

    # ── VERBATIM SEGMENTS PATH ────────────────────────────────────────────
    if report_text.startswith("__VERBATIM_SEGMENTS__"):
        segments = json.loads(report_text[len("__VERBATIM_SEGMENTS__"):])
        story = cover_story[:]

        for seg in segments:
            if seg["type"] == "turn_header":
                role_label = "User" if seg["role"] == "User" else "DIUD Agent"
                bar_color  = _C_BLUE if seg["role"] == "User" else _C_NAVY
                story.append(Spacer(1, 8))
                story.append(Table(
                    [[Paragraph(f"Turn {seg['turn']} — {role_label}", styles["Section_H"])]],
                    colWidths=[_CW],
                    style=TableStyle([
                        ("BACKGROUND",    (0,0),(-1,-1), bar_color),
                        ("TOPPADDING",    (0,0),(-1,-1), 6),
                        ("BOTTOMPADDING", (0,0),(-1,-1), 6),
                        ("LEFTPADDING",   (0,0),(-1,-1), 10),
                    ])
                ))
                story.append(Spacer(1, 6))

            elif seg["type"] == "html_chart":
                png_bytes = _render_html_to_png(seg["html"], width=760)
                if png_bytes:
                    img = RLImage(io.BytesIO(png_bytes), width=_CW, height=_CW * 0.55)
                    img.hAlign = "CENTER"
                    story.append(img)
                    story.append(Spacer(1, 10))
                else:
                    story.append(Paragraph(
                        "[Chart could not be rendered — view in live DIUD interface]",
                        styles["Body"]
                    ))

            elif seg["type"] == "text":
                lines = seg["content"].split("\n")
                i = 0
                while i < len(lines):
                    line = lines[i].strip()
                    if line.startswith("|") and i + 1 < len(lines) and "---" in lines[i + 1]:
                        table_lines = []
                        while i < len(lines) and lines[i].strip().startswith("|"):
                            table_lines.append(lines[i].strip())
                            i += 1
                        story.append(_md_table_to_rl(table_lines, styles))
                        story.append(Spacer(1, 6))
                        continue
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
                    i += 1

        doc.build(story)
        return buf.getvalue()

    # ── SUMMARY PATH (existing logic, unchanged) ──────────────────────────
    sections = _parse_sections(report_text)
    story = cover_story[:]

    if not sections:
        # strip residual HTML before fallback rendering
        clean_text = re.sub(r'```html.*?```', '[Chart omitted]', report_text, flags=re.DOTALL | re.IGNORECASE)
        clean_text = re.sub(r'```.*?```', '', clean_text, flags=re.DOTALL)
        clean_text = re.sub(r'^\s*--\s*$', '', clean_text, flags=re.MULTILINE)
        for line in clean_text.split("\n"):
            line = line.strip()
            if line:
                story.append(Paragraph(_strip_md(line), styles["Body"]))
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

            lines = sec_body.split("\n")
            i = 0
            while i < len(lines):
                line = lines[i].strip()
                if line.startswith("|") and i + 1 < len(lines) and "---" in lines[i + 1]:
                    table_lines = []
                    while i < len(lines) and lines[i].strip().startswith("|"):
                        table_lines.append(lines[i].strip())
                        i += 1
                    story.append(_md_table_to_rl(table_lines, styles))
                    story.append(Spacer(1, 6))
                    continue
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
                i += 1

            story.extend([Spacer(1, 12), PageBreak()])

    doc.build(story)
    return buf.getvalue()


def _md_table_to_rl(table_lines: list, styles: dict):
    data = []
    for idx, line in enumerate(table_lines):
        if "---" in line:
            continue
        cells = [c.strip().replace("\\|", "|") for c in line.strip("|").split("|")]
        if idx == 0:
            row = [Paragraph(_strip_md(c), styles["TH"]) for c in cells]
        else:
            row = [Paragraph(_strip_md(c), styles["TD"]) for c in cells]
        data.append(row)

    if not data:
        return Spacer(1, 1)

    num_cols = max(len(r) for r in data)
    col_w    = _CW / max(num_cols, 1)

    tbl = Table(data, colWidths=[col_w] * num_cols, repeatRows=1)
    tbl.setStyle(TableStyle([
        ("BACKGROUND",    (0, 0), (-1, 0),  _C_NAVY),
        ("TEXTCOLOR",     (0, 0), (-1, 0),  _C_WHITE),
        ("FONTNAME",      (0, 0), (-1, 0),  "Helvetica-Bold"),
        ("FONTSIZE",      (0, 0), (-1, -1), 7),
        ("ROWBACKGROUNDS",(0, 1), (-1, -1), [colors.white, colors.HexColor("#F8FAFC")]),
        ("GRID",          (0, 0), (-1, -1), 0.3, colors.HexColor("#E2E8F0")),
        ("TOPPADDING",    (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("LEFTPADDING",   (0, 0), (-1, -1), 4),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 4),
        ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
    ]))
    return tbl


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
    "deal":      RGBColor(0x1E, 0x88, 0xE5),
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
    prs.slide_width  = Inches(13.33)
    prs.slide_height = Inches(7.5)
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
        _pptx_txt(slide, str(i + 1), 12.5, 0.12, 0.6, 0.6, size=11, color=_C_WHITE_P, align=PP_ALIGN.RIGHT)
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
