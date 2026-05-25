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

app = FastAPI(title="DIUD", description="Decision Intelligence Using Data", version="2.0.0")

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

_ai_client   = anthropic.Anthropic(api_key=os.getenv("ANTHROPIC_API_KEY"))
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

# Will be populated at startup: { "database.table": [{"name":..,"type":..}, ...], ... }
_LIVE_SCHEMA: dict = {}
# Human-readable string injected into the system prompt
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
    """
    Walk /databases → /tables/{db} → /schema/{db}/{table}.
    Returns a formatted schema string for the system prompt.
    Populates _LIVE_SCHEMA global.
    """
    global _LIVE_SCHEMA, _SCHEMA_BLOCK

    print("🔎 Discovering schema from ClickHouse proxy…")

    databases_raw = _proxy_get("/databases")
    if not databases_raw:
        msg = "⚠️  Could not fetch databases — check CLICKHOUSE_API_URL / CLICKHOUSE_API_TOKEN."
        print(msg)
        _SCHEMA_BLOCK = msg
        return msg

    # Normalise: could be a list of strings or list of dicts
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

    # Skip system databases
    SKIP_DBS = {"system", "information_schema", "INFORMATION_SCHEMA"}
    databases = [d for d in databases if d not in SKIP_DBS]
    print(f"   Databases found: {databases}")

    schema_lines = []
    schema_dict  = {}

    for db in databases:
        tables_raw = _proxy_get(f"/tables/{db}")
        if not tables_raw:
            continue

        # Normalise table list
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

            # Normalise columns list
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
    """
    Assembles the full system prompt with a compact schema block.
    Only includes table names, column names, and types — no comments or padding.
    This avoids hitting the 200k token limit with large schemas.
    """
    table_refs = list(_LIVE_SCHEMA.keys())
    primary_table = table_refs[0] if table_refs else "your_database.deals"

    # Compact schema: table(col:type, col:type, ...) — no comments, no padding
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

    # Safety cap — should rarely trigger with compact format, but prevents edge cases
    if len(schema) > 20000:
        schema = schema[:20000] + "\n[schema truncated — too many tables/columns]"

    return f"""
You are DIUD (Decision Intelligence Using Data) — a conversational data assistant.
You have LIVE access to a ClickHouse database via the query_clickhouse tool.

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

  -- QUALIFICATION FIELDS
  is_there_a_confirmation_of_budget  STRING  — 'Yes'/'No'
  who_is_the_decision_maker          STRING  — decision maker name
  use_case                           STRING  — use case description
  what_is_the_estimated_timeline     STRING  — timeline string
  is_this_a_deal_with_inception      STRING  — 'Yes'/'No'

  -- WIN/LOSS FIELDS
  primary_closed_won_reason_         STRING  — win reason
  primary_closed_lost_reason         STRING  — loss reason
  won_loss_notes                     STRING  — freeform notes
  competitors                        STRING  — competitor names
  competition                        STRING  — competition notes

  -- DEAL APPROVAL FIELDS
  cs_deal_approval_status_level_1      STRING
  cs_deal_approval_status_level_2      STRING
  direct_deal_approval_status_level_1  STRING
  direct_deal_approval_status_level_2  STRING
  deal_approval_status_level_1         STRING
  deal_approval_status_level_2         STRING
  deal_approval_status_level_3_cs_only STRING

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

── HELPER TABLES ─────────────────────────────────────────────────
  kore_ai_hubspot.gs_deal_ids_hs
    deal_id_hs  STRING   — valid deal IDs whitelist

  kore_ai_hubspot.gs_Teams
    team_id     STRING
    name        STRING

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

Macro (replace <date_col> with the relevant column):
  toYear(toDate(LEFT(coalesce(<date_col>,'1900-01-01'),10)))
  + if(toMonth(toDate(LEFT(coalesce(<date_col>,'1900-01-01'),10))) >= 4, 1, 0)

FY27 5%  cohort → became_5_deal_date  >= '2026-04-01'
FY27 20% cohort → became_20_deal_date >= '2026-04-01'
FY26 5%  cohort → became_5_deal_date  >= '2025-04-01' AND < '2026-04-01'

=================================================================
COMPUTED COLUMNS — use inline in queries
=================================================================

-- Owner full name (requires LEFT JOIN to owners):
  concat(o.firstName, ' ', o.lastName) AS deal_owner_name

-- Region display mapping:
  CASE
    WHEN d.region = 'japac'       THEN 'JAPAC'
    WHEN d.region = 'Africa'      THEN 'Middle East'
    WHEN d.region = 'india___sea' THEN 'ISEA'
    ELSE d.region
  END AS region

  RAW → DISPLAY:
    'japac'        → 'JAPAC'
    'Africa'       → 'Middle East'
    'india___sea'  → 'ISEA'
    Others unchanged: 'North America', 'EMEA', 'APAC', 'India', 'Latin America'

-- Deal source mapping:
  CASE
    WHEN d.deal_source_rollup IN ('Executive Outreach','Investor') THEN 'Executive Outreach'
    WHEN d.deal_source_rollup IN ('BDR Outbound')                  THEN 'BDR'
    WHEN d.deal_source_rollup IN ('Partner')                       THEN 'Partner - Non Hyperscaler'
    WHEN d.deal_source_rollup IN ('Marketing','Customer Success',
         'AE Outbound','Inception','Hyperscaler')                  THEN d.deal_source_rollup
    ELSE 'Other'
  END AS deal_source

-- Industry mapping:
  CASE
    WHEN d.kore_primary_industry IN ('Financial Services','Banking','Insurance')
         THEN 'Financial Services'
    WHEN d.kore_primary_industry IN ('Manufacturing Discreet','Manufacturing Process','CPG')
         THEN 'Manufacturing'
    WHEN d.kore_primary_industry IN ('Hi-Tech','Telecom / Media / Entertainment')
         THEN 'TMT'
    WHEN d.kore_primary_industry IS NULL
      OR d.kore_primary_industry IN ('Business Services','Government','Energy & Utilities',
         'Education','Restaurants','null','Energy')
         THEN 'Other'
    ELSE d.kore_primary_industry
  END AS industry

-- Stage category:
  CASE
    WHEN d.deal_stage IN ('20% - Solution','30% - Proof','40% - Proposal',
         '60% - Price Negotiation','75% - Contract Review')
         THEN 'Active Pipeline'
    WHEN d.deal_stage IN ('Prospect Disengaged','Closed Lost','Didn''t Qualify')
         THEN 'Fallen Out'
    WHEN d.deal_stage IN ('90% - Deal Desk Review','Closed Won')
         THEN 'Closed Won'
    ELSE 'Pre-Qualification'
  END AS stage_category

-- BANT qualification:
  CASE
    WHEN d.is_there_a_confirmation_of_budget = 'Yes'
     AND d.who_is_the_decision_maker IS NOT NULL
     AND d.use_case IS NOT NULL
     AND d.what_is_the_estimated_timeline IS NOT NULL
    THEN 'Yes' ELSE 'No'
  END AS BANT

-- Account priority grouping:
  CASE
    WHEN d.account_priority_level IN ('P1','P2','P3','P4') THEN 'P1-P4'
    WHEN d.account_priority_level IN ('P5','P6','P7')      THEN 'P5-P7'
    WHEN d.account_priority_level IN ('P8','P9','P10')     THEN 'P8-P10'
    ELSE 'No Priority'
  END AS acct_priority

-- Days in a stage (example for 10% Discovery):
  DATE_DIFF('Day',
    toDate(LEFT(coalesce(d.became_10_deal_date,'1900-01-01'),10)),
    CURRENT_DATE()
  ) AS days_in_10

=================================================================
DEAL STAGE LIST (funnel order, with velocity benchmarks)
=================================================================
  '1% - IQM Scheduled'       → Pre-Qualification  (target ≤ 7 days)
  '5% - IQM Held'            → Pre-Qualification  (target ≤ 21 days)
  '10% - Discovery'          → Pre-Qualification  (target ≤ 28 days)
  '20% - Solution'           → Active Pipeline    (target ≤ 41 days)
  '30% - Proof'              → Active Pipeline    (target ≤ 15 days)
  '40% - Proposal'           → Active Pipeline    (target ≤ 29 days)
  '60% - Price Negotiation'  → Active Pipeline    (target ≤ 27 days)
  '75% - Contract Review'    → Active Pipeline    (target ≤ 34 days)
  '90% - Deal Desk Review'   → Closed Won
  'Closed Won'               → Closed Won
  'Closed Lost'              → Fallen Out
  "Didn't Qualify"           → Fallen Out
  'Prospect Disengaged'      → Fallen Out
  'Deal on Hold'             → Pre-Qualification

=================================================================
BUSINESS DEFINITIONS
=================================================================
"Active pipeline"   → deal_stage IN ('20% - Solution','30% - Proof','40% - Proposal',
                      '60% - Price Negotiation','75% - Contract Review')
"Qualified deals"   → became_20_deal_date <> '1900-01-01' AND became_20_deal_date IS NOT NULL
"Fallen out"        → deal_stage IN ('Prospect Disengaged','Closed Lost','Didn''t Qualify')
"Closed won"        → deal_stage IN ('Closed Won','90% - Deal Desk Review')
"BANT qualified"    → all 4 BANT fields confirmed (budget, decision maker, use_case, timeline)
"High priority"     → account_priority_level IN ('P1','P2','P3','P4')
"FY27 5% cohort"    → became_5_deal_date >= '2026-04-01'
"FY27 20% cohort"   → became_20_deal_date >= '2026-04-01'
"Stalled deal"      → in active pipeline stage for > 2× the stage benchmark days
"At-risk deal"      → closing within 30 days AND still in 20–40% stage
"Coverage ratio"    → active pipeline value ÷ revenue target (healthy ≥ 3×)

=================================================================
QUERY RULES
=================================================================
1. SELECT / WITH only — never INSERT, UPDATE, DELETE, DROP, ALTER, TRUNCATE
2. Always use FINAL on all hs_analytics tables
3. Always apply the 3 mandatory base filters on deals
4. Always LIMIT row-level queries (max 100 rows)
5. Use countDistinct(deal_id) for unique deal counts
6. Use round(sum(amount)/1e6, 1) for $M dollar amounts
7. Use ILIKE for case-insensitive text matching
8. Dates are stored as strings — always cast: toDate(LEFT(coalesce(col,'1900-01-01'),10))
9. Null date sentinel is '1900-01-01' — exclude with: col <> '1900-01-01' AND col IS NOT NULL
10. Default fiscal year context is FY27 (Apr 2026 – Mar 2027) unless user specifies otherwise

=================================================================
SAMPLE QUERIES
=================================================================

-- Count + value of active pipeline (FY27 5% cohort):
SELECT
  countDistinct(d.deal_id) AS active_deals,
  round(sum(d.amount)/1e6, 1) AS pipeline_m
FROM hs_analytics.deals d FINAL
WHERE d.pipeline = 'default'
  AND CASE WHEN d.deal_type IS NULL THEN 'Not Assigned' ELSE d.deal_type END
      NOT IN ('Partner-Led SMB')
  AND toInt64(d.deal_id) IN (SELECT DISTINCT toInt64(deal_id_hs) FROM kore_ai_hubspot.gs_deal_ids_hs)
  AND d.deal_stage IN ('20% - Solution','30% - Proof','40% - Proposal',
                       '60% - Price Negotiation','75% - Contract Review')
  AND d.became_5_deal_date >= '2026-04-01'

-- Count deals in each stage right now (FY27 5% cohort):
SELECT
  d.deal_stage,
  countDistinct(d.deal_id) AS deal_count,
  round(sum(d.amount)/1e6, 1) AS pipeline_m
FROM hs_analytics.deals d FINAL
WHERE d.pipeline = 'default'
  AND CASE WHEN d.deal_type IS NULL THEN 'Not Assigned' ELSE d.deal_type END
      NOT IN ('Partner-Led SMB')
  AND toInt64(d.deal_id) IN (SELECT DISTINCT toInt64(deal_id_hs) FROM kore_ai_hubspot.gs_deal_ids_hs)
  AND d.became_5_deal_date >= '2026-04-01'
GROUP BY d.deal_stage
ORDER BY pipeline_m DESC

-- Top 10 deals by value with owner:
SELECT
  d.deal_name,
  concat(o.firstName,' ',o.lastName) AS owner,
  CASE WHEN d.region='japac' THEN 'JAPAC' WHEN d.region='Africa' THEN 'Middle East'
       WHEN d.region='india___sea' THEN 'ISEA' ELSE d.region END AS region,
  d.deal_stage,
  round(d.amount/1e6, 2) AS amt_m,
  toDate(LEFT(coalesce(d.close_date,'1900-01-01'),10)) AS close_date
FROM hs_analytics.deals d FINAL
LEFT JOIN hs_analytics.owners o FINAL ON d.deal_owner = CAST(o.id AS VARCHAR)
WHERE d.pipeline = 'default'
  AND CASE WHEN d.deal_type IS NULL THEN 'Not Assigned' ELSE d.deal_type END NOT IN ('Partner-Led SMB')
  AND toInt64(d.deal_id) IN (SELECT DISTINCT toInt64(deal_id_hs) FROM kore_ai_hubspot.gs_deal_ids_hs)
  AND d.deal_stage IN ('20% - Solution','30% - Proof','40% - Proposal',
                       '60% - Price Negotiation','75% - Contract Review')
  AND d.became_5_deal_date >= '2026-04-01'
ORDER BY d.amount DESC LIMIT 10

-- Pipeline breakdown by region (FY27):
SELECT
  CASE WHEN d.region='japac' THEN 'JAPAC' WHEN d.region='Africa' THEN 'Middle East'
       WHEN d.region='india___sea' THEN 'ISEA' ELSE d.region END AS region,
  countDistinct(d.deal_id) AS deals,
  round(sum(d.amount)/1e6, 1) AS pipeline_m
FROM hs_analytics.deals d FINAL
WHERE d.pipeline = 'default'
  AND CASE WHEN d.deal_type IS NULL THEN 'Not Assigned' ELSE d.deal_type END NOT IN ('Partner-Led SMB')
  AND toInt64(d.deal_id) IN (SELECT DISTINCT toInt64(deal_id_hs) FROM kore_ai_hubspot.gs_deal_ids_hs)
  AND d.became_5_deal_date >= '2026-04-01'
  AND d.deal_stage IN ('20% - Solution','30% - Proof','40% - Proposal',
                       '60% - Price Negotiation','75% - Contract Review')
GROUP BY region ORDER BY pipeline_m DESC

-- Win rate by deal source (FY27 5% cohort):
SELECT
  CASE
    WHEN d.deal_source_rollup IN ('Executive Outreach','Investor') THEN 'Executive Outreach'
    WHEN d.deal_source_rollup IN ('BDR Outbound') THEN 'BDR'
    WHEN d.deal_source_rollup IN ('Partner') THEN 'Partner - Non Hyperscaler'
    ELSE coalesce(d.deal_source_rollup,'Other')
  END AS deal_source,
  countDistinct(CASE WHEN d.deal_stage IN ('Closed Won','90% - Deal Desk Review') THEN d.deal_id END) AS won,
  countDistinct(CASE WHEN d.deal_stage = 'Closed Lost' THEN d.deal_id END) AS lost,
  round(
    countDistinct(CASE WHEN d.deal_stage IN ('Closed Won','90% - Deal Desk Review') THEN d.deal_id END) * 100.0
    / nullIf(countDistinct(CASE WHEN d.deal_stage IN ('Closed Won','90% - Deal Desk Review','Closed Lost') THEN d.deal_id END), 0)
  , 1) AS win_rate_pct
FROM hs_analytics.deals d FINAL
WHERE d.pipeline = 'default'
  AND CASE WHEN d.deal_type IS NULL THEN 'Not Assigned' ELSE d.deal_type END NOT IN ('Partner-Led SMB')
  AND toInt64(d.deal_id) IN (SELECT DISTINCT toInt64(deal_id_hs) FROM kore_ai_hubspot.gs_deal_ids_hs)
  AND d.became_5_deal_date >= '2026-04-01'
GROUP BY deal_source ORDER BY won DESC

-- BANT qualification rate across active pipeline:
SELECT
  countDistinct(d.deal_id) AS total_active,
  countDistinct(CASE
    WHEN d.is_there_a_confirmation_of_budget = 'Yes'
     AND d.who_is_the_decision_maker IS NOT NULL
     AND d.use_case IS NOT NULL
     AND d.what_is_the_estimated_timeline IS NOT NULL
    THEN d.deal_id END) AS bant_qualified,
  round(
    countDistinct(CASE
      WHEN d.is_there_a_confirmation_of_budget = 'Yes'
       AND d.who_is_the_decision_maker IS NOT NULL
       AND d.use_case IS NOT NULL
       AND d.what_is_the_estimated_timeline IS NOT NULL
      THEN d.deal_id END) * 100.0
    / nullIf(countDistinct(d.deal_id), 0)
  , 1) AS bant_rate_pct
FROM hs_analytics.deals d FINAL
WHERE d.pipeline = 'default'
  AND CASE WHEN d.deal_type IS NULL THEN 'Not Assigned' ELSE d.deal_type END NOT IN ('Partner-Led SMB')
  AND toInt64(d.deal_id) IN (SELECT DISTINCT toInt64(deal_id_hs) FROM kore_ai_hubspot.gs_deal_ids_hs)
  AND d.deal_stage IN ('20% - Solution','30% - Proof','40% - Proposal',
                       '60% - Price Negotiation','75% - Contract Review')
  AND d.became_5_deal_date >= '2026-04-01'

CORE RULES:
- NEVER say you lack database access. You always have it via the tool.
- NEVER fabricate numbers. Query the DB for every metric question.
- NEVER run destructive SQL (INSERT / UPDATE / DELETE / DROP / ALTER / TRUNCATE).
- If query_clickhouse returns text starting with DATABASE CONNECTION FAILED or ERROR:, "
  STOP immediately and show the user this exact message: "
  '⚠️ Database unreachable. Please check your connection and try again.' "
  DO NOT attempt alternative queries or explain what you were going to do."
- ALWAYS use the fully qualified table name: database.table (e.g. {primary_table}).
- Answer in clean markdown: use tables for data, bold for key numbers.
- Be concise but complete.
- When generating export content, use ## section headers.


"""



# Initialise with a placeholder — will be replaced on startup
_SYSTEM_PROMPT = _build_system_prompt()


# =============================================================================
# ClickHouse query runner
# =============================================================================

def run_clickhouse_query(sql: str) -> str:
    """
    POST /query  body: {{"query": "<SQL>"}}
    Returns plain-text table or error string starting with ERROR:/DATABASE.
    Never raises.
    """
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
            json={"query": sql},   # ← "query" field confirmed by your OpenAPI spec
            timeout=30,
        )

        if resp.status_code == 401:
            return "DATABASE CONNECTION FAILED: 401 Unauthorized — check CLICKHOUSE_API_TOKEN."
        if resp.status_code == 403:
            return "DATABASE CONNECTION FAILED: 403 Forbidden."
        if resp.status_code == 422:
            detail = resp.text[:600]
            print(f"   422: {detail}")
            return f"ERROR: Proxy rejected the query (422). Detail: {detail}"
        if resp.status_code == 500:
            detail = resp.text[:600]
            print(f"   500: {detail}")
            return (
                f"DATABASE ERROR: HTTP 500 Internal Server Error from the proxy.\n"
                f"This usually means the SQL references a table or column that doesn't exist.\n"
                f"Detail: {detail}"
            )
        if resp.status_code != 200:
            return f"DATABASE ERROR: HTTP {resp.status_code} — {resp.text[:400]}"

        payload = resp.json()
        print(f"   Response shape: {type(payload).__name__}, "
              f"keys={list(payload.keys()) if isinstance(payload, dict) else 'list'}")

        # Normalise to a list of rows
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

        # Build text table
        if isinstance(rows[0], dict):
            cols   = list(rows[0].keys())
            header = " | ".join(cols)
            lines  = [header, "-" * min(len(header), 140)]
            for row in rows[:100]:
                lines.append(" | ".join(str(row.get(c, "NULL")) for c in cols))
        else:
            lines = [" | ".join(str(v) for v in rows[0]), "-" * 80]
            for row in rows[1:101]:
                lines.append(" | ".join(str(v) for v in row))

        if len(rows) > 100:
            lines.append(f"... ({len(rows) - 100} more rows not shown)")

        result = "\n".join(lines)
        print(f"   ✅ {len(rows)} rows. Preview: {result[:200]}")
        return result

    except httpx.ConnectError as e:
        return f"DATABASE CONNECTION FAILED: Could not reach {base_url}. Detail: {e}"
    except httpx.TimeoutException:
        return "DATABASE CONNECTION FAILED: Query timed out after 30 seconds."
    except Exception as exc:
        traceback.print_exc()
        return f"DATABASE CONNECTION FAILED: {type(exc).__name__}: {exc}"


# =============================================================================
# Startup event — discover schema and build system prompt
# =============================================================================

@app.on_event("startup")
async def on_startup():
    global _SYSTEM_PROMPT
    discover_schema()
    _SYSTEM_PROMPT = _build_system_prompt()
    print("🚀 System prompt built with live schema.")
    print(_SYSTEM_PROMPT[:800])


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
                    "LIMIT row queries to 100."
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
    """Run Claude with query_clickhouse tool. Up to 5 tool rounds."""
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

        sql          = tool_block.input.get("sql", "")
        query_result = run_clickhouse_query(sql)
        is_error     = any(query_result.startswith(p) for p in [
            "DATABASE CONNECTION FAILED", "ERROR:", "DATABASE ERROR:"
        ])

        messages = messages + [
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
            messages=messages,
            tools=[_QUERY_TOOL],
            temperature=0,
            max_tokens=max_tokens,
        )

    reply = _extract_text(response.content)
    if not reply:
        reply = (
            "⚠️ No response was generated.\n\n"
            "This usually means the ClickHouse query failed. "
            "Open **/debug/db** to diagnose, then verify `CLICKHOUSE_API_URL` "
            "and `CLICKHOUSE_API_TOKEN` in `.env` and restart."
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
    """Full connectivity + schema diagnostic. Open: http://localhost:8000/debug/db"""
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

    # Health check
    try:
        r = httpx.get(base_url, timeout=10)
        tests["GET /"] = {"status": r.status_code, "body": r.text[:200]}
    except Exception as e:
        tests["GET /"] = {"error": str(e)}

    # List databases
    try:
        r = httpx.get(f"{base_url}/databases", headers=_auth_headers(), timeout=10)
        tests["GET /databases"] = {"status": r.status_code, "body": r.text[:400]}
    except Exception as e:
        tests["GET /databases"] = {"error": str(e)}

    # Ping query
    ping = run_clickhouse_query("SELECT 1 AS ping")
    query_ok = not any(ping.startswith(p) for p in [
        "DATABASE CONNECTION FAILED", "ERROR:", "DATABASE ERROR:"
    ])
    tests["POST /query (SELECT 1)"] = {"result": ping, "ok": query_ok}

    return {
        "status":           "OK" if query_ok else "FAILED",
        "config":           config,
        "discovered_tables": list(_LIVE_SCHEMA.keys()),
        "tests":            tests,
        "recommendation": (
            "✅ DB connected. Schema loaded. Chat is ready."
            if query_ok else
            "❌ Query test failed. See 'POST /query' result for the exact error.\n"
            "Common causes:\n"
            "  401 → wrong token\n"
            "  ConnectError → wrong URL or proxy down\n"
            "  500 → SQL error (wrong table/column)"
        ),
    }


@app.post("/refresh-schema")
def refresh_schema():
    """Re-discover schema from the proxy and rebuild the system prompt."""
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


@app.post("/export/pdf")
def export_pdf(payload: ExportRequest):
    conv_text = "\n\n".join(
        f"{'User' if m.role == 'user' else 'Assistant'}: {m.content}"
        for m in payload.conversation
    )
    prompt = (
        f"Based on this conversation, write a structured deals intelligence report "
        f"titled '{payload.title}'. Use these ## section headers:\n"
        "## Executive Summary\n## Pipeline Overview\n## Key Metrics\n"
        "## Regional Breakdown\n## Win / Loss Analysis\n## Recommendations\n\n"
        "Query the database for any missing numbers.\n\n"
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
        f"Based on this conversation, write slide content for '{payload.title}'.\n"
        "Output each slide as:\nSLIDE: <Title>\nBULLETS:\n- bullet 1\n- bullet 2\n\n"
        "Include: Overview, Pipeline Health, Key Metrics, Regional Breakdown, "
        "Win/Loss Analysis, Recommendations.\n"
        "Query DB for any missing data.\n\n"
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
