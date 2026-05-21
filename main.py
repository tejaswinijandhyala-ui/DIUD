"""
main.py
PipeGen Chat — FastAPI Backend
Conversational AI over ClickHouse pipeline data, with PDF & PPTX export.
"""

# ── Standard Library ──────────────────────────────────────────────────────────
import io
import json
import os
import re
import traceback
from datetime import date
from typing import List, Literal, Optional

# ── Third-Party: Web Framework ────────────────────────────────────────────────
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import HTMLResponse, StreamingResponse
from pydantic import BaseModel

# ── Third-Party: Environment ──────────────────────────────────────────────────
from dotenv import load_dotenv

# ── Third-Party: AI ───────────────────────────────────────────────────────────
import anthropic

# ── Third-Party: HTTP ─────────────────────────────────────────────────────────
import requests as http_requests

# ── Third-Party: PPTX ────────────────────────────────────────────────────────
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt, Emu

# ── Third-Party: PDF ─────────────────────────────────────────────────────────
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import (
    BaseDocTemplate, Frame, PageTemplate, Flowable,
    HRFlowable, NextPageTemplate, PageBreak,
    Paragraph, Spacer, Table, TableStyle,
)

load_dotenv()

# =============================================================================
# App
# =============================================================================

app = FastAPI(
    title="PipeGen Chat",
    description="Conversational pipeline intelligence powered by Claude and ClickHouse.",
    version="1.0.0",
    redirect_slashes=False,
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

_ai_client = anthropic.Anthropic(api_key=os.getenv("ANTHROPIC_API_KEY"))
_CLAUDE_MODEL = "claude-opus-4-5"


# =============================================================================
# ClickHouse Query Runner
# =============================================================================

def _parse_ch_response(resp: http_requests.Response) -> str:
    """
    Parse a ClickHouse HTTP response into a readable pipe-delimited table.
    Handles JSONCompact, JSON, and plain TSV/CSV fallbacks.
    """
    resp.raise_for_status()
    content_type = resp.headers.get("content-type", "")

    # ── JSONCompact / JSON ────────────────────────────────────────────────────
    if "json" in content_type or resp.text.lstrip().startswith("{"):
        data    = resp.json()
        columns = [c["name"] for c in data.get("meta", [])]
        rows    = data.get("data", [])
        if not rows:
            return "Query returned 0 rows."
        capped  = rows[:100]
        header  = " | ".join(columns)
        divider = "-" * min(len(header), 120)
        lines   = [header, divider]
        for row in capped:
            lines.append(" | ".join(str(v) if v is not None else "NULL" for v in row))
        if len(rows) > 100:
            lines.append(f"... ({len(rows) - 100} more rows — refine your query)")
        return "\n".join(lines)

    # ── TSV / plain text fallback ─────────────────────────────────────────────
    lines = [l for l in resp.text.splitlines() if l.strip()]
    if not lines:
        return "Query returned 0 rows."
    return "\n".join(lines[:101])


def run_clickhouse_query(sql: str) -> str:
    """
    Execute a read-only SELECT/WITH against the ClickHouse HTTP API.

    Tries four strategies in order so the function works regardless of how
    the endpoint is configured:
      1. POST  — SQL in request body, FORMAT in body          (most common for ClickHouse Cloud / proxy)
      2. POST  — SQL as ?query= param, FORMAT appended        (original AI4Looker style)
      3. GET   — SQL as ?query= param                         (some self-hosted setups)
      4. POST  — SQL in body as plain text, no FORMAT param   (bare ClickHouse HTTP interface)

    Auth is tried as Bearer token first, then as basic-auth username from the token.
    """
    api_url   = os.getenv("CLICKHOUSE_API_URL", "").rstrip("/")
    api_token = os.getenv("CLICKHOUSE_API_TOKEN", "")

    if not api_url or not api_token:
        return (
            "ClickHouse not configured.\n"
            "Set CLICKHOUSE_API_URL and CLICKHOUSE_API_TOKEN in your .env file."
        )

    stripped = sql.strip().upper()
    if not stripped.startswith("SELECT") and not stripped.startswith("WITH"):
        return "Error: Only SELECT/WITH queries are permitted."

    sql_with_fmt    = sql.strip() + " FORMAT JSONCompact"
    bearer_headers  = {"Authorization": f"Bearer {api_token}",
                       "Content-Type": "text/plain"}
    errors = []

    # ── Strategy 1: POST body (most common for proxied/cloud ClickHouse) ──────
    try:
        print(f"  🔌 CH strategy 1: POST body")
        resp = http_requests.post(
            api_url,
            data=sql_with_fmt.encode("utf-8"),
            headers=bearer_headers,
            timeout=30,
        )
        if resp.status_code not in (405, 404, 403):
            return _parse_ch_response(resp)
        errors.append(f"S1 {resp.status_code}: {resp.text[:120]}")
    except Exception as e:
        errors.append(f"S1 exc: {e}")

    # ── Strategy 2: POST ?query= param (original AI4Looker approach) ──────────
    try:
        print(f"  🔌 CH strategy 2: POST ?query=")
        resp = http_requests.post(
            api_url,
            params={"query": sql_with_fmt},
            headers={"Authorization": f"Bearer {api_token}"},
            timeout=30,
        )
        if resp.status_code not in (405, 404, 403):
            return _parse_ch_response(resp)
        errors.append(f"S2 {resp.status_code}: {resp.text[:120]}")
    except Exception as e:
        errors.append(f"S2 exc: {e}")

    # ── Strategy 3: GET ?query= param ─────────────────────────────────────────
    try:
        print(f"  🔌 CH strategy 3: GET ?query=")
        resp = http_requests.get(
            api_url,
            params={"query": sql_with_fmt},
            headers={"Authorization": f"Bearer {api_token}"},
            timeout=30,
        )
        if resp.status_code not in (405, 404, 403):
            return _parse_ch_response(resp)
        errors.append(f"S3 {resp.status_code}: {resp.text[:120]}")
    except Exception as e:
        errors.append(f"S3 exc: {e}")

    # ── Strategy 4: POST body without FORMAT (bare ClickHouse HTTP) ───────────
    try:
        print(f"  🔌 CH strategy 4: POST body plain")
        resp = http_requests.post(
            api_url,
            data=sql.strip().encode("utf-8"),
            headers={"Authorization": f"Bearer {api_token}",
                     "Content-Type": "text/plain",
                     "X-ClickHouse-Format": "JSONCompact"},
            timeout=30,
        )
        if resp.status_code < 400:
            return _parse_ch_response(resp)
        errors.append(f"S4 {resp.status_code}: {resp.text[:120]}")
    except Exception as e:
        errors.append(f"S4 exc: {e}")

    # ── All strategies failed ──────────────────────────────────────────────────
    print(f"  ❌ All CH strategies failed: {errors}")
    return (
        "ClickHouse connection failed across all strategies.\n"
        "Please verify CLICKHOUSE_API_URL and CLICKHOUSE_API_TOKEN in .env.\n"
        "Debug info (check server logs): " + " | ".join(errors)
    )


# =============================================================================
# System Prompt — Schema, Definitions, Example Queries
# =============================================================================
#
# IMPORTANT FOR DEVELOPERS:
#   This is the master knowledge prompt injected into every chat turn.
#   Customize the sections marked with TODO to match your exact database,
#   business logic, and KPI definitions before deploying.
#
_SYSTEM_PROMPT = """
You are PipeGen Chat — a conversational pipeline intelligence assistant for Kore.ai.
You have DIRECT, LIVE access to the ClickHouse database via the query_clickhouse tool.

RULES:
- NEVER say you lack database access. You always have it.
- NEVER fabricate numbers. Query the DB if the answer isn't in your context.
- NEVER run destructive SQL (no INSERT/UPDATE/DELETE/DROP/ALTER/TRUNCATE).
- Always answer in clean markdown: use tables for data, bold for KPIs.
- Be concise but complete. If the user asks for a summary, give one.
- When building export content (PDF/PPTX), structure the response with
  clear section headers so the export functions can parse it.

=================================================================
CLICKHOUSE DIRECT ACCESS
=================================================================

You have a tool called query_clickhouse.
Use it for any question about pipeline deals, AEs, regions, industries,
stages, win/loss, competitors, conversions, or any metric not already
in the conversation context.

=================================================================
TABLES
=================================================================

── TABLE 1: hs_analytics.deals ─────────────────────────────────
Primary table. One row per deal.
Always use FINAL keyword: FROM hs_analytics.deals FINAL

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

  -- DATE COLUMNS (stored as strings, always cast to DATE before comparison)
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
  cs_deal_approval_status_level_1    STRING
  cs_deal_approval_status_level_2    STRING
  direct_deal_approval_status_level_1 STRING
  direct_deal_approval_status_level_2 STRING
  deal_approval_status_level_1       STRING
  deal_approval_status_level_2       STRING
  deal_approval_status_level_3_cs_only STRING

── TABLE 2: hs_analytics.owners ─────────────────────────────────
Always use FINAL: FROM hs_analytics.owners FINAL

  id           STRING  — owner ID (join key to deals.deal_owner)
  firstName    STRING  — first name
  lastName     STRING  — last name
  email        STRING  — owner email

── TABLE 3: hs_analytics.companies ──────────────────────────────
Always use FINAL: FROM hs_analytics.companies FINAL

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

FY27 5% cohort  → became_5_deal_date  >= '2026-04-01'
FY27 20% cohort → became_20_deal_date >= '2026-04-01'
FY26 5% cohort  → became_5_deal_date  >= '2025-04-01' AND < '2026-04-01'

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

-- AEs with most stalled deals:
SELECT
  concat(o.firstName,' ',o.lastName) AS owner,
  countDistinct(d.deal_id) AS stalled_deals,
  round(sum(d.amount)/1e6,1) AS at_risk_m
FROM hs_analytics.deals d FINAL
LEFT JOIN hs_analytics.owners o FINAL ON d.deal_owner = CAST(o.id AS VARCHAR)
WHERE d.pipeline = 'default'
  AND CASE WHEN d.deal_type IS NULL THEN 'Not Assigned' ELSE d.deal_type END NOT IN ('Partner-Led SMB')
  AND toInt64(d.deal_id) IN (SELECT DISTINCT toInt64(deal_id_hs) FROM kore_ai_hubspot.gs_deal_ids_hs)
  AND d.deal_stage IN ('20% - Solution','30% - Proof','40% - Proposal',
                       '60% - Price Negotiation','75% - Contract Review')
  AND d.became_5_deal_date >= '2026-04-01'
GROUP BY owner ORDER BY stalled_deals DESC LIMIT 10

-- Lost deals mentioning a specific competitor:
SELECT
  d.deal_name,
  concat(o.firstName,' ',o.lastName) AS owner,
  round(d.amount/1e6,2) AS amt_m,
  d.primary_closed_lost_reason,
  d.competitors,
  d.won_loss_notes
FROM hs_analytics.deals d FINAL
LEFT JOIN hs_analytics.owners o FINAL ON d.deal_owner = CAST(o.id AS VARCHAR)
WHERE d.pipeline = 'default'
  AND CASE WHEN d.deal_type IS NULL THEN 'Not Assigned' ELSE d.deal_type END NOT IN ('Partner-Led SMB')
  AND toInt64(d.deal_id) IN (SELECT DISTINCT toInt64(deal_id_hs) FROM kore_ai_hubspot.gs_deal_ids_hs)
  AND d.deal_stage = 'Closed Lost'
  AND d.competitors ILIKE '%salesforce%'
ORDER BY d.amount DESC LIMIT 20

-- Pipeline by industry at 20%+ stage (FY27):
SELECT
  CASE
    WHEN d.kore_primary_industry IN ('Financial Services','Banking','Insurance') THEN 'Financial Services'
    WHEN d.kore_primary_industry IN ('Manufacturing Discreet','Manufacturing Process','CPG') THEN 'Manufacturing'
    WHEN d.kore_primary_industry IN ('Hi-Tech','Telecom / Media / Entertainment') THEN 'TMT'
    ELSE 'Other'
  END AS industry,
  countDistinct(d.deal_id) AS deals,
  round(sum(d.amount)/1e6,1) AS pipeline_m
FROM hs_analytics.deals d FINAL
WHERE d.pipeline = 'default'
  AND CASE WHEN d.deal_type IS NULL THEN 'Not Assigned' ELSE d.deal_type END NOT IN ('Partner-Led SMB')
  AND toInt64(d.deal_id) IN (SELECT DISTINCT toInt64(deal_id_hs) FROM kore_ai_hubspot.gs_deal_ids_hs)
  AND d.became_20_deal_date >= '2026-04-01'
  AND d.became_20_deal_date <> '1900-01-01'
GROUP BY industry ORDER BY pipeline_m DESC

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

=================================================================
"""

# Tool definition for Claude
_QUERY_TOOL = {
    "name": "query_clickhouse",
    "description": (
        "Execute a SELECT query against the Kore.ai ClickHouse pipeline database. "
        "Use this for any question about deals, AEs, regions, industries, stages, "
        "win/loss data, competitors, BANT, attainment, or metrics not already in context."
    ),
    "input_schema": {
        "type": "object",
        "properties": {
            "sql": {
                "type": "string",
                "description": (
                    "A valid ClickHouse SELECT or WITH query following the schema rules above. "
                    "Always use FINAL, apply mandatory base filters, and LIMIT row queries."
                )
            }
        },
        "required": ["sql"]
    }
}


# =============================================================================
# Pydantic Models
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
    title: str = "Pipeline Intelligence Report"


# =============================================================================
# Core Chat Logic (shared by /chat and /export)
# =============================================================================

def _call_claude_with_tools(messages: list, max_tokens: int = 2000) -> str:
    """
    Call Claude with the query_clickhouse tool.
    Runs the tool loop (up to 4 rounds) and returns the final text reply.
    """
    response = _ai_client.messages.create(
        model=_CLAUDE_MODEL,
        system=_SYSTEM_PROMPT,
        messages=messages,
        tools=[_QUERY_TOOL],
        temperature=0,
        max_tokens=max_tokens,
    )

    rounds = 0
    while response.stop_reason == "tool_use" and rounds < 4:
        rounds += 1
        tool_block = next((b for b in response.content if b.type == "tool_use"), None)
        if not tool_block:
            break

        sql          = tool_block.input.get("sql", "")
        query_result = run_clickhouse_query(sql)

        print(f"  🔍 DB round {rounds}: {sql[:120]}...")
        print(f"  📥 Result preview: {query_result[:200]}")

        messages = messages + [
            {"role": "assistant", "content": response.content},
            {
                "role": "user",
                "content": [
                    {
                        "type": "tool_result",
                        "tool_use_id": tool_block.id,
                        "content": query_result,
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

    reply = next(
        (b.text for b in response.content if hasattr(b, "text") and b.text),
        "I could not generate a response. Please try rephrasing."
    )
    return reply


# =============================================================================
# Routes
# =============================================================================

@app.get("/", response_class=HTMLResponse)
def root():
    with open("chat.html", "r") as f:
        return HTMLResponse(content=f.read())


@app.get("/debug/clickhouse")
def debug_clickhouse():
    """
    Diagnostic endpoint — tests all 4 connection strategies.
    Open in browser: http://localhost:8000/debug/clickhouse
    """
    api_url   = os.getenv("CLICKHOUSE_API_URL", "")
    api_token = os.getenv("CLICKHOUSE_API_TOKEN", "")

    config = {
        "CLICKHOUSE_API_URL":   api_url or "NOT SET",
        "CLICKHOUSE_API_TOKEN": f"set ({len(api_token)} chars)" if api_token else "NOT SET",
    }

    if not api_url or not api_token:
        return {"config": config, "error": "Missing env vars — check .env file"}

    test_sql = "SELECT 1 AS ping FORMAT JSONCompact"
    results  = {}

    try:
        r = http_requests.post(api_url, data=test_sql.encode(),
            headers={"Authorization": f"Bearer {api_token}", "Content-Type": "text/plain"}, timeout=10)
        results["S1_POST_body"] = {"status": r.status_code, "body": r.text[:200]}
    except Exception as e:
        results["S1_POST_body"] = {"error": str(e)}

    try:
        r = http_requests.post(api_url, params={"query": test_sql},
            headers={"Authorization": f"Bearer {api_token}"}, timeout=10)
        results["S2_POST_query_param"] = {"status": r.status_code, "body": r.text[:200]}
    except Exception as e:
        results["S2_POST_query_param"] = {"error": str(e)}

    try:
        r = http_requests.get(api_url, params={"query": test_sql},
            headers={"Authorization": f"Bearer {api_token}"}, timeout=10)
        results["S3_GET_query_param"] = {"status": r.status_code, "body": r.text[:200]}
    except Exception as e:
        results["S3_GET_query_param"] = {"error": str(e)}

    try:
        r = http_requests.post(api_url, data="SELECT 1 AS ping".encode(),
            headers={"Authorization": f"Bearer {api_token}", "Content-Type": "text/plain",
                     "X-ClickHouse-Format": "JSONCompact"}, timeout=10)
        results["S4_POST_xheader"] = {"status": r.status_code, "body": r.text[:200]}
    except Exception as e:
        results["S4_POST_xheader"] = {"error": str(e)}

    working = [k for k, v in results.items() if isinstance(v, dict) and v.get("status", 999) < 400]
    return {
        "config": config,
        "working_strategies": working,
        "all_results": results,
        "recommendation": (
            f"Use strategy: {working[0]}" if working
            else "No strategy worked — check URL, token, network/firewall"
        )
    }


@app.post("/chat")
def chat(req: ChatRequest):

    try:
        print("QUESTION:", req.question)

        # Step 1
        sql_query = generate_sql(req.question)
        print("GENERATED SQL:")
        print(sql_query)

        # Step 2
        validate_sql(sql_query)

        # Step 3
        query_result = run_clickhouse_query(sql_query)
        print("CLICKHOUSE RESULT:")
        print(query_result)

        # Step 4
        final_response = generate_answer(
            question=req.question,
            sql=sql_query,
            data=query_result
        )

        return {
            "question": req.question,
            "sql": sql_query,
            "data": query_result,
            "answer": final_response
        }

    except Exception as e:
        print("ERROR OCCURRED:")
        print(str(e))

        return {
            "error": str(e)
        }


@app.post("/export/pdf")
def export_pdf(payload: ExportRequest):
    """
    Generate a multi-page PDF report from the current conversation.
    Claude synthesises a structured report; reportlab renders it.
    """
    conv_text = "\n\n".join(
        f"{'User' if m.role == 'user' else 'Assistant'}: {m.content}"
        for m in payload.conversation
    )

    export_prompt = (
        f"Based on this conversation, produce a structured pipeline intelligence report "
        f"titled '{payload.title}'. "
        "Format with these exact section headers (use ## for each):\n"
        "## Executive Summary\n"
        "## Pipeline Health\n"
        "## Key Metrics\n"
        "## Regional Breakdown\n"
        "## Risk & Opportunities\n"
        "## Recommended Actions\n\n"
        "Query the database for any missing data. Be data-driven with real numbers.\n\n"
        f"CONVERSATION CONTEXT:\n{conv_text}"
    )

    messages = [{"role": "user", "content": export_prompt}]
    print(f"📄 [export/pdf] Generating report: {payload.title}")
    try:
        report_text = _call_claude_with_tools(messages, max_tokens=3000)
    except Exception as exc:
        raise HTTPException(status_code=502, detail=f"Claude error: {exc}")

    pdf_bytes = _build_pdf(payload.title, report_text)
    filename  = payload.title.replace(" ", "_") + ".pdf"
    return StreamingResponse(
        io.BytesIO(pdf_bytes),
        media_type="application/pdf",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'}
    )


@app.post("/export/pptx")
def export_pptx(payload: ExportRequest):
    """
    Generate a branded PPTX presentation from the current conversation.
    """
    conv_text = "\n\n".join(
        f"{'User' if m.role == 'user' else 'Assistant'}: {m.content}"
        for m in payload.conversation
    )

    export_prompt = (
        f"Based on this conversation, produce slide content for a pipeline presentation "
        f"titled '{payload.title}'. "
        "Output each slide as:\n"
        "SLIDE: <Slide Title>\n"
        "BULLETS:\n- bullet 1\n- bullet 2\n...\n\n"
        "Include these slides: Title/Overview, Pipeline Health, Key Metrics, "
        "Regional Breakdown, Risk & Opportunities, Recommended Actions.\n"
        "Query the database for any missing data. Keep bullets crisp and data-driven.\n\n"
        f"CONVERSATION CONTEXT:\n{conv_text}"
    )

    messages = [{"role": "user", "content": export_prompt}]
    print(f"📊 [export/pptx] Generating deck: {payload.title}")
    try:
        slide_text = _call_claude_with_tools(messages, max_tokens=3000)
    except Exception as exc:
        raise HTTPException(status_code=502, detail=f"Claude error: {exc}")

    pptx_bytes = _build_pptx(payload.title, slide_text)
    filename   = payload.title.replace(" ", "_") + ".pptx"
    return StreamingResponse(
        io.BytesIO(pptx_bytes),
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'}
    )


# =============================================================================
# PDF Builder
# =============================================================================

C_NAVY_PDF  = colors.HexColor("#0D1B3E")
C_BLUE_PDF  = colors.HexColor("#1565C0")
C_WHITE_PDF = colors.white
C_BG_PDF    = colors.HexColor("#F7F9FC")
C_TXT_PDF   = colors.HexColor("#1E293B")
C_MID_PDF   = colors.HexColor("#475569")
C_DIM_PDF   = colors.HexColor("#94A3B8")
C_BORDER    = colors.HexColor("#CBD5E1")
C_ROW_ALT   = colors.HexColor("#EEF4FF")
C_GREEN_PDF = colors.HexColor("#2E7D32")
C_RED_PDF   = colors.HexColor("#B71C1C")
C_AMBER_PDF = colors.HexColor("#F59E0B")

SECTION_PALETTE = {
    "executive summary":    colors.HexColor("#0D1B3E"),
    "pipeline health":      colors.HexColor("#1565C0"),
    "key metrics":          colors.HexColor("#004D40"),
    "regional breakdown":   colors.HexColor("#BF360C"),
    "risk":                 colors.HexColor("#B71C1C"),
    "recommended actions":  colors.HexColor("#1B5E20"),
}

PW, PH = A4
ML = MR = 0.6 * inch
MT      = 0.45 * inch
MB      = 0.40 * inch
HDR_H   = 44
FTR_H   = 20
CW      = PW - ML - MR


def _strip_md(t: str) -> str:
    if not t:
        return ""
    t = re.sub(r'\*{1,3}(.*?)\*{1,3}', r'\1', t)
    t = re.sub(r'#{1,6}\s*', '', t)
    t = re.sub(r'^[\-\*•]\s*', '', t, flags=re.M)
    t = re.sub(r'`(.*?)`', r'\1', t)
    return t.strip()


def _make_styles():
    styles = getSampleStyleSheet()
    base = dict(fontName="Helvetica", textColor=C_TXT_PDF)

    defs = {
        "Cover_Title":   ParagraphStyle("Cover_Title",   fontSize=28, leading=34, textColor=C_WHITE_PDF, fontName="Helvetica-Bold",  spaceAfter=10),
        "Cover_Sub":     ParagraphStyle("Cover_Sub",     fontSize=14, leading=20, textColor=colors.HexColor("#B0BEC5"), fontName="Helvetica",       spaceAfter=6),
        "Section_H":     ParagraphStyle("Section_H",     fontSize=12, leading=16, textColor=C_WHITE_PDF, fontName="Helvetica-Bold",  spaceAfter=0),
        "Body":          ParagraphStyle("Body",          fontSize=9,  leading=14, textColor=C_TXT_PDF,   fontName="Helvetica",       spaceAfter=4, spaceBefore=2),
        "Bullet":        ParagraphStyle("Bullet",        fontSize=9,  leading=14, textColor=C_TXT_PDF,   fontName="Helvetica",       leftIndent=12, firstLineIndent=-8, spaceAfter=3),
        "H2":            ParagraphStyle("H2",            fontSize=11, leading=15, textColor=C_NAVY_PDF,  fontName="Helvetica-Bold",  spaceBefore=10, spaceAfter=4),
        "H3":            ParagraphStyle("H3",            fontSize=9,  leading=13, textColor=C_BLUE_PDF,  fontName="Helvetica-Bold",  spaceBefore=6,  spaceAfter=2),
        "Footer":        ParagraphStyle("Footer",        fontSize=7,  leading=9,  textColor=C_DIM_PDF,   fontName="Helvetica",       alignment=TA_CENTER),
    }
    return defs


class _HdrFtr(Flowable):
    """Canvas-level header/footer — paints outside frames."""
    def __init__(self, title, section_color):
        super().__init__()
        self._title = title
        self._color = section_color

    def draw(self):
        pass  # drawn in page templates below


def _build_pdf(title: str, report_text: str) -> bytes:
    buf    = io.BytesIO()
    styles = _make_styles()

    sections = _parse_report_sections(report_text)

    def _on_page(canvas, doc):
        canvas.saveState()
        # Header bar
        canvas.setFillColor(C_NAVY_PDF)
        canvas.rect(0, PH - HDR_H - MT, PW, HDR_H + MT, fill=1, stroke=0)
        canvas.setFillColor(C_WHITE_PDF)
        canvas.setFont("Helvetica-Bold", 10)
        canvas.drawString(ML, PH - MT - 28, title)
        # Footer
        canvas.setFillColor(C_BG_PDF)
        canvas.rect(0, 0, PW, FTR_H + MB, fill=1, stroke=0)
        canvas.setFillColor(C_DIM_PDF)
        canvas.setFont("Helvetica", 7)
        footer = f"Pipeline Intelligence  |  AI-Generated  |  CONFIDENTIAL  |  {date.today().strftime('%B %Y')}"
        canvas.drawCentredString(PW / 2, MB + 5, footer)
        pg = canvas.getPageNumber()
        canvas.drawRightString(PW - MR, MB + 5, f"Page {pg}")
        canvas.restoreState()

    frame = Frame(ML, MB + FTR_H, CW, PH - HDR_H - MT - MB - FTR_H, id="main")
    template = PageTemplate(id="main", frames=[frame], onPage=_on_page)
    doc = BaseDocTemplate(buf, pagesize=A4, leftMargin=ML, rightMargin=MR,
                          topMargin=MT + HDR_H, bottomMargin=MB + FTR_H)
    doc.addPageTemplates([template])

    story = []

    # Cover
    story.append(Spacer(1, 1.2 * inch))
    story.append(Paragraph(title, styles["Cover_Title"]))
    story.append(Paragraph(f"Generated {date.today().strftime('%B %d, %Y')}", styles["Cover_Sub"]))
    story.append(PageBreak())

    for sec_title, sec_body in sections:
        color_key = next((k for k in SECTION_PALETTE if k in sec_title.lower()), None)
        bar_color = SECTION_PALETTE.get(color_key, C_BLUE_PDF)

        # Section header bar
        story.append(Table(
            [[Paragraph(sec_title.upper(), styles["Section_H"])]],
            colWidths=[CW],
            style=TableStyle([
                ("BACKGROUND", (0, 0), (-1, -1), bar_color),
                ("TOPPADDING",    (0, 0), (-1, -1), 8),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
                ("LEFTPADDING",   (0, 0), (-1, -1), 10),
                ("RIGHTPADDING",  (0, 0), (-1, -1), 10),
            ])
        ))
        story.append(Spacer(1, 6))

        for line in sec_body.split("\n"):
            line = line.strip()
            if not line:
                story.append(Spacer(1, 3))
                continue
            if line.startswith("### "):
                story.append(Paragraph(_strip_md(line), styles["H3"]))
            elif line.startswith("## "):
                story.append(Paragraph(_strip_md(line), styles["H2"]))
            elif line.startswith(("- ", "* ", "• ")):
                story.append(Paragraph("• " + _strip_md(line[2:]), styles["Bullet"]))
            else:
                story.append(Paragraph(_strip_md(line), styles["Body"]))

        story.append(Spacer(1, 12))
        story.append(PageBreak())

    doc.build(story)
    return buf.getvalue()


def _parse_report_sections(text: str):
    """Split report text by ## headers into (title, body) pairs."""
    parts = re.split(r'^##\s+', text, flags=re.MULTILINE)
    sections = []
    for part in parts:
        if not part.strip():
            continue
        lines = part.strip().split("\n", 1)
        sec_title = lines[0].strip()
        sec_body  = lines[1].strip() if len(lines) > 1 else ""
        sections.append((sec_title, sec_body))
    return sections


def _parse_slides(text: str):
    """Parse SLIDE:/BULLETS: formatted text into list of (title, bullets)."""
    slides = []
    current_title   = None
    current_bullets = []
    for line in text.split("\n"):
        line = line.rstrip()
        if line.startswith("SLIDE:"):
            if current_title is not None:
                slides.append((current_title, current_bullets))
            current_title   = line[6:].strip()
            current_bullets = []
        elif line.startswith("- ") and current_title:
            current_bullets.append(line[2:].strip())
        elif line.startswith("BULLETS:"):
            continue
    if current_title is not None:
        slides.append((current_title, current_bullets))
    return slides


# =============================================================================
# PPTX Builder
# =============================================================================

C_NAVY_P  = RGBColor(0x0D, 0x1B, 0x3E)
C_DNAV_P  = RGBColor(0x0A, 0x11, 0x28)
C_BLUE_P  = RGBColor(0x1E, 0x88, 0xE5)
C_WHITE_P = RGBColor(0xFF, 0xFF, 0xFF)
C_LTBG_P  = RGBColor(0xF5, 0xF7, 0xFA)
C_TXT_P   = RGBColor(0x1A, 0x1A, 0x2E)
C_GOLD_P  = RGBColor(0xF5, 0xA6, 0x23)
C_DIM_P   = RGBColor(0x88, 0x99, 0xAA)

SLIDE_SECTION_COLORS = {
    "overview":     RGBColor(0x1E, 0x88, 0xE5),
    "pipeline":     RGBColor(0x00, 0x89, 0x7B),
    "metric":       RGBColor(0x2E, 0x7D, 0x32),
    "regional":     RGBColor(0xBF, 0x36, 0x0C),
    "risk":         RGBColor(0xC6, 0x28, 0x28),
    "opportunit":   RGBColor(0xC6, 0x28, 0x28),
    "recommend":    RGBColor(0x1B, 0x5E, 0x20),
    "action":       RGBColor(0x1B, 0x5E, 0x20),
}

SLIDE_W = Inches(13.33)
SLIDE_H = Inches(7.5)


def _pptx_bg(slide, color: RGBColor):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color


def _pptx_rect(slide, l, t, w, h, color: RGBColor):
    shp = slide.shapes.add_shape(1, Inches(l), Inches(t), Inches(w), Inches(h))
    shp.fill.solid()
    shp.fill.fore_color.rgb = color
    shp.line.fill.background()
    return shp


def _pptx_txt(slide, text, l, t, w, h, bold=False, size=18, color=None, align=PP_ALIGN.LEFT, wrap=True):
    txb = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    txb.word_wrap = wrap
    tf  = txb.text_frame
    tf.word_wrap = wrap
    p   = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color or C_TXT_P
    return txb


def _build_pptx(title: str, slide_text: str) -> bytes:
    slides_data = _parse_slides(slide_text)
    if not slides_data:
        # Fallback: treat whole text as one slide
        slides_data = [(title, [slide_text[:500]])]

    prs = Presentation()
    prs.slide_width  = SLIDE_W
    prs.slide_height = SLIDE_H

    footer_text = (
        f"Pipeline Intelligence  |  AI-Generated  |  CONFIDENTIAL  |  "
        f"{date.today().strftime('%B %Y')}"
    )

    def _add_footer(slide):
        _pptx_rect(slide, 0, 7.1, 13.33, 0.4, C_DNAV_P)
        _pptx_txt(slide, footer_text, 0.3, 7.12, 12, 0.35,
                  size=7, color=C_DIM_P, align=PP_ALIGN.CENTER)

    def _slide_accent(slide_title_lower):
        for k, c in SLIDE_SECTION_COLORS.items():
            if k in slide_title_lower:
                return c
        return C_BLUE_P

    # ── Cover slide ──────────────────────────────────────────────────────────
    cover_layout = prs.slide_layouts[6]  # blank
    cover = prs.slides.add_slide(cover_layout)
    _pptx_bg(cover, C_NAVY_P)
    _pptx_rect(cover, 0, 0, 13.33, 7.5, C_NAVY_P)
    _pptx_rect(cover, 0, 3.2, 13.33, 0.06, C_BLUE_P)
    _pptx_txt(cover, title, 0.8, 1.6, 11.5, 1.4,
              bold=True, size=36, color=C_WHITE_P)
    _pptx_txt(cover, "Pipeline Intelligence Report", 0.8, 3.0, 8, 0.6,
              size=16, color=RGBColor(0xB0, 0xBE, 0xC5))
    _pptx_txt(cover, f"Generated: {date.today().strftime('%B %d, %Y')}",
              0.8, 3.6, 6, 0.45, size=12, color=RGBColor(0x78, 0x90, 0x9C))
    _pptx_txt(cover, "CONFIDENTIAL", 0.8, 6.8, 4, 0.4,
              size=9, color=RGBColor(0xEF, 0x53, 0x50))

    # ── Content slides ────────────────────────────────────────────────────────
    blank_layout = prs.slide_layouts[6]

    for i, (s_title, bullets) in enumerate(slides_data):
        slide = prs.slides.add_slide(blank_layout)
        _pptx_bg(slide, C_LTBG_P)

        accent = _slide_accent(s_title.lower())

        # Top accent bar
        _pptx_rect(slide, 0, 0, 13.33, 0.9, accent)
        _pptx_txt(slide, s_title.upper(), 0.35, 0.1, 12.5, 0.7,
                  bold=True, size=18, color=C_WHITE_P)

        # Slide number
        _pptx_txt(slide, str(i + 1), 12.5, 0.12, 0.6, 0.6,
                  size=11, color=C_WHITE_P, align=PP_ALIGN.RIGHT)

        # Content area background
        _pptx_rect(slide, 0.3, 1.0, 12.73, 5.9, C_WHITE_P)

        # Bullets
        if bullets:
            txb = slide.shapes.add_textbox(
                Inches(0.5), Inches(1.1), Inches(12.3), Inches(5.6)
            )
            txb.word_wrap = True
            tf = txb.text_frame
            tf.word_wrap = True

            for j, bullet in enumerate(bullets[:12]):
                p = tf.add_paragraph() if j > 0 else tf.paragraphs[0]
                p.space_before = Pt(4)
                p.space_after  = Pt(2)

                # Bullet dot
                dot = p.add_run()
                dot.text = "●  "
                dot.font.size  = Pt(8)
                dot.font.color.rgb = accent
                dot.font.bold  = False

                run = p.add_run()
                run.text = bullet
                run.font.size  = Pt(12)
                run.font.color.rgb = C_TXT_P
                run.font.bold  = False
        else:
            _pptx_txt(slide, "No data available.", 0.5, 1.2, 12, 0.5,
                      size=11, color=C_DIM_P)

        _add_footer(slide)

    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()
