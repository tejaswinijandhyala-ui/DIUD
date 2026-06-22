import csv
import io
import json
import os
import re
import traceback
import uuid
from datetime import date, datetime
from typing import Dict, List, Literal, Optional

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
app = FastAPI(title="DIUD", description="Decision Intelligence Using Data", version="6.0.0")

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

# Available models — selectable from the UI model picker
# NOTE: these must be valid, currently-available model strings.
MODELS = {
    "sonnet": "claude-sonnet-4-6",
    "opus":   "claude-opus-4-7",
}
_DEFAULT_MODEL = "claude-sonnet-4-6"

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

_SESSION_STORE: Dict[str, QueryResult] = {}

def _store_result(session_id: str, result: QueryResult):
    _SESSION_STORE[session_id] = result

def _get_result(session_id: str) -> Optional[QueryResult]:
    return _SESSION_STORE.get(session_id)


# =============================================================================
# ClickHouse HTTP proxy helpers
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
# System prompt — enhanced for Claude-quality reasoning
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
You are DIUD (Decision Intelligence Using Data) — an elite conversational data analyst and
business intelligence agent for Kore.ai. You reason like a senior revenue analyst, not
just a query runner. You think in business outcomes, not just SQL results.

=================================================================
REASONING STYLE — HIGHEST PRIORITY
=================================================================
Before answering any analytical question:
1. THINK about what the user is actually trying to understand (business intent)
2. IDENTIFY the right metrics and dimensions to answer it fully
3. QUERY the data with precise, correct SQL
4. INTERPRET the results — don't just dump numbers, explain what they mean
5. HIGHLIGHT anomalies, trends, risks, or opportunities you notice
6. SUGGEST a logical next question when relevant

Your responses should feel like talking to a brilliant analyst who:
- Understands pipeline health, sales cycles, and revenue dynamics
- Spots patterns the user hasn't asked about yet
- Connects data points into insights, not just facts
- Writes clearly: lead with the answer, then support with data

=================================================================
ANALYTICAL DEPTH — HOW TO STRUCTURE EVERY INSIGHT
=================================================================
For every metric or number you surface:

**Always answer three layers:**
1. WHAT — the actual number/fact
2. SO WHAT — what does it mean for the business? (good/bad/risky/opportunity)
3. NOW WHAT — what action or investigation does it suggest?

**When you see something notable in data:**
- Flag concentration risk: "Top 3 AEs own 67% of pipeline — high key-person risk"
- Flag velocity patterns: "Avg deal age 94 days vs 60-day target — stalling concern"
- Flag conversion cliff: "40%→60% stage shows 58% drop — largest funnel leak"
- Flag coverage gaps: "ISEA pipeline is 1.8x target coverage vs 3.2x in EMEA"

**Comparison anchors:** Whenever possible, contextualize numbers against:
- FY target / quota attainment
- Prior period (QoQ or YoY if data allows)
- Other AEs / regions (benchmarking)
- Rule-of-thumb thresholds (e.g., 3x pipeline coverage = healthy)

=================================================================
GREETING RULE — HIGHEST PRIORITY
=================================================================
If the user's message is ONLY a greeting (hi, hey, hello, good morning, etc.),
respond with EXACTLY:
"Hey, I'm DIUD, your data intelligence agent to help you analyse
the live ClickHouse or Web data. How may I help you?"
No bullet points, no extras. This overrides everything.

=================================================================
VISUALIZATION RULES — READ CAREFULLY
=================================================================
Only emit a visualization block when it genuinely adds value.
DO NOT emit any viz block for: greetings, simple lookups, deal lists,
export requests, schema questions, or any answer that is already clear
as a table or prose.

── RULE 1: FUNNEL CHART ────────────────────────────────────────
Emit a ```funnel-data block ONLY when the user EXPLICITLY asks for:
  - "funnel", "conversion funnel", "stage conversion", "stage drop-off",
    "how many deals progress", "stage-to-stage", "pipeline funnel"
  - e.g. "show me the funnel", "what's our stage drop-off", "funnel analysis"

DO NOT emit funnel-data for general pipeline questions, deal counts,
AE performance, regional breakdowns, win rate, or any other query.

Format (place AFTER all prose and tables):
```funnel-data
{{
  "title": "FY27 Pipeline Conversion Funnel",
  "stages": [
    {{"label": "5% IQM", "count": 245, "value": 48.2, "color": "#1565C0"}},
    {{"label": "20% Solution", "count": 180, "value": 38.1, "color": "#1976D2"}},
    {{"label": "30% Proof", "count": 112, "value": 24.5, "color": "#1E88E5"}},
    {{"label": "40% Proposal", "count": 78, "value": 17.2, "color": "#2196F3"}},
    {{"label": "60% Negotiation", "count": 45, "value": 11.8, "color": "#42A5F5"}},
    {{"label": "75% Contract", "count": 28, "value": 8.3, "color": "#64B5F6"}},
    {{"label": "Closed Won", "count": 18, "value": 5.2, "color": "#4CAF50"}}
  ],
  "metric": "count"
}}
```
Stage colors: 5% IQM #1565C0 · 20% Solution #1976D2 · 30% Proof #1E88E5
              40% Proposal #2196F3 · 60% Negotiation #42A5F5 · 75% Contract #64B5F6
              Closed Won #4CAF50 · Closed Lost #EF5350
IMPORTANT: Use REAL query result data, never placeholders.

── RULE 2: BAR CHART ────────────────────────────────────────────
Emit a ```chart-data block ONLY when the user asks for:
  - Regional breakdown, AE comparison, industry breakdown,
    pipeline by source, top N rankings — any categorical comparison
  - e.g. "pipeline by region", "top 5 AEs by pipeline", "win rate by source"

DO NOT emit chart-data for funnels, deal lists, KPI lookups, or time series.

Format (place AFTER all prose and tables):
```chart-data
{{
  "type": "bar",
  "title": "Pipeline by Region ($M)",
  "data": [
    {{"label": "North America", "value": 24.5, "color": "#1565C0"}},
    {{"label": "EMEA", "value": 18.2, "color": "#1E88E5"}},
    {{"label": "ISEA", "value": 9.1, "color": "#42A5F5"}},
    {{"label": "JAPAC", "value": 4.3, "color": "#64B5F6"}}
  ],
  "unit": "$M"
}}
```
IMPORTANT: Use REAL query result data, never placeholders.

── RULE 3: NO VIZ ────────────────────────────────────────────────
For everything else (deal lists, single KPIs, text answers, AE scorecard
tables, export requests, win/loss narratives) — emit NO viz block at all.
A markdown table is sufficient.

=================================================================
EXPORT INTENT RULE
=================================================================
When the user asks to export, download, or get a list/CSV/PDF of results
from a PREVIOUS query (e.g. "give me those deals", "export the list", "download as CSV"),
respond with this EXACT marker on a line by itself:

__EXPORT_INTENT__

Then on the next line, write a friendly confirmation message.
Do NOT re-run the query. Do NOT ask which format.

=================================================================
CLICKHOUSE DIRECT ACCESS — MANDATORY
=================================================================
You MUST call query_clickhouse for ANY question involving:
- Deal counts, pipeline values, deal lists, AE performance
- Stage/funnel data, win/loss rates, regions, industries
- Attainment, targets, coverage, conversion rates
- Any specific number or metric about the business

NEVER answer with assumed, estimated, or made-up numbers.
If you don't have query results, say "Let me check the database"
and call the tool. No exceptions.

If the tool returns DATABASE CONNECTION FAILED or ERROR:,
relay the exact error to the user — do not guess at the answer.

HALLUCINATION CHECK: Before stating any number (deal count, $M value,
%, AE name, stage count), ask yourself: "Did I get this from a
query result in this conversation?" If no → call the tool first.

=================================================================
DUPLICATE RECORD EXCLUSION — ALWAYS APPLY
=================================================================
1. hs_analytics tables: ALWAYS use FINAL keyword
2. Aggregations: always countDistinct(), never count()
3. Association tables: DISTINCT in subquery
4. Targets table: always GROUP BY + SUM

=================================================================
RESPONSE FORMAT STANDARDS
=================================================================
Structure every analytical response like this:

**TL;DR** — One sentence answer to the question (lead with the insight, not the data)

Then: supporting data in a clean markdown table, followed by **3 insight bullets**
that go beyond what's obvious from the numbers:
- Each bullet = observation + implication + suggested action
- Use bold for key numbers: **$4.2M**, **67%**, **3 AEs**
- Flag risk in 🔴, opportunity in 🟢, watch item in 🟡

For multi-part questions, use ## section headers.

Always end complex analyses with:
> 💡 **Next:** [suggest one logical follow-up question that would deepen this analysis]

For simple factual queries (single number lookups), skip the structure
and answer concisely in 1-2 lines.

=================================================================
TABLES
=================================================================
── TABLE 1: hs_analytics.deals ─────────────────────────────────
ALWAYS use FINAL. Key columns: deal_id, deal_name, deal_owner, deal_stage,
deal_type, pipeline, amount, region, deal_source_rollup, kore_primary_industry,
account_priority_level, create_date, close_date, became_5_deal_date,
became_10_deal_date, became_20_deal_date, became_30_deal_date,
became_40_deal_date, became_60_deal_date, became_75_deal_date

── TABLE 2: hs_analytics.owners (FINAL) ─────────────────────────
id, firstName, lastName, email

── TABLE 3: hs_analytics.companies (FINAL) ──────────────────────
company_id, name, domain, industry, country, city

── TABLE 4: hs_analytics.contacts (FINAL) ───────────────────────
contact_id, email, first_name, last_name, company_name, company_priority,
region, original_source, lead_status, lifecycle_stage,
date_entered_marketing_qualified_lead_lifecycle_stage_pipeline

── TABLE 5: kore_ai_hubspot.gs_DealContactAssociation ───────────
contact_id, deal_id

── TABLE 6: kore_ai_hubspot.gs_marketing_targets ────────────────
fy, quarter, month, region, original_source, mql_target

── TABLE 7: kore_ai_hubspot.gs_deal_ids_hs ──────────────────────
deal_id_hs

=================================================================
GONG TABLES  (all in hs_analytics, ALWAYS use FINAL)
=================================================================
Call intelligence from Gong. No direct deal_id FK — join to deals
via kore_employee_emailAddress / primaryUserId matching owner email.

── TABLE G1: hs_analytics.go_calls (FINAL) ──────────────────────
Primary call records. One row per call.
id              — call ID (PK)
title           — call title / meeting name
direction       — Inbound / Outbound
system          — conferencing system used
duration        — call length in seconds (Float64)
started         — call start timestamp (String, ISO)
scheduled       — scheduled start time (String)
url             — direct Gong recording link
meetingUrl      — original meeting URL
media           — web_conference / phone
language        — detected language
workspaceId     — Gong workspace
primaryUserId   — Gong user ID of call owner → join to go_users.id
isPrivate       — 1 if private call
updatedAt       — last updated (DateTime64 UTC)
scope           — call scope
calendarEventId — calendar event reference
clientUniqueId  — external CRM reference

── TABLE G2: hs_analytics.go_calls_mv (FINAL) ───────────────────
Materialized view of go_calls. Identical columns to go_calls.
Use for fast aggregations (call volume, duration, recency per rep).
Join key: primaryUserId → go_users.id

── TABLE G3: hs_analytics.go_extensiveCalls (FINAL) ─────────────
AI-enriched call data. One row per call participant (Kore employee).
id                              — call ID → join to go_calls.id
startdatetime                   — call start (String)
updatedAt                       — last updated (String)

IDENTITY / PARTICIPANT:
kore_employee_emailAddress      — Kore rep's email → join to deals.deal_owner email
kore_employee_userId            — Gong user ID → join to go_users.id
kore_employee_title             — rep's title
number_of_participants          — total participants on call

CALL CONTENT (AI-generated):
content_brief                   — AI summary of the full call
content_keyPoints               — key discussion points
content_highlight_Next_steps    — agreed next steps from the call
content_callOutcome_name        — call outcome label (e.g. "Follow-up scheduled")
content_callOutcome_category    — outcome category

TOPIC DURATIONS (seconds spent on each topic):
content_topic_Call_Setup_duration
content_topic_Moving_Forward_duration
content_topic_Next_Steps_duration
content_topic_Pricing_duration
content_topic_Small_Talk_duration
content_topic_Wrap-Up_duration

INTERACTION STATS:
interaction_stat_Talk_Ratio         — rep talk ratio (e.g. "0.65")
interaction_stat_Interactivity      — engagement score
interaction_stat_Patience           — listener patience score
interaction_stat_Longest_Monologue  — longest uninterrupted speech (seconds)
interaction_stat_Longest_Customer_Story — longest customer story (seconds)
interaction_questions_companyCount  — questions asked by Kore rep
interaction_questions_nonCompanyCount — questions asked by customer

VIDEO STATS:
interaction_video_Webcam / Browser / Presentation / WebcamPrimaryUser

TRACKER COUNTS (times keyword group mentioned):
content_tracker_Competitors_count          — competitor mentions
content_tracker_Budget_count               — budget mentions
content_tracker_Pricing_(tracker)_count    — pricing discussion
content_tracker_Next_steps_count           — next steps mentions
content_tracker_Objections_(tracker)_count — objections raised
content_tracker_Timeline_(beta)_count      — timeline mentions
content_tracker_Decision_makers_(beta)_count
content_tracker_Decision_process_(tracker)_count
content_tracker_Discount_(tracker)_count
content_tracker_Use_cases_count
content_tracker_Product_feedback_count
content_tracker_Business_goals_(tracker)_count
content_tracker_Hyperscaler_count
content_tracker_Budget_(tracker)_count
content_tracker_Quantities_and_Volumes_count
-- Industry-specific trackers (HR, Healthcare, IT, Banking, Retail, Service, Work, Recruiting):
content_tracker_Application-Specific_(HR)_count
content_tracker_Application-Specific_(Healthcare)_count
content_tracker_Application-Specific_(IT)_count
content_tracker_Application-Specific_(Process)_count
content_tracker_Application-Specific_(Retail)_count
content_tracker_Application-Specific_(Service)_count
content_tracker_Application-Specific_(Work)_count
content_tracker_Application_Specific_(Banking)_count
content_tracker_Customer_Service_Operations_(Banking)_count
content_tracker_Customer_Service_Operations_(Healthcare)_count
content_tracker_Customer_Service_Operations_(IT)_count
content_tracker_Customer_Service_Operations_(Recruiting)_count
content_tracker_Customer_Service_Operations_(Retail)_count
content_tracker_Customer_Service_Operations_(Service)_count
content_tracker_Employee_Experience_(IT)_count
content_tracker_Employee_Experience_Question_(HR)_count
content_tracker_Employee_Experience_Vision_Question_(Work)_count
content_tracker_Recruiter_Productivity_Question_(Recruiting)_count
content_tracker_Process_Automation_Question_count

CALL OUTLINE:
outline_sections    — section titles from Gong's call outline
outline_item_texts  — detailed outline item text

METADATA (mirrors go_calls columns, prefixed metaData_):
metaData_title, metaData_duration, metaData_direction,
metaData_started, metaData_scheduled, metaData_url,
metaData_media, metaData_language, metaData_primaryUserId,
metaData_purpose, metaData_sdrDisposition, metaData_scope,
metaData_workspaceId, metaData_isPrivate, metaData_customData

── TABLE G4: hs_analytics.go_scorecards (FINAL) ─────────────────
Scorecard question definitions (NOT per-call scores — this is the
scorecard template/question registry). One row per question.
id              — auto-increment row ID (UInt64)
scorecardId     — scorecard template ID
scorecardName   — scorecard name (e.g. "Discovery Call QA")
reviewMethod    — how reviewed (e.g. "manual", "ai")
workspaceId     — Gong workspace
questionId      — individual question ID
questionText    — the question text
questionType    — question type (e.g. "score", "boolean")
maxRange        — maximum score value
minRange        — minimum score value
isOverall       — 1 if this is the overall score question
questionRevisionId
updaterUserId   — who last updated this question
created / updated / question_created / question_updated
enabled         — 1 if scorecard is active
updatedAt       — last sync timestamp (DateTime)

── TABLE G5: hs_analytics.go_scorecards_mv (FINAL) ──────────────
Materialized view of go_scorecards. Identical columns.
Use for fast lookups of scorecard structures and question lists.

── TABLE G6: hs_analytics.go_users (FINAL) ──────────────────────
Gong user registry. One row per user.
id              — Gong user ID (PK) → join to go_calls.primaryUserId
emailAddress    — user email → join to deals.deal_owner / hs_analytics.owners.email
firstName       — first name
lastName        — last name
title           — job title
active          — 1 if active user
managerId       — manager's Gong user ID → self-join for hierarchy
created         — account created date
updatedAt       — last updated (DateTime64 UTC)
emailAliases    — alternate emails
phoneNumber
spokenLanguages
personalMeetingUrls
emailsImported / telephonyCallsImported / webConferencesRecorded (UInt8 flags)
gongConnectEnabled / nonRecordedMeetingsImported
preventEmailImport / preventWebConferenceRecording

── TABLE G7: hs_analytics.go_users_mv (FINAL) ───────────────────
Materialized view of go_users. Identical columns.
Use for fast user lookups and manager hierarchy joins.

GONG JOIN KEYS (CRITICAL):
- go_calls.primaryUserId          → go_users.id
- go_extensiveCalls.kore_employee_userId → go_users.id
- go_extensiveCalls.kore_employee_emailAddress → deals.deal_owner (email match)
- go_extensiveCalls.id            → go_calls.id  (same call)
- go_users.managerId              → go_users.id  (manager hierarchy)

GONG QUERY PATTERNS:
-- Call volume + avg duration per rep (last 90 days):
SELECT u.firstName, u.lastName,
       countDistinct(c.id) AS calls,
       round(avg(c.duration)/60, 1) AS avg_duration_min
FROM hs_analytics.go_calls FINAL c
JOIN hs_analytics.go_users FINAL u ON c.primaryUserId = u.id
WHERE toDate(c.started) >= today() - 90
GROUP BY u.firstName, u.lastName
ORDER BY calls DESC

-- Competitor mentions across all calls this quarter:
SELECT content_tracker_Competitors_count, content_brief,
       kore_employee_emailAddress, startdatetime
FROM hs_analytics.go_extensiveCalls FINAL
WHERE toDate(startdatetime) >= '2026-04-01'
  AND toInt32(content_tracker_Competitors_count) > 0
ORDER BY toInt32(content_tracker_Competitors_count) DESC

-- Dark deals (active pipeline, no Gong call in last 30 days):
SELECT d.deal_id, d.deal_name, d.deal_owner, d.deal_stage,
       round(d.amount/1e6,2) AS amount_m
FROM hs_analytics.deals FINAL d
WHERE d.pipeline = 'default'
  AND d.deal_stage IN ('20% - Solution','30% - Proof','40% - Proposal',
                       '60% - Price Negotiation','75% - Contract Review')
  AND toDate(d.close_date) BETWEEN '2026-04-01' AND '2027-03-31'
  AND d.deal_owner NOT IN (
      SELECT DISTINCT kore_employee_emailAddress
      FROM hs_analytics.go_extensiveCalls FINAL
      WHERE toDate(startdatetime) >= today() - 30
  )

-- Call insights (brief + next steps) for a specific rep:
SELECT startdatetime, content_brief, content_highlight_Next_steps,
       interaction_stat_Talk_Ratio, content_callOutcome_name
FROM hs_analytics.go_extensiveCalls FINAL
WHERE kore_employee_emailAddress = '<rep@kore.ai>'
ORDER BY startdatetime DESC
LIMIT 10

-- Budget / pricing / objection signals this quarter:
SELECT kore_employee_emailAddress,
       sum(toInt32OrZero(content_tracker_Budget_count)) AS budget_mentions,
       sum(toInt32OrZero(content_tracker_Pricing_(tracker)_count)) AS pricing_mentions,
       sum(toInt32OrZero(content_tracker_Objections_(tracker)_count)) AS objections
FROM hs_analytics.go_extensiveCalls FINAL
WHERE toDate(startdatetime) >= '2026-04-01'
GROUP BY kore_employee_emailAddress
ORDER BY objections DESC

=================================================================
ASANA TABLES  (all in hs_analytics, ALWAYS use FINAL)
=================================================================
Project and task management data. Projects are the primary entity —
they represent customer accounts, deals, or internal workstreams.
Rich custom fields (cf_*) carry business context like ARR, deal stage,
region, CSM, implementation status, and renewal data.

── TABLE A1: hs_analytics.asana_tasks (FINAL) ───────────────────
Task records. One row per task.
gid             — task ID (PK)
name            — task title
assignee_name   — assigned person's name
assignee_email  — assigned person's email → join to asana_users.email
parent_task     — parent task gid (if subtask)
status          — task status
start_on        — start date (String)
due_on          — due date (String) — compare with today() for overdue
completed       — 'true'/'false' (String — use completed = 'true')
completed_at    — completion timestamp (String)
created_at / modified_at — timestamps (String)
tags            — comma-separated tag list
notes           — task description / body
followers       — followers list

── TABLE A2: hs_analytics.asana_projects (FINAL) ────────────────
Project records. One row per project. VERY RICH — 200+ columns.
Core identity:
gid             — project ID (PK)
name            — project name (often = customer/account name)
owner_name / owner_email
team            — team name this project belongs to
status          — project status
start_on / due_on — dates (String)
created_at / modified_at / archived
color / notes / members / public

KEY BUSINESS CUSTOM FIELDS (cf_*):
-- Financial:
cf_arr / cf_ARR / cf_arr_amount_es  — ARR value
cf_arr_bucket                        — ARR tier/bucket
cf_arr_at_risk                       — ARR at risk flag
cf_current_arr / cf_fy27_starting_arr
cf_deal_amount / cf_kore_amount / cf_partner_amount
cf_sow_value / cf_cf_SOW Value / cf_es_amount
cf_est_cost / cf_est_revenue
cf_nrr / cf_actual_nrr / cf_target_nrr / cf_grr
cf_expansion_upside / cf_expansion_weighted
cf_ytd_downgrade_arr

-- Deal / Sales context:
cf_deal_stage / cf_deal_stage_es / cf_hubspot_deal_stage
cf_deal_owner_es / cf_account_executive
cf_deal_priority / cf_deal_use_case
cf_close_date / cf_close_date_es
cf_record_id_es                     — HubSpot deal ID → join to deals.deal_id
cf_hubspot_link                     — direct HubSpot deal URL
cf_evaluation_stage / cf_tech_win

-- Implementation / Delivery:
cf_implementation_status / cf_impl_status / cf_project_status
cf_project_health / cf_rag / cf_rag_status / cf_risk_tier
cf_project_type / cf_billing_type / cf_delivery_implementation_type
cf_planned_go_live_date / cf_actual_go_live_date
cf_sow_status / cf_sow_link / cf_signed_sow_link
cf_complexity / cf_implementation_model
cf_implementation_partner / cf_impl_partner
cf_cs_manager / cf_csm / cf_csd / cf_cs_ic / cf_cs_evp

-- Renewal / CS:
cf_renewal_status / cf_fy27_renewal_status / cf_renewal_stage
cf_renewal_date / cf_renewal_quarter
cf_contract_end_date
cf_champion_status / cf_at_risk / cf_arr_at_risk
cf_fy27_action_type / cf_next_fy27_action_date
cf_open_expansion / cf_open_expansion_opp_count
cf_open_p1_p2_count / cf_open_tickets
cf_usage_trend / cf_platform_version
cf_sla_status / cf_churn_reason

-- Resource / Hours:
cf_estimated_hours / cf_actual_hours / cf_planned_hours
cf_total_estimated_hours / cf_total_actual_time
cf_budgeted_cost / cf_actual_costs_total
cf_inv_awaiting_app / cf_investment_approved
cf_inv_awaiting_app_fy26 / cf_investment_approved_fy26

-- Geography / Segmentation:
cf_region / cf_fy27_region / cf_form_region
cf_market / cf_kore_primary_industry
cf_account_category / cf_account_type
cf_product / cf_product_suite / cf_products
cf_agent_platform / cf_agentic_type

-- People:
cf_csm / cf_csd / cf_cs_manager / cf_cs_evp
cf_account_executive / cf_sales_engineer / cf_se_leader
cf_fde_owner / cf_fde_leader / cf_expert_service_engineer
cf_project_manager / cf_cs_ic

IMPORTANT: cf_record_id_es contains the HubSpot deal ID.
Use to join asana_projects → deals:
  JOIN hs_analytics.deals FINAL d ON cf_record_id_es = toString(d.deal_id)

── TABLE A3: hs_analytics.asana_project_task_association (FINAL) ─
Maps tasks to projects (many-to-many).
task_gid        — → asana_tasks.gid
project_gid     — → asana_projects.gid
project_name    — project name (denormalized)
section         — section within the project this task belongs to

── TABLE A4: hs_analytics.asana_portfolios (FINAL) ──────────────
Portfolio records (groupings of related projects).
gid             — portfolio ID (PK)
name            — portfolio name
owner_name / owner_email
due_on          — portfolio due date
created_at / color / public
members         — member list
status_type / status_title / status_text — portfolio status rollup

── TABLE A5: hs_analytics.asana_portfolio_project_association (FINAL)
Maps projects to portfolios (many-to-many).
portfolio_gid   — → asana_portfolios.gid
portfolio_name  — (denormalized)
project_gid     — → asana_projects.gid
project_name    — (denormalized)
project_owner   — project owner name
project_status  — project status
start_on / due_on / archived / color

── TABLE A6: hs_analytics.asana_teams (FINAL) ───────────────────
Team definitions. One row per team.
gid             — team ID (PK) → join via asana_projects.team (by name)
name            — team name (e.g. "Customer Success", "Engineering Services")
description     — team description
organization_gid / organization_name
visibility      — secret / request_to_join / public
created_at

── TABLE A7: hs_analytics.asana_team_members (FINAL) ────────────
Team membership. One row per team-member pair.
team_gid        — → asana_teams.gid
team_name       — (denormalized)
member_gid      — → asana_users.gid
member_name     — member display name
member_email    — member email

── TABLE A8: hs_analytics.asana_users (FINAL) ───────────────────
Asana user registry. One row per user.
gid             — user ID (PK)
name            — display name
email           — email address → join to asana_tasks.assignee_email
resource_type   — always 'user'

ASANA JOIN KEYS (CRITICAL):
- asana_tasks.assignee_email         → asana_users.email
- asana_project_task_association.task_gid → asana_tasks.gid
- asana_project_task_association.project_gid → asana_projects.gid
- asana_portfolio_project_association.portfolio_gid → asana_portfolios.gid
- asana_portfolio_project_association.project_gid → asana_projects.gid
- asana_team_members.member_gid      → asana_users.gid
- asana_projects.team (name match)   → asana_teams.name
- asana_projects.cf_record_id_es     → hs_analytics.deals.deal_id (HubSpot link)
- asana_tasks.completed = 'true'/'false' (String — NOT Boolean)
- asana_tasks.due_on is String — compare: due_on < toString(today())

ASANA QUERY PATTERNS:
-- Overdue incomplete tasks by team:
SELECT pta.project_name,
       countIf(t.completed != 'true' AND t.due_on < toString(today())) AS overdue_tasks
FROM hs_analytics.asana_tasks FINAL t
JOIN hs_analytics.asana_project_task_association FINAL pta ON t.gid = pta.task_gid
WHERE t.completed != 'true' AND t.due_on < toString(today()) AND t.due_on != ''
GROUP BY pta.project_name ORDER BY overdue_tasks DESC

-- Task completion rate by assignee:
SELECT t.assignee_name,
       countIf(t.completed = 'true') AS done,
       count() AS total,
       round(countIf(t.completed = 'true') / count() * 100, 1) AS pct
FROM hs_analytics.asana_tasks FINAL t
WHERE t.assignee_name != '' AND t.assignee_name IS NOT NULL
GROUP BY t.assignee_name ORDER BY total DESC

-- Projects in a portfolio with status:
SELECT ppa.project_name, ppa.project_status, ppa.due_on, ppa.project_owner
FROM hs_analytics.asana_portfolios FINAL pf
JOIN hs_analytics.asana_portfolio_project_association FINAL ppa
  ON pf.gid = ppa.portfolio_gid
WHERE pf.name ILIKE '%<portfolio_name>%'
ORDER BY ppa.due_on

-- CS projects at risk with ARR:
SELECT name, cf_csm, cf_region, cf_arr, cf_rag, cf_renewal_status,
       cf_renewal_date, cf_champion_status, cf_open_p1_p2_count
FROM hs_analytics.asana_projects FINAL
WHERE cf_at_risk = 'Yes' OR cf_rag IN ('Red','At Risk')
ORDER BY toFloat32OrZero(cf_arr) DESC

-- Link Asana project → HubSpot deal:
SELECT p.name AS project, p.cf_record_id_es AS hs_deal_id,
       d.deal_name, d.deal_stage, d.amount, p.cf_implementation_status
FROM hs_analytics.asana_projects FINAL p
JOIN hs_analytics.deals FINAL d
  ON p.cf_record_id_es = toString(d.deal_id)
WHERE <deals base filters>

GONG + ASANA ANALYTICAL INTENTS:
- "Call activity / insights for a deal"  → go_extensiveCalls by kore_employee_emailAddress
- "Dark deals / no recent calls"         → deals vs go_extensiveCalls, no match in 30d
- "Competitor mentions this quarter"     → go_extensiveCalls.content_tracker_Competitors_count
- "Talk ratio / rep engagement"          → go_extensiveCalls.interaction_stat_Talk_Ratio
- "Budget / pricing signals"             → tracker count columns in go_extensiveCalls
- "Next steps from calls"                → go_extensiveCalls.content_highlight_Next_steps
- "Rep call volume"                      → go_calls + go_users join on primaryUserId
- "Scorecard questions"                  → go_scorecards by scorecardName
- "Overdue tasks"                        → asana_tasks WHERE completed!='true' AND due_on < today
- "Projects at risk / red RAG"           → asana_projects.cf_rag / cf_at_risk
- "Renewal pipeline health"              → asana_projects cf_renewal_status + cf_arr
- "CSM workload"                         → asana_projects GROUP BY cf_csm
- "Deal → implementation status"         → asana_projects JOIN deals on cf_record_id_es
- "Portfolio health"                     → asana_portfolio_project_association + project status
- "Team task backlog"                    → asana_tasks → asana_project_task_association → project.team

=================================================================
MANDATORY BASE FILTERS (every deals query)
=================================================================
WHERE pipeline = 'default'
AND CASE WHEN deal_type IS NULL THEN 'Not Assigned' ELSE deal_type END
    NOT IN ('Partner-Led SMB')
AND toInt64(deal_id) IN (
    SELECT DISTINCT toInt64(deal_id_hs) FROM kore_ai_hubspot.gs_deal_ids_hs
)

=================================================================
FISCAL YEAR
=================================================================
FY27 = Apr 2026 – Mar 2027. Default to FY27 unless specified.
Active pipeline close_date: >= '2026-04-01' AND <= '2027-03-31'
Stages: '20% - Solution','30% - Proof','40% - Proposal',
        '60% - Price Negotiation','75% - Contract Review'

=================================================================
QUERY RULES
=================================================================
1. SELECT / WITH only — no destructive SQL
2. FINAL on all hs_analytics tables
3. Apply all 3 mandatory base filters on deals
4. For LIST queries: NO LIMIT unless user says "top N" or "first N"
5. countDistinct(deal_id) for unique counts
6. round(sum(amount)/1e6, 1) for $M amounts
7. Dates: toDate(LEFT(coalesce(col,'1900-01-01'),10))
8. Always tell the user the TOTAL count (e.g. "Found 256 deals")

=================================================================
REGION / SOURCE / INDUSTRY MAPPINGS (SELECT only, not WHERE)
=================================================================
Region:  japac→JAPAC, Africa→Middle East, india___sea→ISEA
Source:  Executive Outreach+Investor→Executive Outreach, BDR Outbound→BDR
Industry: Financial Services+Banking+Insurance→Financial Services

CORE RULES:
- NEVER fabricate numbers. Query the DB for every metric.
- NEVER run destructive SQL.
- Answer in clean markdown with tables for data, bold for key numbers.

=================================================================
TARGET TABLES
=================================================================
── TABLE T1: kore_ai_hubspot.gs_pipeline_quotas_v1 ─────────────
Pipeline targets, EOP tracking, attainment: actual ÷ quota

── TABLE T2: kore_ai_hubspot.gs_partner_targets_region_wise ─────
Region-level partner targets and attainment

── TABLE T3: kore_ai_hubspot.gs_partner_targets_psd ─────────────
PSD-level partner performance tracking

── TABLE T4: kore_ai_hubspot.gs_marketing_targets ───────────────
Marketing MQL targets: SELECT region, mql_source, SUM(mql_target) GROUP BY

── TABLE T5: kore_ai_hubspot.gs_closed_won_quotas ───────────────
Closed Won quotas: join to deals where deal_stage = 'Closed Won'

ATTAINMENT: round(actual / target * 100, 1)
COVERAGE:   round(pipeline / revenue_target, 1)

=================================================================
DASHBOARD DEFINITIONS
=================================================================

── EOP DASHBOARD ───────────────────────────────────────────────
EOP Pipeline Value, EOP Target (gs_pipeline_quotas_v1), Attainment %,
Stage-wise and Region-wise breakdown. close_date within current quarter.

── EXEC KPI DASHBOARD ──────────────────────────────────────────
Total Active Pipeline ($M), Closed Won ($M), Attainment %, Win Rate %,
Pipeline Coverage, New Logo Count, ACV Weighted Pipeline.

── CS DASHBOARD ────────────────────────────────────────────────
Renewal Pipeline, Upsell/Expansion Pipeline, Renewal Win Rate,
NRR, At-Risk Deals, CS AE Performance.

── GLOBAL PIPELINE GOVERNANCE ──────────────────────────────────
Pipeline by Region ($M), Partner Pipeline, Partner Attainment,
Coverage Ratio, Closed Won Governance, Marketing Sourced Pipeline.

SELECTION RULE: If user names a dashboard, apply its metric definitions.
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

    stripped = sql.strip().upper()
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
            _store_result(session_id, QueryResult(
                sql           = sql,
                columns       = columns,
                rows          = norm_rows,
                total_rows    = total_rows,
                captured_at   = datetime.utcnow().isoformat() + "Z",
                filters_applied = _extract_filters_from_sql(sql),
            ))

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
    if "BECAME_20_DEAL_DATE" in sql_upper:
        filters.append("Cohort: 20% qualified deals")
    if "BECAME_5_DEAL_DATE" in sql_upper:
        filters.append("Cohort: 5% IQM deals")
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
    return "; ".join(filters) if filters else "Standard base filters applied"


# =============================================================================
# Startup
# =============================================================================
@app.on_event("startup")
async def on_startup():
    global _SYSTEM_PROMPT
    discover_schema()
    _SYSTEM_PROMPT = _build_system_prompt()
    print("🚀 DIUD v6 started — enhanced analytics + SVG funnel viz + Sonnet/Opus selector.")


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
    session_id: Optional[str] = None
    model: Optional[str] = None   # "sonnet" | "opus" | full model string

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
# =============================================================================
def _extract_text(content_blocks) -> str:
    return "\n".join(
        b.text for b in content_blocks if hasattr(b, "text") and b.text
    ).strip()


def _resolve_model(model_hint: Optional[str]) -> str:
    """Resolve user model preference to a full model string."""
    if not model_hint:
        return _DEFAULT_MODEL
    # Accept short names
    if model_hint in MODELS:
        return MODELS[model_hint]
    # Accept full model strings
    if model_hint in MODELS.values():
        return model_hint
    # Fuzzy match
    hint_lower = model_hint.lower()
    if "opus" in hint_lower:
        return MODELS["opus"]
    if "sonnet" in hint_lower:
        return MODELS["sonnet"]
    return _DEFAULT_MODEL


_DATA_QUESTION_PATTERNS = re.compile(
    r'\b(how many|how much|what is|what are|show me|list|give me|'
    r'top \d|pipeline|deals?|stage|funnel|region|AE|attain|win rate|'
    r'closed|open|stall|value|amount|quota|target|coverage|'
    r'breakdown|summary|compare|which|who has|count|total|'
    r'conversion|drop.?off|revenue|forecast|'
    r'gong|call|calls|scorecard|talk.?ratio|sentiment|competitor|'
    r'next.?step|dark.?deal|no.?call|brief|rep.?performance|'
    r'asana|task|tasks|project|portfolio|overdue|backlog|assignee|'
    r'completion|due.?date|team.?workload|subtask)\b',
    re.IGNORECASE,
)

_GREETING_PATTERNS = re.compile(
    r'^(hi|hey|hello|good morning|good afternoon|good evening|'
    r'thanks|thank you|ok|okay|got it|sure|sounds good)[!.,\s]*$',
    re.IGNORECASE,
)

def _is_data_question(message: str) -> bool:
    """True when the last user message is a data/analytics question."""
    msg = message.strip()
    if _GREETING_PATTERNS.match(msg):
        return False
    if len(msg) < 15:                     # very short — probably not a query
        return False
    return bool(_DATA_QUESTION_PATTERNS.search(msg))


def _call_claude(messages: list, max_tokens: int = 8192,
                 session_id: Optional[str] = None,
                 model_hint: Optional[str] = None) -> str:
    """Run Claude with query_clickhouse tool. Up to 5 tool rounds."""
    model = _resolve_model(model_hint)
    print(f"🤖 Using model: {model}")

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

    # Determine whether to force tool use on first turn
    last_user_msg = next(
        (m["content"] for m in reversed(safe_messages) if m["role"] == "user"
         and isinstance(m["content"], str)), ""
    )
    force_tool = _is_data_question(last_user_msg)
    tool_choice = {"type": "any"} if force_tool else {"type": "auto"}
    print(f"   tool_choice={tool_choice['type']} (data_question={force_tool})")

    response = _ai_client.messages.create(
        model=model,
        system=_SYSTEM_PROMPT,
        messages=safe_messages,
        tools=[_QUERY_TOOL],
        tool_choice=tool_choice,
        temperature=0,
        max_tokens=max_tokens,
    )

    for _ in range(5):
        if response.stop_reason != "tool_use":
            break

        tool_block = next((b for b in response.content if b.type == "tool_use"), None)
        if not tool_block:
            break

        sql          = tool_block.input.get("sql", "")
        query_result = run_clickhouse_query(sql, session_id=session_id)
        is_error     = any(query_result.startswith(p) for p in [
            "DATABASE CONNECTION FAILED", "ERROR:", "DATABASE ERROR:"
        ])

        safe_messages = safe_messages + [
            {"role": "assistant", "content": response.content},
            {"role": "user", "content": [{
                "type":        "tool_result",
                "tool_use_id": tool_block.id,
                "content":     query_result,
                "is_error":    is_error,
            }]},
        ]

        response = _ai_client.messages.create(
            model=model,
            system=_SYSTEM_PROMPT,
            messages=safe_messages,
            tools=[_QUERY_TOOL],
            tool_choice={"type": "auto"},   # subsequent rounds: auto
            temperature=0,
            max_tokens=max_tokens,
        )

    reply = _extract_text(response.content)
    if not reply:
        reply = "⚠️ No response generated. Check **/debug/db** to diagnose connectivity."
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
    print(f"💬 [chat] session={payload.session_id} model={payload.model} msg={payload.message[:80]}")

    try:
        reply = _call_claude(messages, session_id=payload.session_id, model_hint=payload.model)
    except Exception as exc:
        traceback.print_exc()
        raise HTTPException(status_code=502, detail=f"Claude error: {exc}")

    has_dataset = payload.session_id is not None and payload.session_id in _SESSION_STORE
    stored = _SESSION_STORE.get(payload.session_id) if payload.session_id else None

    funnel_data = _extract_funnel_data(reply)
    chart_data  = _extract_chart_data(reply)

    return {
        "reply":         reply,
        "has_dataset":   has_dataset,
        "dataset_rows":  stored.total_rows if stored else 0,
        "export_intent": "__EXPORT_INTENT__" in reply,
        "funnel_data":   funnel_data,
        "chart_data":    chart_data,
    }


def _extract_funnel_data(reply: str) -> Optional[dict]:
    """Extract funnel-data JSON block from reply if present."""
    match = re.search(r'```funnel-data\s*\n(.*?)\n```', reply, re.DOTALL)
    if match:
        try:
            return json.loads(match.group(1))
        except Exception:
            return None
    return None


def _extract_chart_data(reply: str) -> Optional[dict]:
    """Extract chart-data JSON block from reply if present."""
    match = re.search(r'```chart-data\s*\n(.*?)\n```', reply, re.DOTALL)
    if match:
        try:
            return json.loads(match.group(1))
        except Exception:
            return None
    return None


# =============================================================================
# Retry
# =============================================================================
class RetryRequest(BaseModel):
    history:    List[ChatMessage] = []
    session_id: Optional[str] = None
    model:      Optional[str] = None


@app.post("/chat/retry")
def chat_retry(payload: RetryRequest):
    if not payload.history:
        raise HTTPException(status_code=400, detail="history must not be empty for retry.")

    clean_history = list(payload.history)
    while clean_history and clean_history[-1].role == "assistant":
        clean_history.pop()

    if not clean_history or clean_history[-1].role != "user":
        raise HTTPException(status_code=400, detail="No user message found to retry.")

    last_user_msg = clean_history[-1].content
    prior_history = clean_history[:-1]

    print(f"🔄 [retry] session={payload.session_id} retrying: {last_user_msg[:80]}")

    messages = [{"role": m.role, "content": m.content} for m in prior_history]
    messages.append({"role": "user", "content": last_user_msg})

    try:
        reply = _call_claude(messages, session_id=payload.session_id, model_hint=payload.model)
    except Exception as exc:
        traceback.print_exc()
        raise HTTPException(status_code=502, detail=f"Claude error on retry: {exc}")

    has_dataset = payload.session_id is not None and payload.session_id in _SESSION_STORE
    stored      = _SESSION_STORE.get(payload.session_id) if payload.session_id else None
    funnel_data = _extract_funnel_data(reply)
    chart_data  = _extract_chart_data(reply)

    return {
        "reply":         reply,
        "has_dataset":   has_dataset,
        "dataset_rows":  stored.total_rows if stored else 0,
        "export_intent": "__EXPORT_INTENT__" in reply,
        "funnel_data":   funnel_data,
        "chart_data":    chart_data,
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
            raise HTTPException(status_code=404, detail="No query result found for this session.")
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


def _build_csv(stored: QueryResult) -> bytes:
    buf = io.StringIO()
    buf.write(f"# Title: {stored.sql[:80]}\n")
    buf.write(f"# Generated: {date.today().isoformat()}\n")
    buf.write(f"# Total Records: {stored.total_rows}\n")
    buf.write(f"# Filters: {stored.filters_applied}\n")
    buf.write(f"# Captured at: {stored.captured_at}\n")
    buf.write("#\n")
    writer = csv.DictWriter(buf, fieldnames=stored.columns, extrasaction="ignore", lineterminator="\n")
    writer.writeheader()
    for row in stored.rows:
        writer.writerow(row)
    return buf.getvalue().encode("utf-8")


def _generate_export_content(
    conversation:    List[ChatMessage],
    title:           str,
    export_type:     str,
    detail_level:    str = "detailed",
    stored_dataset:  Optional[QueryResult] = None,
) -> str:
    conv_text = "\n\n".join(
        f"{'USER' if m.role == 'user' else 'DIUD AGENT'}: {m.content}"
        for m in conversation
    )

    format_hint = (
        "Format as a PowerPoint: use SLIDE: <title> for each slide, then bullet points."
        if export_type == "pptx"
        else "Format as a professional PDF report: ## section headers, narrative prose, tables."
    )
    detail_hint = (
        "Include all metrics and insights. Write [DEAL_TABLE_PLACEHOLDER] where the full deal table should appear."
        if detail_level == "detailed"
        else "Executive summary only — key metrics and top insights, no raw deal list."
    )

    dataset_hint = ""
    if stored_dataset:
        dataset_hint = (
            f"\n\nDATASET CONTEXT: The query returned {stored_dataset.total_rows} records "
            f"with columns: {', '.join(stored_dataset.columns[:12])}. "
            f"Filters: {stored_dataset.filters_applied}. "
            f"The complete table will be injected at [DEAL_TABLE_PLACEHOLDER]."
        )

    prompt = f"""You are preparing a professional {export_type.upper()} export document.

CONVERSATION:
{conv_text}
{dataset_hint}

TASK: Create "{title}"

{format_hint}
{detail_hint}

REQUIREMENTS:
- Executive summary at the start with key numbers
- Logical sections: summary, key metrics, insights, recommendations
- If this is a deal list export, include [DEAL_TABLE_PLACEHOLDER] where the full table belongs
- Bold key numbers; clean professional tone
- Today: {date.today().strftime('%B %d, %Y')}

Generate the document now:"""

    response = _ai_client.messages.create(
        model   = _DEFAULT_MODEL,
        system  = "You are a professional business report writer. Generate clean, well-structured documents.",
        messages= [{"role": "user", "content": prompt}],
        temperature = 0,
        max_tokens  = 4096,
    )
    ai_text = _extract_text(response.content)

    if stored_dataset and stored_dataset.total_rows > 0:
        table_md  = _rows_to_markdown_table(stored_dataset)
        meta_line = (
            f"**Total records:** {stored_dataset.total_rows:,} | "
            f"**Filters:** {stored_dataset.filters_applied} | "
            f"**Exported:** {date.today().strftime('%B %d, %Y')}"
        )
        full_section = f"\n\n## Deal List ({stored_dataset.total_rows:,} records)\n\n{meta_line}\n\n{table_md}"

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
        for line in report_text.split("\n"):
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
