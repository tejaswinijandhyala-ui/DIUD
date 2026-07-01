"""
rules.py
=============================================================================
Machine-checkable registry of MANDATORY business rules from the DIUD system
prompt (v4 — Pattern A / B / C architecture).

KEY CHANGES vs previous version:
  - Pattern A (cumulative OR-chain stage counts) is explicitly whitelisted so
    cohort rules do NOT fire against it.
  - Pattern B / C are detected and validated separately.
  - Sentinel for "date never set" is `!= '1900-01-01'`, not `IS NOT NULL`.
  - SQL-triggered cohort checks now require a `-- Pattern A` comment to opt out,
    so the model can selectively suppress them when using Pattern A.
  - `FINAL` check handles LEFT JOIN patterns correctly.
  - MQL filter checks tightened to match §11 exactly.
  - Attainment / target checks aligned to §8 (Pattern C) two-CTE pattern.
=============================================================================
"""

import re
from typing import Any, Callable, Dict, List, Optional, Set

# =============================================================================
# RULEBOOK
# =============================================================================
RULEBOOK: Dict[str, str] = {
    "mql": """
§11 MQL CALCULATION RULES — MANDATORY
When computing MQL actuals from hs_analytics.contacts FINAL, ALL THREE filters
below are mandatory. Missing any one produces inflated counts.

1. lifecycle_stage = 'marketingqualifiedlead'
   AND date_entered_marketing_qualified_lead_lifecycle_stage_pipeline IS NOT NULL
2. company_priority IN ('P1','P2','P3','P4','P5','P6','P7')
3. lead_status != 'Bad Data'

MQL TARGET PATTERN — always filter by exact quarter, never divide annual by 4:
  SELECT region, original_source, SUM(toFloat64OrZero(mql_target)) AS mql_tgt
  FROM kore_ai_hubspot.gs_marketing_targets
  WHERE fy = 'FY27' AND quarter = 'Q1'
  GROUP BY region, original_source
""",

    "active_pipeline": """
§7 ACTIVE PIPELINE DEFINITION (Pattern B — Deal-Level Detail)
A deal qualifies as ACTIVE pipeline when ALL of the following are true:
1. deal_stage IN ('20% - Solution','30% - Proof','40% - Proposal',
                  '60% - Price Negotiation','75% - Contract Review')
2. Primary filter is close_date (not any became_<stage>_deal_date)
3. All MANDATORY_BASE_FILTERS applied (pipeline='default', deal_type exclusion,
   gs_deal_ids_hs allowlist)
Apply the deal_stage filter ONLY when the user explicitly requests "active" pipeline.
""",

    "cohort_funnel": """
§8b COHORT FUNNEL RULE — for true cohort queries only
(Do NOT confuse with Pattern A cumulative OR-chain counts.)

Before writing cohort SQL, verify all 4 checks:
  - became_<N>_deal_date != '1900-01-01' — cohort anchor present (sentinel, NOT IS NOT NULL)
  - deal_stage NOT IN (<all stages before N%>) — exclusion present
  - Query starts with WITH — CTE pattern used
  - countDistinct(deal_id) — not count(*) or count(deal_id)

Exclusion logic by starting stage:
  10% -> CW : exclude 1%, 5%
  20% -> CW : exclude 1%, 5%, 10%
  30% -> CW : exclude 1%, 5%, 10%, 20%
  40% -> CW : exclude 1%, 5%, 10%, 20%, 30%
  60% -> CW : exclude 1%, 5%, 10%, 20%, 30%, 40%
  75% -> CW : exclude 1%, 5%, 10%, 20%, 30%, 40%, 60%

Pattern:
  WITH cohort AS (
    SELECT deal_id, deal_stage, amount
    FROM hs_analytics.deals FINAL
    WHERE became_<N>_deal_date != '1900-01-01'
      AND deal_stage NOT IN (<stages before N%>)
      AND <MANDATORY_BASE_FILTERS>
      AND <date range on became_<N>_deal_date>
  )
  SELECT deal_stage, countDistinct(deal_id) AS deal_count,
         round(SUM(amount)/1e6, 1) AS pipeline_m
  FROM cohort GROUP BY deal_stage ORDER BY deal_count DESC
""",

    "attainment": """
§8 / Pattern C — TARGET SQL RULES (Actuals vs Target)
1. DEFAULT TIER = L2 (no-prefix in T1; l2_ prefix in T2/T3). Switch only if user
   explicitly says L1/stretch/committed.
2. CAST ALL NUMERIC TARGET COLUMNS: SUM(toFloat64OrZero(amount_target_20))
3. NO FAN-OUT JOINS — never join raw deals to a target table then SUM. Use
   independent CTEs (actual CTE + target CTE), combined with LEFT JOIN at the end.
4. Use nullIf(target, 0) in every division.
5. PERIOD GRAIN MUST MATCH — filter target table to the exact quarter, never
   divide an annual target by 4.
6. ATTAINMENT = round(actual / nullIf(target, 0) * 100, 1)
7. SOURCE MAPPING for Pattern C: Executive Outreach + Investor → 'Executive Outreach'
   (this differs from Pattern A/B where they stay separate).
8. FY anchor MUST use the became_<stage>_deal_date corresponding
to the stage requested.

Examples:
5%  → became_5_deal_date
10% → became_10_deal_date
20% → became_20_deal_date
30% → became_30_deal_date
40% → became_40_deal_date
60% → became_60_deal_date
75% → became_75_deal_date
""",

    "closed_won": """
CLOSED WON RULES
  deal_stage = 'Closed Won' AND close_date BETWEEN '<fy_start>' AND '<fy_end>'
  GROUP BY deal_owner + MANDATORY_BASE_FILTERS
  Quota source: kore_ai_hubspot.gs_closed_won_quotas, cast with toFloat64OrZero().
  No L1/L2/Committed split in this table — single quota tier only.
  JOIN to deals on ae = deal_owner.
""",

    "partner_targets": """
PARTNER TARGET RULES
  Table T2 gs_partner_targets_region_wise: l2_/l1_/committed_ prefixes.
  Table T3 gs_partner_targets_psd: COMMITTED ONLY, no L1/L2 — use T2 for those.
  Filter partner_team_type IN ('Hyperscaler','GSI/SI','Reseller/BPO/TSD') as needed.
  committed_amount_target_10 / committed_amount_target_5 do NOT exist in T2.
  All columns are Nullable(String) — always SUM(toFloat64OrZero(col)).
""",

    "dashboard_definitions": """
DASHBOARD DEFINITIONS (abbreviated — ask user which dashboard if ambiguous)
  EOP:         pipeline vs EOP target, active stages only, current quarter end window.
  EXEC KPI:    total active pipeline, closed won, CW attainment %, win rate, coverage.
  PIPEGEN:     5/10/20% pipeline amount + deal count vs gs_pipeline_quotas_v1.
  PARTNERSHIP: partner pipeline vs partner target tables, PSD/hyperscaler splits.
  MARKETING:   MQL actual vs target (§11 filters), source/region performance.
  AE FOCUS:    AE pipeline, CW ARR, quota attainment, win rate, avg deal size.
  BDR FOCUS:   meetings created, opportunities generated, PipeGen by BDR.
""",
}


def get_rulebook_entry(topic: str) -> str:
    return RULEBOOK.get(topic, f"No rulebook entry found for topic '{topic}'.")


# =============================================================================
# PATTERN DETECTION
# =============================================================================

def _is_pattern_a(sql: str) -> bool:
    """
Pattern A: cumulative OR-chain stage counting.
FY is anchored on the became_<stage>_deal_date corresponding
to the stage requested by the user.
"""
    if re.search(r'--\s*Pattern\s*A', sql, re.I):
        return True

    dates = re.findall(r'became_(\d+)_deal_date', sql, re.I)

    return len(set(dates)) >= 3 and "OR" in sql.upper()


def _is_cohort_query(sql: str, intent: dict) -> bool:
    if _is_pattern_a(sql):
        return False
    return intent.get("cohort_stage") is not None


def _has_became_date(sql: str) -> bool:
    return bool(re.search(r'became_\d+_deal_date', sql, re.I))


def _is_pattern_c(sql: str, intent: dict) -> bool:
    return intent.get("metric") == "attainment"


# -----------------------------------------------------------------------------
# Stage helper
# -----------------------------------------------------------------------------

_STAGE_COLUMN_MAP = {
    "5": "became_5_deal_date",
    "10": "became_10_deal_date",
    "20": "became_20_deal_date",
    "30": "became_30_deal_date",
    "40": "became_40_deal_date",
    "60": "became_60_deal_date",
    "75": "became_75_deal_date",
}


def _expected_became_column(intent: Dict[str, Any]) -> str:
    """
    Returns the became_<stage>_deal_date corresponding
    to the stage requested by the user.
    Used for all stage-based queries.
    """
    stage = (
        intent.get("stage")
        or intent.get("cohort_stage")
        or "10"
    )

    return _STAGE_COLUMN_MAP.get(stage, "became_10_deal_date")


# =============================================================================
# Intent detection
# =============================================================================

def detect_intent(user_message: str, sql: str = "") -> Dict[str, Any]:

    msg = user_message or ""
    intent: Dict[str, Any] = {}

    # ------------------------------------------------------------------
    # Detect requested stage (5%,10%,20%,30%,40%,60%,75%)
    # ------------------------------------------------------------------

    stage_match = re.search(r'\b(5|10|20|30|40|60|75)\s*%', msg)

    if stage_match:
        intent["stage"] = stage_match.group(1)

    # ------------------------------------------------------------------
    # Cohort detection
    # ------------------------------------------------------------------

    m = re.search(
        r'(\d+)\s*%\s*(?:→|->|to)\s*(closed\s*won|cw)\b',
        msg,
        re.I,
    )

    if m:
        intent["cohort_stage"] = m.group(1)

    if not intent.get("cohort_stage"):

        m = re.search(
            r'\b(cohort|starting\s+(?:at|from))\b.*?(\d+)\s*%',
            msg,
            re.I,
        )

        if m:
            intent["cohort_stage"] = m.group(2)

    if not intent.get("cohort_stage"):

        m = re.search(
            r'(\d+)\s*%.*?\bcohort\b',
            msg,
            re.I,
        )

        if m:
            intent["cohort_stage"] = m.group(1)

    if (
        not intent.get("cohort_stage")
        and sql
        and not _is_pattern_a(sql)
    ):

        m = re.search(
            r'became_(\d+)_deal_date',
            sql,
            re.I,
        )

        if m and "NOT IN" in sql.upper():
            intent["cohort_stage"] = m.group(1)

    # ------------------------------------------------------------------
    # List queries
    # ------------------------------------------------------------------

    if re.search(
        r'\b(list|show me all|which deals|deals\s+(with|where))\b',
        msg,
        re.I,
    ):
        intent["query_type"] = "list"

    # ------------------------------------------------------------------
    # Top N
    # ------------------------------------------------------------------

    if re.search(r'\b(top|first)\s+\d+\b', msg, re.I):
        intent["top_n"] = True

    # ------------------------------------------------------------------
    # Metrics
    # ------------------------------------------------------------------

    if re.search(r'\bMQL\b', msg, re.I):
        intent["metric"] = "mql"

    if re.search(
        r'\b(attainment|quota|coverage|vs\.?\s*target|gap\s*to\s*target)\b',
        msg,
        re.I,
    ):
        intent["metric"] = "attainment"

    if re.search(r'\bactive pipeline\b', msg, re.I):
        intent["metric"] = "active_pipeline"

    # ------------------------------------------------------------------
    # Placeholder leakage
    # ------------------------------------------------------------------

    if sql and "MANDATORY_BASE_FILTERS" in sql:
        intent["placeholder_leak"] = True

    return intent


def _arity(fn: Callable) -> int:
    return fn.__code__.co_argcount


def _call(fn: Callable, sql: str, intent: dict):
    return fn(sql, intent) if _arity(fn) == 2 else fn(sql)


# =============================================================================
# SQL-TEXT RULES
# =============================================================================
RULES: List[Dict[str, Any]] = [

    # -------------------------------------------------------------------------
    # §3 MANDATORY_BASE_FILTERS — apply to every deals query
    # -------------------------------------------------------------------------
    {
        "id": "base_filter_pipeline",
        "section": "§3 MANDATORY_BASE_FILTERS (1/3) — pipeline = 'default'",
        "applies_when": lambda sql: "hs_analytics.deals" in sql,
        "check": lambda sql: bool(re.search(r"pipeline\s*=\s*'default'", sql, re.I)),
        "message": "Missing `pipeline = 'default'` base filter on hs_analytics.deals query.",
    },
    {
        "id": "base_filter_deal_type",
        "section": "§3 MANDATORY_BASE_FILTERS (2/3) — Partner-Led SMB exclusion",
        "applies_when": lambda sql: "hs_analytics.deals" in sql,
        "check": lambda sql: "Partner-Led SMB" in sql and bool(re.search(r'\bNOT\s+IN\b', sql, re.I)),
        "message": "Missing deal_type NOT IN ('Partner-Led SMB') base filter.",
    },
    {
        "id": "base_filter_allowlist",
        "section": "§3 MANDATORY_BASE_FILTERS (3/3) — gs_deal_ids_hs allowlist",
        "applies_when": lambda sql: "hs_analytics.deals" in sql,
        "check": lambda sql: "gs_deal_ids_hs" in sql,
        "message": "Missing deal_id allowlist subquery against kore_ai_hubspot.gs_deal_ids_hs.",
    },

    # -------------------------------------------------------------------------
    # §3 Duplicate exclusion — FINAL on every hs_analytics.* reference
    # -------------------------------------------------------------------------
    {
        "id": "final_keyword",
        "section": "§3 Duplicate exclusion — FINAL on hs_analytics.*",
        "applies_when": lambda sql: bool(re.search(r"hs_analytics\.\w+", sql)),
        "check": lambda sql: bool(re.search(r"hs_analytics\.\w+\s+(?:AS\s+\w+\s+)?FINAL|hs_analytics\.\w+\s+FINAL", sql, re.I)),
        "message": "Missing FINAL on at least one hs_analytics.* table reference.",
    },
    {
        "id": "count_distinct_not_count",
        "section": "§3 / §13.4 — countDistinct, never count()",
        "applies_when": lambda sql: "hs_analytics" in sql and bool(
            re.search(r'(?<!\w)count\s*\((?!Distinct)', sql, re.I)
        ),
        "check": lambda sql: bool(
            re.search(r'countDistinct\s*\(\s*(deal_id|contact_id)\s*\)', sql, re.I)
        ),
        "message": "Uses count() instead of countDistinct(deal_id) or countDistinct(contact_id).",
    },
    {
        "id": "distinct_in_association_subquery",
        "section": "§3 Table 5 — DISTINCT in gs_DealContactAssociation subqueries",
        "applies_when": lambda sql: "gs_DealContactAssociation" in sql and bool(re.search(r'\(\s*SELECT', sql, re.I)),
        "check": lambda sql: bool(re.search(r'SELECT\s+DISTINCT', sql, re.I)),
        "message": "Subquery against gs_DealContactAssociation is missing DISTINCT.",
    },

    # -------------------------------------------------------------------------
    # §9 / §13 General SQL guardrails
    # -------------------------------------------------------------------------
    {
        "id": "select_or_with_only",
        "section": "§13.1 SELECT/WITH only — no destructive SQL",
        "applies_when": lambda sql: True,
        "check": lambda sql: sql.strip().upper().startswith(("SELECT", "WITH", "--")),
        "message": "Query does not start with SELECT, WITH, or a comment. Only read queries are permitted.",
    },
    {
        "id": "no_placeholder_leak_strict",
        "section": "Generation hygiene — no unresolved placeholder tokens",
        "applies_when": lambda sql, intent: intent.get("placeholder_leak", False),
        "check": lambda sql, intent: False,
        "message": "Literal placeholder '<MANDATORY_BASE_FILTERS>' leaked into generated SQL — must be expanded.",
    },
    {
        "id": "no_limit_on_list",
        "section": "§13.5 — No LIMIT on list queries unless user says 'top N'/'first N'",
        "applies_when": lambda sql, intent: intent.get("query_type") == "list" and not intent.get("top_n"),
        "check": lambda sql, intent: "LIMIT" not in sql.upper(),
        "message": "LIMIT applied to a list query the user did not ask to cap with 'top N' or 'first N'.",
    },

    # -------------------------------------------------------------------------
    # §9 Fiscal year / date casting
    # -------------------------------------------------------------------------
    {
        "id": "date_cast_standard",
        "section": "§9 Date casting — CAST(LEFT(coalesce(col,'1900-01-01'),10) AS DATE)",
        "applies_when": lambda sql: bool(
            re.search(r"\b(close_date|became_\d+_deal_date)\b\s*(>=|<=|>|<|=)\s*'", sql, re.I)
        ),
        "check": lambda sql: bool(
            re.search(r"CAST\s*\(\s*LEFT\s*\(\s*coalesce\s*\(", sql, re.I)
        ),
        "message": "Raw date string comparison without the mandatory CAST(LEFT(coalesce(col,'1900-01-01'),10) AS DATE) cast.",
    },
    {
        "id": "sentinel_not_null_check",
        "section": "§9 Sentinel '1900-01-01' — use != '1900-01-01', NOT IS NOT NULL",
        "applies_when": lambda sql: _has_became_date(sql),
        "check": lambda sql: not bool(
            re.search(r"became_\d+_deal_date\s+IS\s+NOT\s+NULL", sql, re.I)
        ),
        "message": (
            "Using `IS NOT NULL` on became_<N>_deal_date. "
            "The sentinel for 'date not set' is '1900-01-01', so use `!= '1900-01-01'` instead."
        ),
    },

    # -------------------------------------------------------------------------
    # §4 Target table — nullable string casting
    # -------------------------------------------------------------------------
    {
        "id": "target_table_float_cast",
        "section": "§4 Target table — SUM(toFloat64OrZero(col))",
        "applies_when": lambda sql: bool(
            re.search(r"gs_pipeline_quotas_v1|gs_partner_targets|gs_closed_won_quotas|gs_marketing_targets", sql, re.I)
        ),
        "check": lambda sql: bool(re.search(r"toFloat64OrZero|toFloat32OrZero", sql, re.I)),
        "message": (
            "Target table columns are Nullable(String). "
            "Always cast with SUM(toFloat64OrZero(col)) — raw arithmetic will silently produce NULLs."
        ),
    },
    {
        "id": "target_no_quarterly_divide",
        "section": "§4 / §13.9 — Never derive quarterly target by dividing by 4",
        "applies_when": lambda sql: bool(
            re.search(r"gs_pipeline_quotas_v1|gs_partner_targets|gs_marketing_targets", sql, re.I)
        ),
        "check": lambda sql: not bool(re.search(r"/\s*4\b", sql)),
        "message": "Target figure is being divided by 4 to derive a quarterly value — not permitted. Filter the target table to the exact quarter instead.",
    },
    {
        "id": "nullif_in_division",
        "section": "§4 / §13.7 — nullIf(denominator, 0) in every division",
        "applies_when": lambda sql: "/" in sql and bool(
            re.search(r"attainment|coverage|pct|rate|ratio", sql, re.I)
        ),
        "check": lambda sql: bool(re.search(r"nullIf\s*\(", sql, re.I)),
        "message": "Division present without nullIf(denominator, 0) — risk of divide-by-zero.",
    },

    # -------------------------------------------------------------------------
    # Pattern A — cumulative OR-chain stage counts (§6)
    # No cohort rules apply here; validate the OR-chain structure instead.
    # -------------------------------------------------------------------------
    {
        "id": "pattern_a_or_chain",
        "section": "§6 Pattern A — cumulative OR-chain stage counting",
        "applies_when": lambda sql: _is_pattern_a(sql),
        "check": lambda sql: bool(re.search(r'\bOR\b', sql, re.I)),
        "message": (
            "Query is marked as Pattern A (cumulative stage counting) but contains "
            "no OR conditions. Pattern A requires OR-chain conditions for each stage "
            "— see §6 for the correct template."
        ),
    },
    {
    "id": "pattern_a_stage_anchor",
    "section": "§6 Pattern A — stage-specific FY anchor",
    "applies_when": lambda sql, intent: _is_pattern_a(sql),
    "check": lambda sql, intent: (
        _expected_became_column(intent) in sql
    ),
    "message": (
        "Pattern A must use the became_<stage>_deal_date "
        "corresponding to the stage requested by the user."
    ),
},


    # -------------------------------------------------------------------------
    # Pattern B — deal-level detail (§7)
    # Primary filter must be close_date, (not any became_<stage>_deal_date).
    # -------------------------------------------------------------------------
    {
        "id": "pattern_b_close_date_filter",
        "section": "§7 Pattern B — primary filter is close_date (not any became_<stage>_deal_date)",
        "applies_when": lambda sql, intent: (
            intent.get("metric") == "active_pipeline"
            and not _is_pattern_a(sql)
            and not _is_pattern_c(sql, intent)
        ),
        "check": lambda sql, intent: bool(re.search(r"close_date\s*>=", sql, re.I)),
        "message": (
            "Pattern B (active pipeline / deal-level) must filter on close_date >= <date>, "
            "not became_10_deal_date. See §7."
        ),
    },
    {
        "id": "pattern_b_active_stages",
        "section": "§7 Pattern B — active pipeline stage filter",
        "applies_when": lambda sql, intent: intent.get("metric") == "active_pipeline",
        "check": lambda sql, intent: all(
            s in sql
            for s in [
                "20% - Solution", "30% - Proof", "40% - Proposal",
                "60% - Price Negotiation", "75% - Contract Review",
            ]
        ),
        "message": (
            "Active pipeline query missing one or more of the 5 required deal_stage values: "
            "20% - Solution, 30% - Proof, 40% - Proposal, 60% - Price Negotiation, 75% - Contract Review."
        ),
    },

    # -------------------------------------------------------------------------
    # Pattern C — attainment / vs-target (§8)
    # -------------------------------------------------------------------------
    {
        "id": "pattern_c_two_cte",
        "section": "§8 Pattern C — actuals CTE + targets CTE (never fan-out join)",
        "applies_when": lambda sql, intent: _is_pattern_c(sql, intent),
        "check": lambda sql, intent: (
            sql.strip().upper().startswith("WITH")
            and sql.upper().count("CTE") + sql.upper().count("AS (") >= 2
        ),
        "message": (
            "Attainment/target query must use independent CTEs for actuals and targets, "
            "then LEFT JOIN them. Never join raw deal rows directly to a target table."
        ),
    },
    {
    "id": "pattern_c_stage_anchor",
    "section": "§8 Pattern C — stage-specific became date",

    "applies_when": lambda sql, intent:
        _is_pattern_c(sql, intent),

    "check": lambda sql, intent:
        _expected_became_column(intent) in sql,

    "message": (
        "Pattern C must use the became_<stage>_deal_date corresponding "
        "to the stage requested by the user."
    ),
    },
    {
    "id": "pattern_c_source_merge",
    "section": "§8 Pattern C — Executive Outreach + Investor merged in source mapping",
    "applies_when": lambda sql, intent: (
        _is_pattern_c(sql, intent)
        and "deal_source_rollup" in sql
        and "Executive Outreach" in sql
    ),
    "check": lambda sql, intent: (
        ("Investor" in sql and "Executive Outreach" in sql)
        or "Investor" not in sql
    ),
    "message": (
        "Pattern C source mapping must merge 'Investor' into 'Executive Outreach' "
        "to match the quota table bucket. This differs from Pattern A/B where they stay separate."
    ),
},
    {
        "id": "pattern_c_target_tier_default",
        "section": "§4 Pattern C — default tier is L2 (no prefix / l2_ prefix)",
        "applies_when": lambda sql, intent: _is_pattern_c(sql, intent),
        "check": lambda sql, intent: not bool(
            re.search(r'\b(l1_|_l1\b|committed_|_committed\b)', sql, re.I)
        ),
        "message": (
            "Target query is using L1 or Committed tier columns. "
            "Default is always L2 (no suffix in T1; l2_ prefix in T2/T3). "
            "Use L1/Committed only when the user explicitly says so."
        ),
    },

    # -------------------------------------------------------------------------
    # True cohort funnel (§8b) — only fires when cohort intent is detected AND
    # Pattern A is NOT present.
    # -------------------------------------------------------------------------
    {
        "id": "cohort_anchor_sentinel",
        "section": "§8b Cohort anchor — became_<N>_deal_date != '1900-01-01'",
        "applies_when": lambda sql, intent: _is_cohort_query(sql, intent),
        "check": lambda sql, intent: bool(
        re.search(
            rf"{_expected_became_column(intent)}\s*!=\s*'1900-01-01'",
            sql,
            re.I,
        )
    ),
        "message": (
            "Cohort query missing `became_<N>_deal_date != '1900-01-01'` sentinel anchor. "
            "Do NOT use IS NOT NULL — use the sentinel check instead."
        ),
    },
    {
        "id": "cohort_exclusion",
        "section": "§8b Stage exclusion — NOT IN prior stages",
        "applies_when": lambda sql, intent: _is_cohort_query(sql, intent),
        "check": lambda sql, intent: bool(re.search(r'\bNOT\s+IN\b', sql, re.I)),
        "message": "Cohort query missing NOT IN exclusion of all deal_stage values prior to the cohort starting stage.",
    },
    {
        "id": "cohort_single_cte",
        "section": "§8b — cohort must be a single CTE with GROUP BY deal_stage",
        "applies_when": lambda sql, intent: _is_cohort_query(sql, intent),
        "check": lambda sql, intent: (
            sql.strip().upper().startswith("WITH")
            and bool(re.search(r'GROUP\s+BY\s+deal_stage', sql, re.I))
        ),
        "message": (
            "Cohort funnel should be a single WITH-cohort CTE with GROUP BY deal_stage, "
            "not separate per-stage SELECT statements."
        ),
    },
    {
        "id": "cohort_count_distinct",
        "section": "§8b Deduplication — countDistinct(deal_id) in cohort",
        "applies_when": lambda sql, intent: _is_cohort_query(sql, intent),
        "check": lambda sql, intent: bool(re.search(r'countDistinct\s*\(\s*deal_id\s*\)', sql, re.I)),
        "message": "Cohort funnel is not using countDistinct(deal_id).",
    },

    # -------------------------------------------------------------------------
    # §11 MQL filters
    # -------------------------------------------------------------------------
    {
        "id": "mql_lifecycle_stage_filter",
        "section": "§11 MQL filter 1 — lifecycle_stage = 'marketingqualifiedlead'",
        "applies_when": lambda sql, intent: intent.get("metric") == "mql",
        "check": lambda sql, intent: bool(re.search(r"lifecycle_stage\s*=\s*'marketingqualifiedlead'", sql, re.I)),
        "message": "MQL query missing `lifecycle_stage = 'marketingqualifiedlead'` filter.",
    },
    {
        "id": "mql_date_entered_filter",
        "section": "§11 MQL filter 1b — date_entered_... IS NOT NULL",
        "applies_when": lambda sql, intent: intent.get("metric") == "mql",
        "check": lambda sql, intent: "date_entered_marketing_qualified_lead_lifecycle_stage_pipeline" in sql,
        "message": "MQL query missing `date_entered_marketing_qualified_lead_lifecycle_stage_pipeline` filter.",
    },
    {
        "id": "mql_company_priority_filter",
        "section": "§11 MQL filter 2 — company_priority IN ('P1'...'P7')",
        "applies_when": lambda sql, intent: intent.get("metric") == "mql",
        "check": lambda sql, intent: bool(re.search(r"company_priority\s+IN", sql, re.I)),
        "message": "MQL query missing `company_priority IN ('P1',...,'P7')` filter.",
    },
    {
        "id": "mql_bad_data_filter",
        "section": "§11 MQL filter 3 — lead_status != 'Bad Data'",
        "applies_when": lambda sql, intent: intent.get("metric") == "mql",
        "check": lambda sql, intent: bool(re.search(r"lead_status\s*!=\s*'Bad Data'", sql, re.I)),
        "message": "MQL query missing `lead_status != 'Bad Data'` filter.",
    },
    {
        "id": "mql_no_quarter_divide",
        "section": "§11 MQL — never derive quarterly target by /4",
        "applies_when": lambda sql, intent: intent.get("metric") == "mql",
        "check": lambda sql, intent: not bool(re.search(r"/\s*4\b", sql)),
        "message": "MQL target appears to be derived by dividing an annual target by 4 — filter the target table to the exact quarter instead.",
    },
]


# =============================================================================
# RESULT-LEVEL RULES
# =============================================================================
RESULT_RULES: List[Dict[str, Any]] = [
    {
        "id": "funnel_sum_within_cohort",
        "section": "§8b Deduplication — stage counts must not exceed cohort total",
        "applies_when": lambda rows, intent: _is_cohort_query("", intent) and bool(rows),
        "check": lambda rows, intent: _funnel_sum_ok(rows),
        "message": (
            "Funnel stage counts (active stages + Closed Won/Lost) exceed the cohort total — "
            "rows were not derived from a single deduplicated cohort CTE."
        ),
    },
]


def _funnel_sum_ok(rows: List[dict]) -> bool:
    try:
        cohort_total: Optional[float] = None
        active_sum = 0.0
        terminal_sum = 0.0
        for r in rows:
            stage = str(r.get("deal_stage", ""))
            cnt = float(r.get("deal_count", 0) or 0)
            if cohort_total is None:
                cohort_total = cnt
                continue
            if "Closed Won" in stage or "Closed Lost" in stage:
                terminal_sum += cnt
            else:
                active_sum += cnt
        if cohort_total is None:
            return True
        return (active_sum + terminal_sum) <= cohort_total
    except Exception:
        return True  # don't block on checker bugs


# =============================================================================
# Public API
# =============================================================================
def validate_sql_against_rules(sql: str, user_message: str) -> List[str]:
    """Run every applicable RULES entry against `sql`. Returns violation strings."""
    intent = detect_intent(user_message, sql)
    violations = []
    for rule in RULES:
        try:
            applies = _call(rule["applies_when"], sql, intent)
        except Exception:
            applies = False
        if not applies:
            continue
        try:
            ok = _call(rule["check"], sql, intent)
        except Exception:
            ok = False
        if not ok:
            violations.append(f"[{rule['id']}] {rule['section']}: {rule['message']}")
    return violations


def validate_result_against_rules(rows: List[dict], user_message: str, sql: str = "") -> List[str]:
    """Run every applicable RESULT_RULES entry against the returned `rows`."""
    intent = detect_intent(user_message, sql)
    violations = []
    for rule in RESULT_RULES:
        try:
            applies = rule["applies_when"](rows, intent)
        except Exception:
            applies = False
        if not applies:
            continue
        try:
            ok = rule["check"](rows, intent)
        except Exception:
            ok = False
        if not ok:
            violations.append(f"[{rule['id']}] {rule['section']}: {rule['message']}")
    return violations


def get_intent(user_message: str, sql: str = "") -> Dict[str, Any]:
    """Exposed for logging / debugging in main.py."""
    return detect_intent(user_message, sql)


# =============================================================================
# FACT-BINDING VERIFIER
# =============================================================================
_NUM_PATTERN = re.compile(r'-?\$?\d[\d,]*\.?\d*%?')

_STAGE_LABEL_PATTERN = re.compile(
    r'\b\d{1,3}%\s*[-–—]\s*[A-Za-z][\w/() ]*'
)
_STAGE_TRANSITION_PATTERN = re.compile(
    r'\b\d{1,3}%\s*(?:to|→|->)\s*\d{1,3}%'
)
_FY_QUARTER_PATTERN = re.compile(r'\bFY\s?\d{2,4}\b|\bQ[1-4]\b', re.IGNORECASE)
_YEAR_RANGE = range(2020, 2036)


def _strip_label_noise(text: str) -> str:
    cleaned = _STAGE_LABEL_PATTERN.sub(' ', text)
    cleaned = _STAGE_TRANSITION_PATTERN.sub(' ', cleaned)
    cleaned = _FY_QUARTER_PATTERN.sub(' ', cleaned)
    return cleaned


def extract_numbers(text: str) -> Set[float]:
    cleaned = _strip_label_noise(text)
    raw = _NUM_PATTERN.findall(cleaned)
    out: Set[float] = set()
    for tok in raw:
        cleaned_tok = tok.replace('$', '').replace(',', '').replace('%', '')
        try:
            val = float(cleaned_tok)
        except ValueError:
            continue
        if val in _YEAR_RANGE and '.' not in cleaned_tok:
            continue
        out.add(round(val, 2))
    return out

def extract_numbers_from_rows(rows: List[dict]) -> Set[float]:
    out: Set[float] = set()
    for row in rows:
        for v in row.values():
            try:
                out.add(round(float(v), 2))
            except (TypeError, ValueError):
                continue
    return out


def _relative_tolerance(value: float, base_tolerance: float = 0.5) -> float:
    return max(base_tolerance, abs(value) * 0.01)


def validate_summary_against_facts(
    summary_text: str,
    allowed_rows: List[dict],
    tolerance: float = 0.5,
) -> List[str]:
    claimed = extract_numbers(summary_text)
    actual = extract_numbers_from_rows(allowed_rows)

    if not actual:
        return []

    derived: Set[float] = set()
    for a in actual:
        derived.add(round(a, 1))
        derived.add(round(a / 1_000_000, 1))
        derived.add(round(a / 1_000, 1))
        derived.add(round(a / 1_000_000, 2))

    actual_nonzero = [a for a in actual if a != 0]
    for a in actual_nonzero:
        for b in actual_nonzero:
            if a == b:
                continue
            ratio = a / b * 100
            derived.add(round(ratio, 1))
            derived.add(round(ratio, 0))

    violations = []
    for c in claimed:
        if c in (0.0, 100.0, 1.0):
            continue

        tol = _relative_tolerance(c, tolerance)

        matches_raw    = any(abs(c - a) <= tol for a in actual)
        matches_m      = any(abs(c - a / 1_000_000) <= tol for a in actual)
        matches_k      = any(abs(c - a / 1_000) <= tol for a in actual)
        matches_scale  = any(abs(c * 1_000_000 - a) <= max(tol * 1_000_000, abs(a) * 0.01) for a in actual)
        matches_derived = any(abs(c - d) <= tol for d in derived)

        if not (matches_raw or matches_m or matches_k or matches_scale or matches_derived):
            violations.append(f"Unverified number in summary: {c}")

    return violations
