

import re
import json
from typing import Dict, List, Optional, Tuple

PALETTE = [
    "#4E79A7", "#F28E2B", "#E15759", "#76B7B2", "#59A14F",
    "#EDC948", "#B07AA1", "#FF9DA7", "#9C755F", "#BAB0AC",
]

# Canonical funnel stage order — used to sort detected stage rows correctly
_STAGE_ORDER = [
    "5% - IQM Held", "10% - Discovery", "20% - Solution", "30% - Proof",
    "40% - Proposal", "60% - Price Negotiation", "75% - Contract Review",
    "90% - Deal Desk Review", "Closed Won",
]


def _stage_rank(stage: str) -> int:
    for i, s in enumerate(_STAGE_ORDER):
        if s.split(" - ")[0].strip() in str(stage):
            return i
    return 99


def _fmt_money(v: float) -> str:
    if abs(v) >= 1_000_000:
        return f"${v/1_000_000:.1f}M"
    if abs(v) >= 1_000:
        return f"${v/1_000:.0f}K"
    return f"${v:.0f}"


def _to_float(v) -> float:
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0.0


def _is_numericish(v) -> bool:
    try:
        float(v)
        return True
    except (TypeError, ValueError):
        return False


# =============================================================================
# DETECTION — figure out what kind of chart (if any) fits this result
# =============================================================================

def detect_chart_type(columns: List[str], rows: List[dict]) -> Optional[str]:
    """
    Returns one of: 'funnel', 'attainment', 'donut', 'bar_h', 'bar_v', None.
    None means: don't chart this (e.g. raw deal-list / single-value result).
    """
    if not rows or not columns:
        return None

    cols_lower = {c.lower() for c in columns}

    # Pure deal-list (Pattern B) — has deal_id / deal_name as a row-identity
    # column and many distinct rows → it's a list, not an aggregate. Skip.
    if "deal_id" in cols_lower and "deal_name" in cols_lower:
        return None

    # Pattern A — funnel: has a 'stage' column + a count column (deals)
    if "stage" in cols_lower and any(c.lower() in ("deals", "deal_count") for c in columns):
        return "funnel"

    # Pattern C — attainment: actual vs target columns present
    has_actual = any("actual" in c.lower() for c in columns)
    has_target = any("target" in c.lower() for c in columns)
    if has_actual and has_target:
        return "attainment"

    # Generic aggregate: one categorical column + one numeric column
    numeric_cols = [c for c in columns if any(_is_numericish(r.get(c)) for r in rows)]
    categorical_cols = [c for c in columns if c not in numeric_cols]

    if len(numeric_cols) >= 1 and len(categorical_cols) >= 1 and len(rows) >= 2:
        if len(rows) <= 6:
            return "donut"
        return "bar_h"

    # Single aggregate row (e.g. one total) — not chart-worthy on its own
    return None


# =============================================================================
# DATA SHAPING
# =============================================================================

def _aggregate_funnel_rows(rows: List[dict]) -> List[Tuple[str, int, float]]:
    """Collapse possibly-multiple rows per stage (different region/source
    breakdowns) into one (stage, total_deals, total_amount) tuple per stage."""
    agg: Dict[str, List[float]] = {}
    for r in rows:
        stage = str(r.get("stage", "")).strip()
        if not stage:
            continue
        deals = _to_float(r.get("deals") or r.get("deal_count") or 0)
        amount = _to_float(r.get("amount") or r.get("pipeline_m") or 0)
        if stage not in agg:
            agg[stage] = [0.0, 0.0]
        agg[stage][0] += deals
        agg[stage][1] += amount

    out = [(stage, vals[0], vals[1]) for stage, vals in agg.items()]
    out.sort(key=lambda x: _stage_rank(x[0]))
    return out


def _aggregate_attainment_rows(rows: List[dict], columns: List[str]) -> Tuple[List[str], List[float], List[float]]:
    label_col = next(
        (c for c in columns if c.lower() in ("region", "deal_source_rollup", "create_quarter", "quarter")),
        columns[0],
    )
    actual_col = next((c for c in columns if "actual" in c.lower()), None)
    target_col = next((c for c in columns if "target" in c.lower() and "attainment" not in c.lower()), None)

    agg: Dict[str, List[float]] = {}
    for r in rows:
        label = str(r.get(label_col, "")).strip() or "Unknown"
        a = _to_float(r.get(actual_col)) if actual_col else 0
        t = _to_float(r.get(target_col)) if target_col else 0
        if label not in agg:
            agg[label] = [0.0, 0.0]
        agg[label][0] += a
        agg[label][1] += t

    labels = list(agg.keys())
    actuals = [agg[l][0] for l in labels]
    targets = [agg[l][1] for l in labels]
    return labels, actuals, targets


def _aggregate_generic_rows(rows: List[dict], columns: List[str]) -> Tuple[List[str], List[float], str]:
    numeric_cols = [c for c in columns if any(_is_numericish(r.get(c)) for r in rows)]
    categorical_cols = [c for c in columns if c not in numeric_cols]
    label_col = categorical_cols[0] if categorical_cols else columns[0]
    value_col = numeric_cols[0] if numeric_cols else columns[-1]

    agg: Dict[str, float] = {}
    for r in rows:
        label = str(r.get(label_col, "")).strip() or "Unknown"
        agg[label] = agg.get(label, 0.0) + _to_float(r.get(value_col))

    items = sorted(agg.items(), key=lambda kv: kv[1], reverse=True)[:12]
    labels = [k for k, _ in items]
    values = [v for _, v in items]
    return labels, values, value_col


# =============================================================================
# HTML BUILDERS
# =============================================================================

def _wrap_html(body: str, title: str, subtitle: str) -> str:
    return f"""```html
<!DOCTYPE html>
<html><head>
<meta charset="UTF-8">
<style>
  * {{ box-sizing: border-box; margin:0; padding:0; font-family: 'Inter', system-ui, sans-serif; }}
  body {{ background:#fff; padding:16px; }}
  .chart-card {{ background:#fff; border:1px solid #E2E8F0; border-radius:12px; padding:18px; }}
  .chart-title {{ font-size:14px; font-weight:700; color:#0D1B3E; margin-bottom:2px; }}
  .chart-subtitle {{ font-size:11px; color:#94A3B8; margin-bottom:14px; }}
  .legend {{ display:flex; flex-wrap:wrap; gap:10px; margin-top:12px; }}
  .legend-item {{ display:flex; align-items:center; gap:6px; font-size:11px; color:#334155; }}
  .legend-swatch {{ width:10px; height:10px; border-radius:3px; flex-shrink:0; }}
</style>
</head>
<body>
<div class="chart-card">
  <div class="chart-title">{title}</div>
  <div class="chart-subtitle">{subtitle}</div>
  {body}
</div>
</body></html>
```"""


def build_funnel_html(stored_rows: List[dict], subtitle: str) -> Optional[str]:
    stages = _aggregate_funnel_rows(stored_rows)
    if not stages:
        return None

    max_deals = max(s[1] for s in stages) or 1
    rows_html = []
    for i, (stage, deals, amount) in enumerate(stages):
        pct_width = max((deals / max_deals) * 88, 18)
        color = PALETTE[i % len(PALETTE)]
        conv = f"{(deals / stages[0][1] * 100):.0f}%" if stages[0][1] else "—"
        rows_html.append(f"""
        <div style="display:flex;align-items:center;gap:12px;margin-bottom:10px;">
          <div style="width:160px;font-size:11px;color:#334155;text-align:right;flex-shrink:0;">{stage}</div>
          <div style="flex:1;background:#F1F5F9;border-radius:6px;height:28px;position:relative;overflow:hidden;">
            <div style="width:{pct_width:.1f}%;height:100%;background:{color};border-radius:6px;
                        display:flex;align-items:center;justify-content:flex-end;padding-right:8px;">
              <span style="font-size:11px;color:white;font-weight:700;white-space:nowrap;">
                {int(deals):,} deals
              </span>
            </div>
          </div>
          <div style="width:110px;font-size:11px;color:#475569;flex-shrink:0;">
            {_fmt_money(amount * (1_000_000 if amount < 10000 else 1))} · {conv}
          </div>
        </div>""")

    body = "<div>" + "".join(rows_html) + "</div>"
    return _wrap_html(body, "Pipeline Funnel", subtitle)


def build_attainment_html(stored_rows: List[dict], columns: List[str], subtitle: str) -> Optional[str]:
    labels, actuals, targets = _aggregate_attainment_rows(stored_rows, columns)
    if not labels:
        return None

    max_val = max(max(actuals, default=0), max(targets, default=0)) or 1
    rows_html = []
    for i, label in enumerate(labels):
        a, t = actuals[i], targets[i]
        a_pct = max((a / max_val) * 100, 1)
        t_pct = max((t / max_val) * 100, 1)
        attain = f"{(a / t * 100):.0f}%" if t else "—"
        rows_html.append(f"""
        <div style="margin-bottom:14px;">
          <div style="display:flex;justify-content:space-between;font-size:11px;color:#334155;margin-bottom:4px;">
            <span style="font-weight:600;">{label}</span>
            <span style="color:#1565C0;font-weight:700;">{attain}</span>
          </div>
          <div style="background:#F1F5F9;border-radius:5px;height:14px;margin-bottom:3px;position:relative;">
            <div style="width:{a_pct:.1f}%;height:100%;background:#1E88E5;border-radius:5px;"></div>
          </div>
          <div style="background:#F8FAFC;border:1px dashed #CBD5E1;border-radius:5px;height:8px;width:{t_pct:.1f}%;"></div>
          <div style="font-size:10px;color:#94A3B8;margin-top:2px;">
            Actual: {_fmt_money(a)} &nbsp;|&nbsp; Target: {_fmt_money(t)}
          </div>
        </div>""")

    legend = """
    <div class="legend">
      <div class="legend-item"><span class="legend-swatch" style="background:#1E88E5;"></span>Actual</div>
      <div class="legend-item"><span class="legend-swatch" style="background:#CBD5E1;"></span>Target</div>
    </div>"""
    body = "<div>" + "".join(rows_html) + "</div>" + legend
    return _wrap_html(body, "Actual vs Target", subtitle)


def build_bar_html(stored_rows: List[dict], columns: List[str], horizontal: bool, subtitle: str) -> Optional[str]:
    labels, values, value_col = _aggregate_generic_rows(stored_rows, columns)
    if not labels:
        return None

    max_val = max(values, default=0) or 1
    rows_html = []
    for i, (label, val) in enumerate(zip(labels, values)):
        pct = max((val / max_val) * 100, 2)
        color = PALETTE[i % len(PALETTE)]
        display_val = _fmt_money(val) if val >= 1000 else f"{val:,.1f}"
        rows_html.append(f"""
        <div style="display:flex;align-items:center;gap:10px;margin-bottom:9px;">
          <div style="width:140px;font-size:11px;color:#334155;text-align:right;flex-shrink:0;
                      overflow:hidden;text-overflow:ellipsis;white-space:nowrap;" title="{label}">{label}</div>
          <div style="flex:1;background:#F1F5F9;border-radius:5px;height:22px;position:relative;">
            <div style="width:{pct:.1f}%;height:100%;background:{color};border-radius:5px;
                        display:flex;align-items:center;justify-content:flex-end;padding-right:8px;">
              <span style="font-size:10.5px;color:white;font-weight:700;white-space:nowrap;">{display_val}</span>
            </div>
          </div>
        </div>""")

    body = "<div>" + "".join(rows_html) + "</div>"
    return _wrap_html(body, value_col.replace("_", " ").title(), subtitle)


def build_donut_html(stored_rows: List[dict], columns: List[str], subtitle: str) -> Optional[str]:
    labels, values, value_col = _aggregate_generic_rows(stored_rows, columns)
    if not labels:
        return None
    total = sum(values) or 1

    legend_items = []
    for i, (label, val) in enumerate(zip(labels, values)):
        color = PALETTE[i % len(PALETTE)]
        pct = val / total * 100
        legend_items.append(f"""
        <div class="legend-item">
          <span class="legend-swatch" style="background:{color};"></span>
          {label}: {_fmt_money(val) if val >= 1000 else f'{val:,.1f}'} ({pct:.0f}%)
        </div>""")

    # Conic-gradient donut, pure CSS — no JS dependency, always renders
    gradient_parts = []
    cum = 0.0
    for i, val in enumerate(values):
        color = PALETTE[i % len(PALETTE)]
        start = cum / total * 360
        cum += val
        end = cum / total * 360
        gradient_parts.append(f"{color} {start:.1f}deg {end:.1f}deg")
    gradient = ", ".join(gradient_parts)

    body = f"""
    <div style="display:flex;align-items:center;gap:24px;flex-wrap:wrap;">
      <div style="width:140px;height:140px;border-radius:50%;
                  background:conic-gradient({gradient});
                  position:relative;flex-shrink:0;">
        <div style="position:absolute;inset:22px;background:#fff;border-radius:50%;
                    display:flex;align-items:center;justify-content:center;
                    font-size:11px;color:#64748B;text-align:center;">
          {_fmt_money(total) if total >= 1000 else f'{total:,.0f}'}
        </div>
      </div>
      <div class="legend" style="flex-direction:column;gap:8px;">{"".join(legend_items)}</div>
    </div>"""
    return _wrap_html(body, value_col.replace("_", " ").title(), subtitle)


# =============================================================================
# PUBLIC ENTRY POINT
# =============================================================================

def build_chart_html(
    columns: List[str],
    rows: List[dict],
    filters_applied: str = "",
) -> Optional[str]:
    """
    Deterministically build an embeddable ```html chart block from stored
    query results. Returns None if the result isn't chart-worthy (e.g. a
    raw deal list or a single scalar value) — callers should NOT force a
    chart in that case.
    """
    chart_type = detect_chart_type(columns, rows)
    if chart_type is None:
        return None

    subtitle = filters_applied[:140] if filters_applied else f"{len(rows)} rows"

    if chart_type == "funnel":
        return build_funnel_html(rows, subtitle)
    if chart_type == "attainment":
        return build_attainment_html(rows, columns, subtitle)
    if chart_type == "donut":
        return build_donut_html(rows, columns, subtitle)
    if chart_type == "bar_h":
        return build_bar_html(rows, columns, horizontal=True, subtitle=subtitle)

    return None


def reply_already_has_chart(reply: str) -> bool:
    """Detect if Claude already included an ```html chart block, so we don't
    double-inject."""
    return bool(re.search(r'```html', reply, re.IGNORECASE))


