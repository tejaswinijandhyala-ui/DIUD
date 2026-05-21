# PipeGen Chat

Conversational pipeline intelligence powered by **Claude** + **ClickHouse**.  
Ask natural language questions about your deals, AEs, regions, win/loss, BANT, and funnel metrics — and export the conversation as a branded PDF report or PPTX deck.

---

## Architecture

```
chat.html          ← Single-page UI (sidebar suggestions, chat bubbles, export buttons)
main.py            ← FastAPI backend
  ├── POST /chat          → Claude + query_clickhouse tool loop
  ├── POST /export/pdf    → Claude generates structured report → ReportLab renders PDF
  └── POST /export/pptx   → Claude generates slide content → python-pptx renders deck
```

**Flow for every chat message:**
1. Full conversation history is sent to Claude with the ClickHouse schema prompt
2. Claude decides whether to query the DB (`query_clickhouse` tool use)
3. Up to 4 tool-use rounds are executed automatically
4. Final reply returned as markdown, rendered in the browser

**Flow for export:**
1. Full conversation is summarised by Claude into a structured report / slide content  
2. Claude may run additional DB queries to fill in missing data  
3. Server builds the PDF/PPTX and streams it as a download  

---

## Setup

### 1. Clone / copy files
```
pipegen-chat/
├── main.py
├── chat.html
├── requirements.txt
└── .env
```

### 2. Install dependencies
```bash
pip install -r requirements.txt
```

### 3. Configure `.env`
```bash
cp .env.example .env
```
Edit `.env`:
```
ANTHROPIC_API_KEY=sk-ant-...
CLICKHOUSE_API_URL=https://your-clickhouse-host/query
CLICKHOUSE_API_TOKEN=your_bearer_token
```

### 4. Run
```bash
uvicorn main:app --host 0.0.0.0 --port 8000 --reload
```
Open **http://localhost:8000**

---

## Customising the Prompt

The entire schema, business definitions, and example queries live in `_SYSTEM_PROMPT` inside `main.py`.  
Sections to update for your environment:

| Section | What to change |
|---|---|
| `TABLES` | Column names, table names, helper tables |
| `MANDATORY BASE FILTERS` | Your whitelist / pipeline filter logic |
| `FISCAL YEAR CALCULATION` | Your FY start month |
| `REGION MAP` | Your raw → display region values |
| `INDUSTRY MAP` | Your industry groupings |
| `DEAL STAGE LIST` | Your stage names and velocity benchmarks |
| `BUSINESS DEFINITIONS` | Your KPI and segment definitions |
| `SAMPLE QUERIES` | Representative queries for your use case |

---

## Export Formats

### PDF
- Multi-page A4 portrait  
- Navy branded header/footer on every page  
- Sections: Executive Summary, Pipeline Health, Key Metrics, Regional Breakdown, Risk & Opportunities, Recommended Actions  
- Each section has a colour-coded accent bar  

### PPTX
- 16:9 widescreen (13.33″ × 7.5″)  
- Dark navy cover slide  
- One content slide per section with accent bars and bullet points  
- Branded footer on every slide  

---

## API Endpoints

| Method | Path | Description |
|---|---|---|
| GET | `/` | Serves `chat.html` |
| POST | `/chat` | `{message, history[]}` → `{reply}` |
| POST | `/export/pdf` | `{title, conversation[]}` → PDF stream |
| POST | `/export/pptx` | `{title, conversation[]}` → PPTX stream |

---

## Security Notes

- The `run_clickhouse_query` function only allows `SELECT`/`WITH` statements  
- Always tighten `allow_origins` in CORSMiddleware before deploying to production  
- Store secrets in `.env` — never commit it  
