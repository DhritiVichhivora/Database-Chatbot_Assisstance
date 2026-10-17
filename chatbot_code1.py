#!/usr/bin/env python
# coding: utf-8

# ══════════════════════════════════════════════════════════════
# SUPERSTORE SALES DATABASE CHATBOT — Refactored
# Changes from previous version:
#   1. .env file support for all config (URL, model, DB)
#   2. Single call_model() utility function
#   3. Intent router — classifies every question
#   4. Mistral for generic/contextual/recommendation
#   5. Gemma3 only for SQL generation
#   6. Customers table JOIN support
#   7. Recommendation + subjective question handling
# ══════════════════════════════════════════════════════════════


import requests
import pandas as pd
import re
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from sqlalchemy import create_engine, text
from datetime import datetime, timedelta
import textwrap
import os
import io

# ── Predictive analytics imports ──────────────────────────────
try:
    import numpy as np
    from sklearn.linear_model import LinearRegression
    SKLEARN_AVAILABLE = True
except ImportError:
    SKLEARN_AVAILABLE = False
    print("  ⚠️  scikit-learn not installed — pip install scikit-learn numpy")

# ── LOAD .env FILE ────────────────────────────────────────────
# All configuration lives in .env — see .env.example for the
# full list of required keys (DB_URL, GROQ_API_KEY, etc.)

def load_env():
    env = {}
    env_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), '.env')
    if os.path.exists(env_path):
        with open(env_path) as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith('#') and '=' in line:
                    key, val = line.split('=', 1)
                    env[key.strip()] = val.strip()
        print("  ✅ .env file loaded")
    else:
        print("  ⚠️  No .env file found — using environment variables")
    # Real environment variables (e.g. set on Render) take precedence
    for key in ('OLLAMA_URL', 'SQL_MODEL', 'FAST_MODEL', 'DB_URL',
                'GROQ_API_KEY', 'GROQ_SQL_MODEL', 'GROQ_FAST_MODEL'):
        if os.environ.get(key):
            env[key] = os.environ[key]
    return env

env = load_env()

# ── CONFIG ────────────────────────────────────────────────────
OLLAMA_URL    = env.get('OLLAMA_URL', '')
SQL_MODEL     = env.get('SQL_MODEL', 'gemma3:12b')
FAST_MODEL    = env.get('FAST_MODEL', 'mistral')
DB_URL        = env.get('DB_URL', '')
GROQ_API_KEY  = env.get('GROQ_API_KEY', '')
GROQ_SQL_MODEL  = env.get('GROQ_SQL_MODEL', 'llama-3.3-70b-versatile')
GROQ_FAST_MODEL = env.get('GROQ_FAST_MODEL', 'llama-3.3-70b-versatile')

# ── Word doc export ───────────────────────────────────────────
try:
    from docx import Document
    from docx.shared import Inches, Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False

# ── DATABASE ──────────────────────────────────────────────────
engine = create_engine(DB_URL)

print()
print("=" * 60)
print("       SUPERSTORE SALES DATABASE CHATBOT")
print("=" * 60)
print()

try:
    with engine.connect() as conn:
        conn.execute(text("SELECT 1"))
    print("  ✅ Database connected successfully")
except Exception as e:
    print(f"  ❌ Database connection failed: {e}")

if GROQ_API_KEY:
    print(f"  ✅ AI model connected (Groq — {GROQ_SQL_MODEL})")
else:
    print("  ⚠️  No GROQ_API_KEY found in .env — AI calls will fail")

if DOCX_AVAILABLE:
    print("  ✅ Word export available")
else:
    print("  ⚠️  Word export unavailable — pip install python-docx")

print()
print("=" * 60)

context = []
latest_chart_buf = None


# ══════════════════════════════════════════════════════════════
# UTILITY FUNCTION — call_model()
# Single function to call any model. All config lives in .env —
# no URLs, keys, or model names hardcoded here.
# ══════════════════════════════════════════════════════════════

def call_model(prompt, model=None, temperature=0.4, top_p=0.9,
               repeat_penalty=1.1, num_ctx=4096, timeout=120):
    if model is None:
        model = FAST_MODEL

    try:
        from groq import Groq
        client = Groq(api_key=GROQ_API_KEY)
        groq_model = GROQ_SQL_MODEL if model == SQL_MODEL else GROQ_FAST_MODEL
        response = client.chat.completions.create(
            model=groq_model,
            messages=[{"role": "user", "content": prompt}],
            temperature=min(temperature, 1.0),
            max_tokens=2048
        )
        return response.choices[0].message.content.strip()
    except Exception as e:
        print(f"  ❌ Groq error: {e}")
        return ""
# ══════════════════════════════════════════════════════════════
# DATABASE METADATA
# ══════════════════════════════════════════════════════════════

def get_metadata():
    cache_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'metadata_cache.txt')
    if os.path.exists(cache_path):
        with open(cache_path, 'r') as f:
            return f.read()
    metadata = "DATABASE STRUCTURE:\n"
    try:
        with engine.connect() as conn:
            tables = conn.execute(text("""
                SELECT table_name FROM information_schema.tables
                WHERE table_schema = 'public' ORDER BY table_name
            """)).fetchall()
            for table in tables:
                table_name = table[0]
                metadata += f"\nTable: {table_name}\n"
                metadata += "Columns:\n"
                columns = conn.execute(text(f"""
                    SELECT column_name, data_type
                    FROM information_schema.columns
                    WHERE table_name = '{table_name}'
                    ORDER BY ordinal_position
                """)).fetchall()
                for col in columns:
                    metadata += f"  - {col[0]} ({col[1]})\n"
                count = conn.execute(text(f"SELECT COUNT(*) FROM {table_name}")).fetchone()[0]
                metadata += f"Total rows: {count}\n"
    except Exception as e:
        print(f"  ❌ Metadata error: {e}")
    try:
        cache_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'metadata_cache.txt')
        with open(cache_path, 'w') as f:
            f.write(metadata)
    except Exception:
        pass
    return metadata

print("\n  📋 Loading metadata...")
metadata = get_metadata()
print("  ✅ Metadata ready")
print("=" * 60)

def get_special_columns():
    special_cols = {}
    try:
        with engine.connect() as conn:
            tables = conn.execute(text("""
                SELECT table_name, column_name FROM information_schema.columns
                WHERE table_schema = 'public' AND column_name ~ '[^a-zA-Z0-9_]'
            """)).fetchall()
            for table_name, col_name in tables:
                normalized = col_name.replace('-', '_').replace(' ', '_').replace('.', '_')
                special_cols[normalized] = f'"{col_name}"'
                special_cols[col_name] = f'"{col_name}"'
    except Exception as e:
        print(f"  ⚠️ Could not get special columns: {e}")
    return special_cols

special_columns = get_special_columns()
print(f"  ✅ Special columns detected: {special_columns}")

def get_table_columns():
    table_cols = {}
    try:
        with engine.connect() as conn:
            tables = conn.execute(text("""
                SELECT table_name, column_name FROM information_schema.columns
                WHERE table_schema = 'public' ORDER BY table_name, ordinal_position
            """)).fetchall()
            for table_name, col_name in tables:
                if table_name not in table_cols:
                    table_cols[table_name] = []
                table_cols[table_name].append(col_name.lower())
    except Exception as e:
        print(f"  ⚠️ Could not get table columns: {e}")
    return table_cols

table_columns = get_table_columns()
print(f"  ✅ Table columns loaded: {list(table_columns.keys())}")


# ══════════════════════════════════════════════════════════════
# DATE INJECTION SYSTEM
# ══════════════════════════════════════════════════════════════

MONTH_NAMES = {
    'january': 1, 'february': 2, 'march': 3, 'april': 4,
    'may': 5, 'june': 6, 'july': 7, 'august': 8,
    'september': 9, 'october': 10, 'november': 11, 'december': 12,
    'jan': 1, 'feb': 2, 'mar': 3, 'apr': 4,
    'jun': 6, 'jul': 7, 'aug': 8, 'sep': 9,
    'sept': 9, 'oct': 10, 'nov': 11, 'dec': 12
}
QUARTER_MONTHS = {1: 1, 2: 4, 3: 7, 4: 10}

def get_quarter_range(year, quarter):
    start_month = QUARTER_MONTHS[quarter]
    start = datetime(year, start_month, 1)
    if quarter == 4:
        end = datetime(year + 1, 1, 1)
    else:
        end = datetime(year, QUARTER_MONTHS[quarter + 1], 1)
    return start.strftime('%Y-%m-%d'), end.strftime('%Y-%m-%d')

def detect_date_context(question):
    q = question.lower().strip()
    today = datetime.now()
    now_year = today.year
    now_month = today.month
    result = None

    m = re.search(r'\b(\d{1,2})(?:st|nd|rd|th)?\s+(january|february|march|april|may|june|july|august|september|october|november|december|jan|feb|mar|apr|jun|jul|aug|sep|sept|oct|nov|dec)\s+(\d{4})\b', q)
    if m:
        day, month_str, year = int(m.group(1)), m.group(2), int(m.group(3))
        month = MONTH_NAMES[month_str]
        date_str = f"{year}-{month:02d}-{day:02d}"
        result = {'filter': f"order_date = '{date_str}'::date", 'desc': f"on {date_str}"}

    if not result:
        m = re.search(r'\b(january|february|march|april|may|june|july|august|september|october|november|december|jan|feb|mar|apr|jun|jul|aug|sep|sept|oct|nov|dec)\s+(\d{1,2})(?:st|nd|rd|th)?\s+(\d{4})\b', q)
        if m:
            month_str, day, year = m.group(1), int(m.group(2)), int(m.group(3))
            month = MONTH_NAMES[month_str]
            date_str = f"{year}-{month:02d}-{day:02d}"
            result = {'filter': f"order_date = '{date_str}'::date", 'desc': f"on {date_str}"}

    if not result:
        m = re.search(r'\b(\d{4}-\d{2}-\d{2})\b', q)
        if m:
            date_str = m.group(1)
            result = {'filter': f"order_date = '{date_str}'::date", 'desc': f"on {date_str}"}

    if not result:
        m = re.search(r'\b(january|february|march|april|may|june|july|august|september|october|november|december|jan|feb|mar|apr|jun|jul|aug|sep|sept|oct|nov|dec)\s+(\d{4})\b', q)
        if m:
            month_str, year = m.group(1), int(m.group(2))
            month = MONTH_NAMES[month_str]
            start = f"{year}-{month:02d}-01"
            end_month = month + 1 if month < 12 else 1
            end_year = year if month < 12 else year + 1
            end = f"{end_year}-{end_month:02d}-01"
            result = {'filter': f"order_date >= '{start}'::date AND order_date < '{end}'::date", 'desc': f"in {m.group(1).title()} {year}"}

    if not result:
        m = re.search(r'\bq([1-4])\s+(?:of\s+)?(\d{4})\b', q)
        if m:
            quarter, year = int(m.group(1)), int(m.group(2))
            start, end = get_quarter_range(year, quarter)
            result = {'filter': f"order_date >= '{start}'::date AND order_date < '{end}'::date", 'desc': f"Q{quarter} {year}"}

    if not result:
        quarter_words = {'first': 1, 'second': 2, 'third': 3, 'fourth': 4}
        m = re.search(r'\b(first|second|third|fourth)\s+quarter\s+(?:of\s+)?(\d{4})\b', q)
        if m:
            quarter, year = quarter_words[m.group(1)], int(m.group(2))
            start, end = get_quarter_range(year, quarter)
            result = {'filter': f"order_date >= '{start}'::date AND order_date < '{end}'::date", 'desc': f"Q{quarter} {year}"}

    if not result:
        m = re.search(r'\b(in|during|for|of)\s+(202[0-9])\b', q)
        if not m:
            m = re.search(r'\b(202[0-9])\b', q)
            if m:
                year = int(m.group(1))
                result = {'filter': f"order_date >= '{year}-01-01'::date AND order_date < '{year+1}-01-01'::date", 'desc': f"in {year}"}
        else:
            year = int(m.group(2))
            result = {'filter': f"order_date >= '{year}-01-01'::date AND order_date < '{year+1}-01-01'::date", 'desc': f"in {year}"}

    if not result and re.search(r'\blast\s+week\b', q):
        start = (today - timedelta(days=7)).strftime('%Y-%m-%d')
        result = {'filter': f"order_date >= '{start}'::date AND order_date < CURRENT_DATE", 'desc': "last week"}

    if not result and re.search(r'\bthis\s+week\b', q):
        start_of_week = (today - timedelta(days=today.weekday())).strftime('%Y-%m-%d')
        result = {'filter': f"order_date >= '{start_of_week}'::date AND order_date < CURRENT_DATE", 'desc': "this week"}

    if not result and re.search(r'\blast\s+month\b', q):
        if now_month == 1:
            start = f"{now_year - 1}-12-01"
            end = f"{now_year}-01-01"
        else:
            start = f"{now_year}-{now_month-1:02d}-01"
            end = f"{now_year}-{now_month:02d}-01"
        result = {'filter': f"order_date >= '{start}'::date AND order_date < '{end}'::date", 'desc': "last month"}

    if not result and re.search(r'\bmonth\s+before\s+(last|that)\b|\b2\s+months\s+ago\b', q):
        two_months_ago = now_month - 2
        year_adj = now_year
        if two_months_ago <= 0:
            two_months_ago += 12
            year_adj -= 1
        one_month_ago = now_month - 1
        year_adj2 = now_year
        if one_month_ago <= 0:
            one_month_ago += 12
            year_adj2 -= 1
        start = f"{year_adj}-{two_months_ago:02d}-01"
        end = f"{year_adj2}-{one_month_ago:02d}-01"
        result = {'filter': f"order_date >= '{start}'::date AND order_date < '{end}'::date", 'desc': "month before last"}

    if not result and re.search(r'\bthis\s+month\b', q):
        start = f"{now_year}-{now_month:02d}-01"
        result = {'filter': f"order_date >= '{start}'::date AND order_date < CURRENT_DATE", 'desc': "this month"}

    if not result and re.search(r'\blast\s+quarter\b', q):
        current_quarter = (now_month - 1) // 3 + 1
        last_quarter = current_quarter - 1
        year = now_year
        if last_quarter == 0:
            last_quarter = 4
            year -= 1
        start, end = get_quarter_range(year, last_quarter)
        result = {'filter': f"order_date >= '{start}'::date AND order_date < '{end}'::date", 'desc': "last quarter"}

    if not result and re.search(r'\bthis\s+quarter\b', q):
        current_quarter = (now_month - 1) // 3 + 1
        start, end = get_quarter_range(now_year, current_quarter)
        result = {'filter': f"order_date >= '{start}'::date AND order_date < CURRENT_DATE", 'desc': "this quarter"}

    if not result and re.search(r'\blast\s+year\b', q):
        start = f"{now_year - 1}-01-01"
        end = f"{now_year}-01-01"
        result = {'filter': f"order_date >= '{start}'::date AND order_date < '{end}'::date", 'desc': "last year"}

    if not result and re.search(r'\bthis\s+year\b', q):
        start = f"{now_year}-01-01"
        result = {'filter': f"order_date >= '{start}'::date AND order_date < CURRENT_DATE", 'desc': "this year"}

    if not result:
        m = re.search(r'\blast\s+(\d+)\s+days?\b', q)
        if m:
            n = int(m.group(1))
            start = (today - timedelta(days=n)).strftime('%Y-%m-%d')
            result = {'filter': f"order_date >= '{start}'::date AND order_date < CURRENT_DATE", 'desc': f"last {n} days"}

    if not result:
        m = re.search(r'\blast\s+(\d+)\s+months?\b', q)
        if m:
            n = int(m.group(1))
            month = now_month - n
            year = now_year
            while month <= 0:
                month += 12
                year -= 1
            start = f"{year}-{month:02d}-01"
            result = {'filter': f"order_date >= '{start}'::date AND order_date < CURRENT_DATE", 'desc': f"last {n} months"}

    if not result:
        m = re.search(r'\bbetween\s+(january|february|march|april|may|june|july|august|september|october|november|december|jan|feb|mar|apr|jun|jul|aug|sep|sept|oct|nov|dec)\s+(\d{4})\s+and\s+(january|february|march|april|may|june|july|august|september|october|november|december|jan|feb|mar|apr|jun|jul|aug|sep|sept|oct|nov|dec)\s+(\d{4})\b', q)
        if m:
            m1, y1 = MONTH_NAMES[m.group(1)], int(m.group(2))
            m2, y2 = MONTH_NAMES[m.group(3)], int(m.group(4))
            start = f"{y1}-{m1:02d}-01"
            end_month = m2 + 1 if m2 < 12 else 1
            end_year = y2 if m2 < 12 else y2 + 1
            end = f"{end_year}-{end_month:02d}-01"
            result = {'filter': f"order_date >= '{start}'::date AND order_date < '{end}'::date", 'desc': f"between {m.group(1).title()} {y1} and {m.group(3).title()} {y2}"}

    return result


def inject_date_into_prompt(question):
    comparison_words = [
        "compare", "vs", "versus", "month before", "last month vs",
        "this year vs", "last year vs", "previous", "month before last",
        "2 months ago", "last quarter vs", "this quarter vs"
    ]
    if any(word in question.lower() for word in comparison_words):
        print(f"  📅 Comparison question — skipping date injection")
        return None
    date_ctx = detect_date_context(question)
    if date_ctx:
        print(f"  📅 Date detected: {date_ctx['desc']} → {date_ctx['filter']}")
        return date_ctx
    return None


# ══════════════════════════════════════════════════════════════
# INTENT ROUTER
# Classifies every question before processing
# Uses rule-based checks first (fast), then AI for ambiguous ones
# ══════════════════════════════════════════════════════════════

GREETING_EXACT = {
    "hello", "hi", "hey", "good morning", "good afternoon",
    "good evening", "how are you", "thanks", "thank you",
    "bye", "goodbye", "see you", "thats all", "that's all",
    "good job", "well done", "nice work", "what can you do","amazing","okay nice"
}

CONTEXTUAL_REFS = {
    "that", "it", "its", "those", "they", "them", "this",
    "that region", "that category", "that customer", "that product",
    "that state", "that city", "same", "also", "too", "as well",
    "more about", "tell me more", "explain", "elaborate", "details","first", "second", "third", "fourth", "fifth",
    "first one", "second one", "top one", "number one",
    "that person", "that customer", "he", "she", "they"
}

RECOMMENDATION_KEYWORDS = {
    "should", "recommend", "suggest", "ideal", "best strategy",
    "what would you", "advice", "advise", "optimize", "improve",
    "increase", "decrease", "focus on", "invest", "priority",
    "inventory", "reorder", "stock", "forecast", "predict",
    "what can we do", "how can we", 
    "how to improve", "action", "plan"
}

SUBJECTIVE_SQL_KEYWORDS = {
    "why", "explain why", "what caused", "reason for",
    "what drives", "insight", "analyse", "analyze",
    "why is", "why are", "why does", "performing better",
    "performing worse", "outperforming", "underperforming"
}

def classify_intent(question):
    """
    Classifies into: generic | sql | recommendation | subjective_sql | contextual
    Hard rules for clear cases, AI for everything else.
    """
    q = question.lower().strip().rstrip("!?.")

    # Rule 1: Exact greetings
    # Rule 1: Exact greetings and acknowledgements
    if q in GREETING_EXACT:
        return "generic"
    
    # Short acknowledgements — never contextual
    if q in {"okay", "ok", "got it", "understood", "sure", "alright", 
             "fine", "noted", "great", "nice", "cool", "good", "perfect",
             "i see", "i understand", "makes sense", "interesting","nice"}:
        return "generic"
    
    # Get real context (skip summary pseudo-entry)
    # Get real context (skip summary pseudo-entry)
    real_ctx = [ex for ex in context if ex.get("question") != "__summary__"]

    # If context exists — always contextual, no exceptions
    if real_ctx:
        return "contextual"

    # No context — first question, normal routing
    sql_phrases = [
        "how many", "how much", "total sales", "total profit", "total orders",
        "total returns", "show me", "give me all details", "all details",
        "details of", "which city", "which state", "which region",
        "which category", "which segment", "which channel", "which device",
        "which month", "which year", "which quarter", "which payment",
        "top ", "bottom ", "highest", "lowest", "most", "least",
        "average delivery", "average order"
    ]
    if any(p in q for p in sql_phrases):
        return "sql"

    if any(w in q for w in ['recommend', 'suggestion', 'suggest']):
        return "recommendation"

    # Predictive intent — forecast/prediction questions
    predictive_keywords = [
        'predict', 'forecast', 'next month', 'next quarter', 'next year',
        'will sales', 'will profit', 'will return', 'will orders',
        'expected', 'projection', 'trend for next', 'future sales',
        'future profit', 'how will', 'likely to',
        'going to be', 'estimate for', 'predicted', 'if we reduce',
        'if we increase', 'what would sales', 'what would profit',
        'scenario', 'when will', 'peak month', 'growth rate',
        'compared to last year', 'vs last year', 'year over year',
        'this time next year', 'will it hit', 'at risk of',
        'churn next', 'will demand', 'will customer', 'how many new customers',
        'next 3 months', 'next 6 months', 'next 12 months', 'will improve',
        'will underperform', 'will decline', 'will grow', 'will increase',
        'will decrease', 'q1 20', 'q2 20', 'q3 20', 'q4 20'
    ]
    if any(w in q for w in predictive_keywords):
        return "predictive"
    
    return "sql"

# ════════════════════════════════════════════ ══════════════════
# HANDLERS — one per intent type
# ══════════════════════════════════════════════════════════════

def handle_generic(question):
    """Handles greetings and casual conversation using Mistral."""
    prompt = f"""You are a friendly Sales Analytics Assistant for Superstore.
The user said: '{question}'
Respond naturally in 1-2 sentences. Be warm and helpful. Be direct too. 
If they greet you, greet back and offer to help with sales data.
Never mention SQL or technical terms.Never give any details as to from where you have gotten the answer the from too. Never start with 'I'."""
    return call_model(prompt, model=FAST_MODEL, temperature=0.8, num_ctx=512, timeout=30)


def _resolve_reference(question, prev_result_str):
    """
    Parse the previous raw result and resolve vague references like
    'lowest', 'highest', 'first', 'second', 'that one' into exact values.
    Returns a dict with resolved info, or empty dict if can't resolve.
    """
    if not prev_result_str:
        return {}
    try:
        import pandas as pd
        import io
        df = pd.read_csv(io.StringIO(prev_result_str), sep=r'\s{2,}', engine='python')
        df = df.dropna(how='all')
        if df.empty:
            return {}

        q = question.lower()
        resolved = {}

        # Detect ordinal references
        ordinal_map = {
            'first': 0, '1st': 0, 'number one': 0, 'top one': 0,
            'second': 1, '2nd': 1, 'number two': 1,
            'third': 2, '3rd': 2, 'number three': 2,
            'fourth': 3, '4th': 3,
            'fifth': 4, '5th': 4,
        }
        for word, idx in ordinal_map.items():
            if word in q and idx < len(df):
                row = df.iloc[idx]
                resolved['row'] = row.to_dict()
                resolved['description'] = f"the {word} row: {row.to_dict()}"
                return resolved

        # Detect lowest/minimum reference
        if any(w in q for w in ['lowest', 'minimum', 'least', 'bottom', 'smallest', 'worst']):
            # Find numeric columns
            num_cols = df.select_dtypes(include='number').columns.tolist()
            if num_cols:
                col = num_cols[-1]  # last numeric col (usually the metric)
                idx = df[col].idxmin()
                row = df.loc[idx]
                resolved['row'] = row.to_dict()
                resolved['description'] = f"the row with the LOWEST {col} value ({row[col]}): {row.to_dict()}"
                return resolved

        # Detect highest/maximum reference
        if any(w in q for w in ['highest', 'maximum', 'most', 'top', 'best', 'largest']):
            num_cols = df.select_dtypes(include='number').columns.tolist()
            if num_cols:
                col = num_cols[-1]
                idx = df[col].idxmax()
                row = df.loc[idx]
                resolved['row'] = row.to_dict()
                resolved['description'] = f"the row with the HIGHEST {col} value ({row[col]}): {row.to_dict()}"
                return resolved

        return {}
    except Exception as e:
        print(f"  ⚠️ _resolve_reference error: {e}")
        return {}


def _classify_contextual_subintent(question, prev_question, prev_answer, prev_result):
    q = question.lower().strip()

    # Subjective — why/explain questions
    if any(w in q for w in [
        "why", "explain", "what caused", "reason", "what drives",
        "how come", "what happened", "insight", "analyse", "analyze"
    ]):
        return "subjective"

    # Text only — no DB needed
    # Text only — no DB needed, and question doesn't ask for data
    if any(w in q for w in [
        "bundle", "group them", "categorise", "categorize",
        "summarise", "summarize"
    ]):
        return "text"

    # Everything else — go to DB
    # ── Force sql_lookup for any data-seeking question ─────────
    data_words = [
        "who", "which", "how many", "how much", "what did", "what has",
        "what have", "where", "when", "top", "most", "least", "lowest",
        "highest", "list", "show", "give", "compare", "bought", "purchased",
        "returned", "sold", "customer", "product", "region", "category"
    ]
    if any(w in q for w in data_words):
        return "sql_lookup"

def handle_similar_customers(customer_name):
    v = customer_name.replace("'", "''")
    sql = f"""
WITH target_products AS (
    SELECT DISTINCT product_name
    FROM orders WHERE LOWER(customer_name) = LOWER('{v}')
),
target_total AS (
    SELECT COUNT(DISTINCT product_name) AS total_products
    FROM orders WHERE LOWER(customer_name) = LOWER('{v}')
),
candidate_totals AS (
    SELECT o.customer_name, COUNT(DISTINCT o.product_name) AS total_owned
    FROM orders o
    WHERE LOWER(o.customer_name) != LOWER('{v}')
    GROUP BY o.customer_name
),
shared AS (
    SELECT o.customer_name,
        COUNT(DISTINCT o.product_name) AS shared_products
    FROM orders o
    WHERE LOWER(o.customer_name) != LOWER('{v}')
    AND o.product_name IN (SELECT product_name FROM target_products)
    GROUP BY o.customer_name
)
SELECT c.customer_name, c.age, c.sex, c.prime_status,
    MAX(o.region) AS region,
    s.shared_products,
    ct.total_owned,
    ROUND(s.shared_products::numeric / NULLIF(ct.total_owned, 0), 2) AS overlap_ratio
FROM shared s
JOIN candidate_totals ct ON s.customer_name = ct.customer_name
JOIN customers c ON c.customer_name = s.customer_name
JOIN orders o ON o.customer_name = s.customer_name
GROUP BY c.customer_name, c.age, c.sex, c.prime_status, s.shared_products, ct.total_owned
ORDER BY s.shared_products DESC, overlap_ratio DESC
LIMIT 5
"""
    try:
        result = run_query(sql, "similar customers")
        if result is None or result.empty:
            return f"No similar customers found for {customer_name} based on shared product purchases."
        lines = [f"Here are customers most similar to {customer_name}, based purely on overlapping product purchases:\n"]
        for i, row in result.iterrows():
            overlap_pct = float(row.get('overlap_ratio', 0)) * 100
            lines.append(f"{i+1}. {row['customer_name']} — {row.get('shared_products', 0)} shared products ({overlap_pct:.0f}% of their purchases overlap with {customer_name}'s)")
        return "\n".join(lines)
    except Exception as e:
        print(f"  ❌ Similar customers error: {e}")
        return f"Could not find similar customers for {customer_name}."
    
def handle_contextual(question, ctx):
    """
    Handles follow-up questions with 3 sub-intents:
    - sql_lookup  : re-queries DB using previous result as reference
    - subjective  : interprets previous answer with Mistral
    - text        : pure text response using previous answer
    """
    # Filter out summary pseudo-entries
    real_ctx = [ex for ex in ctx if ex.get("question") != "__summary__"]
    summary_entry = next((ex for ex in ctx if ex.get("question") == "__summary__"), None)

    if not real_ctx:
        return handle_generic(question)

    last = real_ctx[-1]
    prev_question = last['question']
    prev_answer = last['answer']
    prev_result = last.get('result', '')

    # Pure acknowledgements — no action needed
    q_stripped = question.lower().strip().rstrip("!.")
    ack_words = {"okay", "ok", "got it", "understood", "sure", "alright",
                 "fine", "noted", "great", "nice", "cool", "good", "perfect",
                 "i see", "makes sense", "wow", "oh", "ah", "hmm", "right"}
    if q_stripped in ack_words:
        return handle_generic(question)

    # Route recommendations to recommendation handler
    # Route recommendations to recommendation handler
    if any(w in question.lower() for w in ['recommend', 'suggestion', 'suggest', 'buy next', 'will buy', 'purchase next', 'next purchase']):
        entity = None
        for ex in reversed(real_ctx):
            carried = ex.get('resolved_entity', {})
            if carried.get('customer_name'):
                entity = carried['customer_name']
                break
        if entity:
            return handle_customer_recommendation(question, entity)
        else:
            answer, _ = handle_recommendation(question, real_ctx)
            context.append({"question": question, "sql": "", "answer": answer, "result": ""})
            if len(context) > 10: context.pop(0)
            return answer

    # ── QUERY REWRITING — resolve vague follow-up into standalone question ──
    prev_result_preview = ""
    if prev_result:
        try:
            df_prev = pd.read_csv(io.StringIO(prev_result))
            prev_result_preview = df_prev.head(10).to_csv(index=False)
        except Exception:
            prev_result_preview = prev_result[:600]

    rewrite_prompt = f"""A user is chatting with a sales analytics bot about a Superstore database.

Conversation so far:
Q: {prev_question}
A: {prev_answer[:300]}

Previous data returned:
{prev_result_preview}

User follow-up: "{question}"

Rewrite the follow-up as a complete standalone question that:
- Replaces all pronouns (it, they, him, her, them, that, this) with the actual names/values from the previous data
- Resolves vague references like "the lowest one", "the first one", "that region" using the actual value from the previous data
- Is self-contained with no references to previous context
- Keeps the same intent as the original follow-up
- If the question asks about similar customers, keep it simple: just ask "Find customers similar to [name] based on age, sex and prime status" — do not mention products
- If the question asks about cities, states, or locations in a region, keep it simple: just ask "What cities are in the [region name] region?" — do not include any sales numbers or metrics

Output ONLY the rewritten question. No explanation, no quotes."""

    rewritten = call_model(rewrite_prompt, model=FAST_MODEL, temperature=0.1, num_ctx=2048, timeout=45)
    rewritten = rewritten.strip().strip('"').strip("'")
    if not rewritten or len(rewritten) < 5:
        rewritten = question
    print(f"  ✏️  Rewritten question: {rewritten}")

    sub = _classify_contextual_subintent(rewritten, prev_question, prev_answer, prev_result)
    print(f"  🔀 Contextual sub-intent: {sub}")

    # ── SQL LOOKUP ─────────────────────────────────────────────
    if sub == "sql_lookup":
        # Step 1: Resolve entity from REWRITTEN question
        resolved = _resolve_reference(rewritten, prev_result)
        resolved_row = resolved.get('row', {})

        # Step 2: Direct name match in question
        # Step 2: Direct name match in REWRITTEN question (takes priority over carried entity)
        try:
            with engine.connect() as conn:
                customers = conn.execute(text("SELECT customer_name FROM customers")).fetchall()
                customer_names = [row[0] for row in customers]
            matched_names = [name for name in customer_names if name.lower() in rewritten.lower()]
            if matched_names:
                if len(matched_names) == 1:
                    resolved_row = {'customer_name': matched_names[0]}
                else:
                    resolved_row = {'customer_name': matched_names}
                print(f"  👤 Direct name match in rewritten: {matched_names}")
        except Exception:
            pass

        # Step 3: Carry forward from previous exchanges
        # Check if rewritten question mentions a product directly
        try:
            with engine.connect() as conn:
                products = conn.execute(text("SELECT DISTINCT product_name FROM orders")).fetchall()
                product_names = [row[0] for row in products]
            for pname in product_names:
                if pname.lower() in rewritten.lower():
                    resolved_row = {'product_name': pname}
                    print(f"  📦 Direct product match in rewritten: {pname}")
                    break
        except Exception:
            pass

        # Step 3: Carry forward from previous exchanges
        if not resolved_row:
            for ex in reversed(real_ctx):
                carried = ex.get('resolved_entity', {})
                if carried:
                    resolved_row = carried
                    print(f"  🔄 Carrying forward resolved_entity")
                    break
                raw = ex.get('result', '')
                if raw:
                    parsed = _resolve_reference(question, raw)
                    if parsed.get('row'):
                        resolved_row = parsed['row']
                        print(f"  🔄 Carrying forward entity from parsed result")
                        break

        # Step 4: Detect entity type and value
        entity_type = None
        entity_value = None
        for key in ['product_name', 'category', 'sub_category', 'region',
                    'state', 'city', 'customer_name', 'segment', 'ship_mode']:
            if key in resolved_row:
                entity_type = key
                entity_value = resolved_row[key] if isinstance(resolved_row[key], list) else str(resolved_row[key]).strip()
                break

        print(f"  🔍 Resolved entity: {entity_type} = {entity_value}")

        # If question asks about products but entity is a product, look for region instead
        q_lower = question.lower()
        if entity_type == 'product_name' and any(w in q_lower for w in ['product', 'products', 'selling', 'sold', 'top selling', 'best selling']):
            for ex in reversed(real_ctx):
                carried = ex.get('resolved_entity', {})
                if carried.get('region'):
                    entity_type = 'region'
                    entity_value = carried['region']
                    print(f"  🔄 Switched to region entity: {entity_value}")
                    break

        # Step 5: Re-resolve if no entity found
        if not entity_type and not entity_value:
            q_lower = question.lower()
            if any(w in q_lower for w in ['least', 'lowest', 'minimum', 'worst', 'bottom']):
                try:
                    df = pd.read_csv(io.StringIO(prev_result))
                    df.columns = [str(c).strip().lower().replace(' ', '_') for c in df.columns]
                    num_cols = [c for c in df.select_dtypes(include='number').columns
                               if not any(x in c.lower() for x in ['index', 'rank', 'row', 'id'])]
                    if num_cols:
                        col = num_cols[-1]
                        row = df.loc[df[col].idxmin()]
                        resolved_row = row.to_dict()
                        for key in ['product_name', 'category', 'region', 'customer_name']:
                            if key in resolved_row:
                                entity_type = key
                                entity_value = str(resolved_row[key]).strip()
                                print(f"  🔍 Re-resolved: {entity_type} = {entity_value}")
                                break
                except Exception as e:
                    print(f"  ⚠️ Re-resolve error: {e}")

        # Step 6: Detect topic from FULL context history
        all_prev = " ".join(ex.get('question', '') for ex in real_ctx).lower()
        is_about_returns = any(w in all_prev for w in ['return', 'returned', 'returns'])
        is_about_sales   = any(w in all_prev for w in ['sales', 'revenue', 'sold'])
        is_about_profit  = any(w in all_prev for w in ['profit', 'profitable'])

        # Step 7: Detect follow-up intent
        q = question.lower()
        is_asking_who = (
            any(w in q for w in ['who', 'which customer', 'buyer', 'bought', 'purchase']) or
            ('most' in q and any(w in q for w in ['customer', 'who', 'bought', 'buyer', 'purchased']))
        )
        is_asking_region    = any(w in q for w in ['region', 'where', 'location', 'area'])
        is_asking_category  = any(w in q for w in ['category', 'type', 'kind', 'drives'])
        is_asking_compare   = any(w in q for w in ['compare', 'vs', 'versus', 'best', 'worst', 'all region'])
        is_asking_products  = any(w in q for w in ['product', 'products', 'item', 'items', 'sell', 'sold', 'selling'])
        is_asking_purchases = any(w in q for w in ['buy', 'bought', 'purchase', 'purchased', 'order', 'ordered', 'what did', 'what has', 'what have'])
        is_asking_returns_q = any(w in q for w in ['return', 'returned', 'returns', 'give back', 'sent back'])
        is_asking_similar = any(w in q for w in ['similar', 'like him', 'like her', 'like them', 'same type', 'other customers'])

        sql = None

        if is_asking_similar:
            v_name = entity_value if isinstance(entity_value, str) else (entity_value[0] if entity_value else None)
            if not v_name:
                for ex in reversed(real_ctx):
                    carried = ex.get('resolved_entity', {})
                    if carried.get('customer_name'):
                        v_name = carried['customer_name']
                        break
            if v_name:
                answer = handle_similar_customers(v_name)
                context.append({"question": question, "sql": "", "answer": answer, "result": "", "resolved_entity": {"customer_name": v_name}})
                if len(context) > 10: context.pop(0)
                return answer

        if entity_type == 'product_name' and entity_value:
            v = entity_value.replace("'", "''")
            if is_asking_who:
                if is_about_returns:
                    sql = f"""SELECT o.customer_name,
                        COUNT(DISTINCT r.return_id) AS return_count,
                        ROUND(SUM(o.sales)::numeric,2) AS total_sales
                        FROM orders o JOIN returns r ON o.order_id = r.order_id
                        WHERE o.product_name = '{v}'
                        GROUP BY o.customer_name ORDER BY return_count DESC LIMIT 10"""
                else:
                    sql = f"""SELECT o.customer_name,
                        COUNT(DISTINCT o.order_id) AS order_count,
                        ROUND(SUM(o.sales)::numeric,2) AS total_sales
                        FROM orders o WHERE o.product_name = '{v}'
                        GROUP BY o.customer_name ORDER BY order_count DESC LIMIT 10"""
            elif is_asking_region:
                if is_about_returns:
                    sql = f"""SELECT o.region,
                        COUNT(DISTINCT r.return_id) AS return_count,
                        ROUND(SUM(o.sales)::numeric,2) AS total_sales
                        FROM orders o JOIN returns r ON o.order_id = r.order_id
                        WHERE o.product_name = '{v}'
                        GROUP BY o.region ORDER BY return_count DESC"""
                else:
                    sql = f"""SELECT o.region,
                        COUNT(DISTINCT o.order_id) AS order_count,
                        ROUND(SUM(o.sales)::numeric,2) AS total_sales
                        FROM orders o WHERE o.product_name = '{v}'
                        GROUP BY o.region ORDER BY order_count DESC"""
            elif is_asking_category:
                sql = f"""SELECT o.category, o.sub_category,
                    COUNT(DISTINCT o.order_id) AS order_count,
                    ROUND(SUM(o.sales)::numeric,2) AS total_sales,
                    ROUND(SUM(o.profit)::numeric,2) AS total_profit
                    FROM orders o WHERE o.product_name = '{v}'
                    GROUP BY o.category, o.sub_category ORDER BY order_count DESC"""
            else:
                if is_about_returns:
                    sql = f"""SELECT o.product_name, o.category, o.sub_category,
                        COUNT(DISTINCT r.return_id) AS total_returns,
                        COUNT(DISTINCT o.order_id) AS total_orders,
                        ROUND(SUM(o.sales)::numeric,2) AS total_sales,
                        ROUND(SUM(o.profit)::numeric,2) AS total_profit,
                        ROUND(AVG(o.discount)::numeric*100,1) AS avg_discount_pct                       
                        FROM orders o LEFT JOIN returns r ON o.order_id = r.order_id
                        WHERE o.product_name = '{v}'
                        GROUP BY o.product_name, o.category, o.sub_category"""
                else:
                    sql = f"""SELECT o.product_name, o.category, o.sub_category,
                        COUNT(DISTINCT o.order_id) AS total_orders,
                        ROUND(SUM(o.sales)::numeric,2) AS total_sales,
                        ROUND(SUM(o.profit)::numeric,2) AS total_profit,
                        ROUND(AVG(o.discount)::numeric*100,1) AS avg_discount_pct
                        FROM orders o WHERE o.product_name = '{v}'
                        GROUP BY o.product_name, o.category, o.sub_category"""

        elif entity_type == 'customer_name' and entity_value:
            # Handle both single name and list of names
            if isinstance(entity_value, list):
                names_sql = ", ".join([f"'{n.replace(chr(39), chr(39)*2)}'" for n in entity_value])
                v_filter = f"o.customer_name IN ({names_sql})"
            else:
                v = entity_value.replace("'", "''")
                v_filter = f"o.customer_name = '{v}'"
            if is_asking_similar:
                v_name = entity_value if isinstance(entity_value, str) else entity_value[0]
                v_name = v_name.replace("'", "''")
                sql = f"""SELECT c.customer_name, c.age, c.sex, c.prime_status,
                    c.payment_mode, c.channel, c.device,
                    COUNT(DISTINCT o.order_id) AS total_orders,
                    ROUND(SUM(o.sales)::numeric,2) AS total_sales,
                    ROUND(SUM(o.profit)::numeric,2) AS total_profit,
                    COUNT(DISTINCT r.return_id) AS total_returns,
                    MAX(o.region) AS region, MAX(o.segment) AS segment
                    FROM customers c
                    LEFT JOIN orders o ON c.customer_name = o.customer_name
                    LEFT JOIN returns r ON o.order_id = r.order_id
                    WHERE c.sex = (SELECT sex FROM customers WHERE customer_name = '{v_name}')
                    AND c.age BETWEEN (SELECT age FROM customers WHERE customer_name = '{v_name}') - 5
                        AND (SELECT age FROM customers WHERE customer_name = '{v_name}') + 5
                    AND c.prime_status = (SELECT prime_status FROM customers WHERE customer_name = '{v_name}')
                    AND c.customer_name != '{v_name}'
                    GROUP BY c.customer_name, c.age, c.sex, c.prime_status,
                    c.payment_mode, c.channel, c.device
                    ORDER BY total_sales DESC LIMIT 5"""
            elif is_asking_returns_q:
                sql = f"""SELECT o.product_name, o.category,
                    COUNT(DISTINCT r.return_id) AS return_count,
                    ROUND(SUM(o.sales)::numeric,2) AS total_sales
                    FROM orders o JOIN returns r ON o.order_id = r.order_id
                    WHERE {v_filter}
                    GROUP BY o.product_name, o.category
                    ORDER BY return_count DESC"""
            elif is_asking_purchases:
                sql = f"""SELECT o.customer_name, o.product_name, o.category, o.sub_category,
                    COUNT(DISTINCT o.order_id) AS order_count,
                    ROUND(SUM(o.sales)::numeric,2) AS total_sales
                    FROM orders o WHERE {v_filter}
                    GROUP BY o.customer_name, o.product_name, o.category, o.sub_category
                    ORDER BY order_count DESC LIMIT 10"""
            elif is_asking_region:
                sql = f"""SELECT o.customer_name, o.region, o.state, o.city,
                    COUNT(DISTINCT o.order_id) AS order_count,
                    ROUND(SUM(o.sales)::numeric,2) AS total_sales
                    FROM orders o WHERE {v_filter}
                    GROUP BY o.customer_name, o.region, o.state, o.city ORDER BY order_count DESC"""
            elif is_asking_category:
                sql = f"""SELECT o.customer_name, o.category, o.sub_category,
                    COUNT(DISTINCT o.order_id) AS order_count,
                    ROUND(SUM(o.sales)::numeric,2) AS total_sales
                    FROM orders o WHERE {v_filter}
                    GROUP BY o.customer_name, o.category, o.sub_category ORDER BY order_count DESC"""
            else:
                sql = f"""SELECT c.customer_name, c.age, c.sex, c.prime_status,
                    c.payment_mode, c.channel, c.device,
                    COUNT(DISTINCT o.order_id) AS total_orders,
                    ROUND(SUM(o.sales)::numeric,2) AS total_sales,
                    ROUND(SUM(o.profit)::numeric,2) AS total_profit,
                    COUNT(DISTINCT r.return_id) AS total_returns,
                    MAX(o.region) AS region, MAX(o.segment) AS segment
                    FROM customers c
                    LEFT JOIN orders o ON c.customer_name = o.customer_name
                    LEFT JOIN returns r ON o.order_id = r.order_id
                    WHERE {v_filter.replace('o.customer_name', 'c.customer_name')}
                    GROUP BY c.customer_name, c.age, c.sex, c.prime_status,
                    c.payment_mode, c.channel, c.device"""

        elif entity_type == 'region' and entity_value:
            v = entity_value.replace("'", "''")
            if any(w in q for w in ['city', 'cities', 'town', 'towns', 'state', 'states']):
                sql = f"""SELECT o.city, o.state,
                    COUNT(DISTINCT o.order_id) AS order_count,
                    ROUND(SUM(o.sales)::numeric,2) AS total_sales
                    FROM orders o WHERE o.region = '{v}'
                    GROUP BY o.city, o.state ORDER BY total_sales DESC"""
                
            elif is_asking_products:
                        sql = f"""SELECT o.product_name,
                    COUNT(DISTINCT o.order_id) AS order_count,
                    ROUND(SUM(o.sales)::numeric,2) AS total_sales
                    FROM orders o WHERE o.region = '{v}'
                    GROUP BY o.product_name ORDER BY order_count DESC LIMIT 10"""
                
            elif is_asking_who:
                sql = f"""SELECT o.customer_name,
                    COUNT(DISTINCT o.order_id) AS order_count,
                    ROUND(SUM(o.sales)::numeric,2) AS total_sales
                    FROM orders o WHERE o.region = '{v}'
                    GROUP BY o.customer_name ORDER BY order_count DESC LIMIT 10"""
            elif is_asking_category:
                if is_about_returns:
                    sql = f"""SELECT o.category,
                        COUNT(DISTINCT r.return_id) AS return_count,
                        ROUND(SUM(o.sales)::numeric,2) AS total_sales
                        FROM orders o JOIN returns r ON o.order_id = r.order_id
                        WHERE o.region = '{v}'
                        GROUP BY o.category ORDER BY return_count DESC"""
                else:
                    sql = f"""SELECT o.category,
                        COUNT(DISTINCT o.order_id) AS order_count,
                        ROUND(SUM(o.sales)::numeric,2) AS total_sales,
                        ROUND(SUM(o.profit)::numeric,2) AS total_profit
                        FROM orders o WHERE o.region = '{v}'
                        GROUP BY o.category ORDER BY total_sales DESC"""
            elif is_asking_compare:
                sql = f"""SELECT o.region,
                    COUNT(DISTINCT o.order_id) AS total_orders,
                    ROUND(SUM(o.sales)::numeric,2) AS total_sales,
                    ROUND(SUM(o.profit)::numeric,2) AS total_profit,
                    COUNT(DISTINCT r.return_id) AS total_returns
                    FROM orders o LEFT JOIN returns r ON o.order_id = r.order_id
                    GROUP BY o.region ORDER BY total_sales DESC"""
            else:
                sql = f"""SELECT o.region,
                    COUNT(DISTINCT o.order_id) AS total_orders,
                    ROUND(SUM(o.sales)::numeric,2) AS total_sales,
                    ROUND(SUM(o.profit)::numeric,2) AS total_profit,
                    COUNT(DISTINCT r.return_id) AS total_returns
                    FROM orders o LEFT JOIN returns r ON o.order_id = r.order_id
                    WHERE o.region = '{v}'
                    GROUP BY o.region"""

        elif entity_type == 'category' and entity_value:
            v = entity_value.replace("'", "''")
            if is_about_returns:
                sql = f"""SELECT o.sub_category,
                    COUNT(DISTINCT r.return_id) AS return_count,
                    ROUND(SUM(o.sales)::numeric,2) AS total_sales
                    FROM orders o JOIN returns r ON o.order_id = r.order_id
                    WHERE o.category = '{v}'
                    GROUP BY o.sub_category ORDER BY return_count DESC"""
            else:
                sql = f"""SELECT o.sub_category,
                    COUNT(DISTINCT o.order_id) AS order_count,
                    ROUND(SUM(o.sales)::numeric,2) AS total_sales,
                    ROUND(SUM(o.profit)::numeric,2) AS total_profit
                    FROM orders o WHERE o.category = '{v}'
                    GROUP BY o.sub_category ORDER BY total_sales DESC"""

        # Fallback to Gemma3 — use rewritten question for better context
        if not sql:
            print(f"  ⚠️ No template matched — falling back to Gemma3")
            enriched = f"""{rewritten}
CONTEXT: The conversation has been about '{all_prev[:200]}'.
Entity: {entity_type} = '{entity_value}'
Use this exact value in WHERE clause. Do NOT query globally.
Use COUNT(DISTINCT) for all counts.
For returns always JOIN orders o JOIN returns r ON o.order_id = r.order_id. Never query returns table alone.
If asked about cities or states in a region, SELECT o.city, o.state, COUNT(DISTINCT o.order_id) AS order_count, ROUND(SUM(o.sales)::numeric,2) AS total_sales FROM orders o WHERE o.region = 'region_name' GROUP BY o.city, o.state ORDER BY total_sales DESC."""
            date_ctx = inject_date_into_prompt(rewritten)
            sql = text_to_sql(enriched, date_ctx)

        print(f"  ├─ Contextual SQL: {sql}")

        if not sql or "UNRELATED" in sql.upper() or not sql.strip().upper().startswith(("SELECT", "WITH")):
            sub = "text"
        else:
            result = run_query(sql, question)
            if result is None or result.empty:
                sub = "text"
            else:
                if needs_visualisation(question):
                    create_visualisation(result, question)
                answer = summarise(question, result)
               # Update resolved_entity — prefer original resolved_row over new result
                try:
                    new_entity = resolved_row if resolved_row else {}
                    if not new_entity:
                        result_cols = result.columns.tolist()
                        for key in ['customer_name', 'product_name', 'region', 'category']:
                            if key in result_cols:
                                new_entity = {key: str(result.iloc[0][key])}
                                break
                except Exception:
                    new_entity = {}

                context.append({
                    "question": question,
                    "sql": sql,
                    "answer": answer,
                    "result": result.to_csv(index=False),
                    "resolved_entity": new_entity
                })
                if len(context) > 10:
                    context.pop(0)
                return answer
    # ── SUBJECTIVE ─────────────────────────────────────────────
    if sub == "subjective":
        summary_text = summary_entry['answer'] if summary_entry else ""
        prompt = f"""You are an expert Sales Analytics Consultant for Superstore.

{f'Conversation summary so far: {summary_text}' if summary_text else ''}

Previous question: '{prev_question}'
Previous answer: {prev_answer[:600]}
Raw data: {prev_result[:400]}

User now asks: '{question}'

Provide a short analytical interpretation (2-3 sentences max).
Answer the 'why' or 'what this means' using the data above.
Be specific — reference actual numbers/names from the data. Never start with 'I'."""
        answer = call_model(prompt, model=FAST_MODEL, temperature=0.5, num_ctx=1024, timeout=45)
        context.append({
            "question": question,
            "sql": sql,
            "answer": answer,
            "result": result.to_string(),
            "resolved_entity": {"customer_name": customer_name}
        })
        if len(context) > 10:
            context.pop(0)
        return answer

    # ── TEXT (default) ─────────────────────────────────────────
    summary_text = summary_entry['answer'] if summary_entry else ""
    prompt = f"""You are a Sales Analytics Assistant for Superstore.

{f'Conversation summary so far: {summary_text}' if summary_text else ''}

Previous question: '{prev_question}'
Previous answer:
{prev_answer}

User now asks: '{question}'

STRICT RULES:
- ONLY use data explicitly listed in the previous answer above
- Do NOT invent products, names, or numbers not in the previous answer
- If asked to bundle, group ONLY items from the previous answer by their category
- If the previous answer has no relevant data, say so clearly
- Never start with 'I'
- Be concise"""

    answer = call_model(prompt, model=FAST_MODEL, temperature=0.4, num_ctx=1024, timeout=45)
    if not answer.strip():
        answer = "I couldn't generate a response. Please try rephrasing."
    context.append({
        "question": question,
        "sql": "",
        "answer": answer,
        "result": ""
    })
    if len(context) > 10:
        context.pop(0)
    return answer


def handle_recommendation(question, ctx):
    """
    Handles recommendation questions like Amazon.
    Top 2 products from each of 5 categories = 10 products total.
    Just list the products, and after listing give an overall one liner explanantion, not next to the products
    """
    # If question references previous answer, redirect to contextual
    context_refs = [
        "above", "those", "from the list", "suggested", "the ones", "bundle",
        "from that", "of those", "these products", "the above", "from above",
        "mentioned", "listed", "previous", "from the above"
    ]
    if ctx and any(ref in question.lower() for ref in context_refs):
        return handle_contextual(question, ctx), ctx[-1].get('result', '')

    # Build WHERE clause from question
    q = question.lower()
    filters = []

    age_match = re.search(r'(\d+)\s*[-–to]+\s*(\d+)', q)
    if age_match:
        filters.append(f"c.age BETWEEN {age_match.group(1)} AND {age_match.group(2)}")
    elif 'young' in q:
        filters.append("c.age BETWEEN 21 AND 30")
    elif 'middle aged' in q or 'middle-aged' in q:
        filters.append("c.age BETWEEN 41 AND 50")
    elif 'senior' in q or 'elderly' in q:
        filters.append("c.age > 50")
    elif 'teenager' in q or 'teen' in q:
        filters.append("c.age BETWEEN 11 AND 20")

    if 'non-prime' in q or 'non prime' in q:
        filters.append("c.prime_status = 'Non-Prime'")
    elif 'prime' in q:
        filters.append("c.prime_status = 'Prime'")

    if 'mobile' in q and 'phone' not in q:
        filters.append("c.channel = 'Mobile'")
    elif 'web' in q:
        filters.append("c.channel = 'Web'")

    if 'android' in q:
        filters.append("c.device = 'Android'")
    elif 'ios' in q:
        filters.append("c.device = 'iOS'")

    where_clause = "WHERE " + " AND ".join(filters) if filters else ""

    # Hardcoded SQL — top 2 per category guaranteed
    sql = f"""
WITH ranked AS (
    SELECT
        o.product_name,
        o.category,
        COUNT(DISTINCT o.order_id) AS purchase_count,
        ROW_NUMBER() OVER (
            PARTITION BY o.category
            ORDER BY COUNT(DISTINCT o.order_id) DESC
        ) AS rn
    FROM orders o
    JOIN customers c ON o.customer_name = c.customer_name
    {where_clause}
    GROUP BY o.product_name, o.category
)
SELECT product_name, category, purchase_count
FROM ranked
WHERE rn <= 2
ORDER BY category, rn
"""

    relevant_data_str = ""
    product_list_from_db = []

    try:
        result = run_query(sql, question)
        if result is not None and not result.empty and 'product_name' in result.columns:
            relevant_data_str = result.to_string()
            product_list_from_db = result['product_name'].tolist()
    except Exception as e:
        print(f"  ⚠️ Recommendation SQL error: {e}")

    if product_list_from_db:
        if age_match:
            group_desc = f"customers aged {age_match.group(1)}-{age_match.group(2)}"
        elif 'non-prime' in q or 'non prime' in q:
            group_desc = "Non-Prime members"
        elif 'prime' in q:
            group_desc = "Prime members"
        elif 'android' in q:
            group_desc = "Android users"
        elif 'ios' in q:
            group_desc = "iOS users"
        elif 'mobile' in q and 'phone' not in q:
            group_desc = "Mobile channel customers"
        elif 'web' in q:
            group_desc = "Web channel customers"
        else:
            group_desc = "all customers in our database"

        header = f"Here are the recommended products, which others from {group_desc} have also bought:\n"
        numbered = "\n".join([f"{i+1}. {p}" for i, p in enumerate(product_list_from_db)])
        why_prompt = f"""In one sentence under 20 words, explain why these specific products suit {group_desc}, based only on the fact that they are top sellers among that group.
Products: {', '.join(product_list_from_db[:5])}
Start with: These are the top products purchased by"""
        why = call_model(why_prompt, model=FAST_MODEL, temperature=0.1, num_ctx=256, timeout=30)
        why = why.strip().split('\n')[0]
        answer = f"{header}\n{numbered}\n\n{why}"
        return answer, relevant_data_str

    # No data found for these filters — give an honest answer, never invent products
    return (f"Couldn't find enough purchase data for that specific group. "
            f"Try a broader request like 'recommend products for prime members' "
            f"or 'recommend products for age 20-30'."), ""

def handle_customer_recommendation(question, customer_name):
    try:
        sql = f"""
WITH customer_products AS (
    SELECT DISTINCT o.product_name, o.category
    FROM orders o
    WHERE LOWER(o.customer_name) = LOWER('{customer_name}')
),
similar_customers AS (
    SELECT o.customer_name,
        COUNT(DISTINCT o.product_name) AS shared_products
    FROM orders o
    WHERE LOWER(o.customer_name) != LOWER('{customer_name}')
    AND o.product_name IN (SELECT product_name FROM customer_products)
    GROUP BY o.customer_name
    HAVING COUNT(DISTINCT o.product_name) >= 1
    ORDER BY shared_products DESC
    LIMIT 30
),
recommended_products AS (
    SELECT 
        o.product_name,
        o.category,
        COUNT(DISTINCT o.customer_name) AS bought_by,
        ROUND(AVG(o.sales)::numeric, 2) AS avg_price
    FROM orders o
    WHERE o.customer_name IN (SELECT customer_name FROM similar_customers)
    AND o.product_name NOT IN (SELECT product_name FROM customer_products)
    GROUP BY o.product_name, o.category
    ORDER BY bought_by DESC
    LIMIT 10
)
SELECT product_name, category, bought_by AS similar_customers_bought, avg_price
FROM recommended_products
ORDER BY bought_by DESC
"""
        result = run_query(sql, question)

        if result is None or result.empty:
            return f"Couldn't find recommendations for {customer_name}. They may not have enough purchase history yet, or similar customers haven't bought anything new."

        top_10 = result.head(10)['product_name'].tolist()
        top_3 = result.head(3)['product_name'].tolist()

        is_next_purchase_q = any(w in question.lower() for w in ['buy next', 'will buy', 'purchase next', 'next purchase'])

        if is_next_purchase_q:
            lines = [f"{customer_name}'s shopping twins have spoken — based on what customers with the most overlapping purchases buy next, here's where the smart money is:\n"]
            for i, p in enumerate(top_3, 1):
                lines.append(f"{i}. {p}")
            lines.append(f"\nThese are the products most frequently picked up by customers whose buying habits closely mirror {customer_name}'s — a strong signal for what comes next.")
            answer = "\n".join(lines)
        else:
            lines = [f"We looked at customers whose shopping carts overlap most with {customer_name}'s, and pulled out the 10 products they keep buying that {customer_name} hasn't tried yet:\n"]
            for i, p in enumerate(top_10, 1):
                lines.append(f"{i}. {p}")
            lines.append(f"\nThink of this as a curated shortlist built from real buying patterns, not guesswork — every item here has already proven popular with shoppers most like {customer_name}.")
            answer = "\n".join(lines)

        context.append({
            "question": question,
            "sql": sql,
            "answer": answer,
            "result": result.to_string(),
            "resolved_entity": {"customer_name": customer_name}
        })
        if len(context) > 10:
            context.pop(0)

        return answer

    except Exception as e:
        print(f"  ❌ Recommendation error: {e}")
        return f"Could not generate recommendations for {customer_name}. Please try again."

def handle_subjective_sql(question, ctx):
    """
    Handles questions needing data AND interpretation.
    Gemma3 gets the data, Mistral interprets it.
    """
    # Step 1: Get data with Gemma3
    date_ctx = inject_date_into_prompt(question)
    sql = text_to_sql(question, date_ctx)
    print(f"  ├─ SQL: {sql}")

    data_str = ""
    result = None

    if sql and not "UNRELATED" in sql.upper() and sql.strip().upper().startswith(("SELECT", "WITH")):
        result = run_query(sql, question)
        if result is not None and not result.empty:
            data_str = result.head(20).to_string()
            if needs_visualisation(question):
                create_visualisation(result, question)

    # Step 2: Interpret with Mistral
    context_str = ""
    if ctx:
        context_str = f"\nPrevious context: {ctx[-1]['question']} → {ctx[-1]['answer']}"

    prompt = f"""You are an expert Sales Analytics Consultant for Superstore.

The user asks: '{question}'
Database results: {data_str if data_str else 'No specific data found'}
{context_str}

Provide an insightful interpretation and explanation. Answer the 'why' behind the numbers.
Be analytical and business-focused. Max 2 sentences. Never start with 'I'."""

    answer = call_model(prompt, model=FAST_MODEL, temperature=0.6,
                       num_ctx=4096, timeout=120)

    # Save to context
    if result is not None and not result.empty:
        context.append({
            "question": question,
            "sql": sql if sql else "",
            "answer": answer,
            "result": data_str
        })
        if len(context) > 10:
            context.pop(0)

    return answer


def handle_sql(question):
    """
    Handles data questions: generates SQL with Gemma3, summarises with Mistral.
    """
    global latest_chart_buf

    date_ctx = inject_date_into_prompt(question)
    sql = text_to_sql(question, date_ctx)
    print(f"  ├─ SQL: {sql}")

    if "UNRELATED" in sql.upper():
        # Check if question contains a name — retry with explicit customer hint
        name_hint_phrases = ["details of", "tell me about", "info on", "all details"]
        if any(phrase in question.lower() for phrase in name_hint_phrases):
            retry_q = f"{question} — look up this person in the customers table and orders table"
            sql = text_to_sql(retry_q, date_ctx)
            if "UNRELATED" in sql.upper() or not sql.strip().upper().startswith(("SELECT", "WITH")):
                return "Couldn't find details for that person. Please check the name and try again."
        elif context and len(question.split()) <= 6:
            retry_q = f"{question} (referring to previous answer about {context[-1]['question']})"
            sql = text_to_sql(retry_q, date_ctx)
            if "UNRELATED" in sql.upper() or not sql.strip().upper().startswith(("SELECT", "WITH")):
                return "That question is outside my expertise! I'm your Superstore Sales Analytics Assistant — ask me about sales, profits, customers, products, returns, regions or trends and I'll have the answer ready for you."
        else:
           return "That question is outside my expertise! I'm your Superstore Sales Analytics Assistant — ask me about sales, profits, customers, products, returns, regions or trends and I'll have the answer ready for you."
    if not sql.strip().upper().startswith(("SELECT", "WITH")):
        return "That question is outside my expertise! I'm your Superstore Sales Analytics Assistant — ask me about sales, profits, customers, products, returns, regions or trends and I'll have the answer ready for you."

    result = run_query(sql, question)

    # For product listing questions — return clean list, skip summarise
    listing_phrases = [
        "what products", "which products", "list products",
        "show products", "what items", "which items"
    ]
    if any(phrase in question.lower() for phrase in listing_phrases):
        if 'product_name' in result.columns:
            products = result['product_name'].tolist()
            answer = "\n".join([f"{i+1}. {p}" for i, p in enumerate(products)])
            context.append({"question": question, "sql": sql, "answer": answer, "result": result.to_string()})
            if len(context) > 10: context.pop(0)
            return answer

    if result is None or result.empty:
       return "That question is outside my expertise! I'm your Superstore Sales Analytics Assistant — ask me about sales, profits, customers, products, returns, regions or trends and I'll have the answer ready for you."
    detail_phrases = ["all orders", "all data", "give me all", "give me details", "show me details"]
    is_person_query = any(phrase in question.lower() for phrase in 
                        ["of ", "for ", "about ", "details of", "data of", "data for"])

    if any(phrase in question.lower() for phrase in detail_phrases) and not is_person_query:
        answer = result.to_string(index=False)
        context.append({"question": question, "sql": sql, "answer": answer, "result": result.to_string()})
        if len(context) > 10: context.pop(0)
        return answer

    if needs_visualisation(question):
        create_visualisation(result, question)

    answer = summarise(question, result)

    context.append({
        "question": question,
        "sql": sql,
        "answer": answer,
        "result": result.to_string() if result is not None else ""
    })
    if len(context) > 10:
        context.pop(0)

    return answer


# ══════════════════════════════════════════════════════════════
# SQL GENERATION — text_to_sql()
# Uses Gemma3 (SQL_MODEL) — best model for SQL
# ══════════════════════════════════════════════════════════════

def text_to_sql(question, date_ctx=None):
    context_str = ""
    summary_entry = next((ex for ex in context if ex.get("question") == "__summary__"), None)
    real_ctx_sql = [ex for ex in context if ex.get("question") != "__summary__"]
    if summary_entry:
        context_str += f"\nCONVERSATION SUMMARY:\n{summary_entry['answer']}\n"
    if real_ctx_sql:
        context_str += "\n\nRECENT EXCHANGES:\n"
        for exchange in real_ctx_sql[-3:]:
            context_str += f"User: {exchange['question']}\n"
            context_str += f"SQL: {exchange['sql']}\n"
            context_str += f"Answer: {exchange['answer'][:300]}\n"
            context_str += f"Result data: {exchange.get('result', '')[:400]}\n\n"
        context_str += """CRITICAL FOLLOW UP RULES:
- Read the conversation history AND result data above very carefully
- The result data shows the EXACT values returned by the database
- When user says 'that region' or 'it' or 'its' — look at the previous answer to find the exact region name and add WHERE region = 'exact name'
- When user says 'that customer' — add WHERE customer_name = 'exact name from previous answer'
- When user says 'that category' — add WHERE category = 'exact name from previous answer'
- When user says 'lowest', 'least', 'first', 'second' — find that specific value from the previous result data and use it in WHERE clause
- NEVER query globally when there is conversation history — always filter to the specific entity discussed
- Never return UNRELATED when there is conversation history\n"""

    date_instruction = ""
    if date_ctx:
        date_instruction = f"""
CRITICAL DATE INSTRUCTION:
The user is asking about: {date_ctx['desc']}
Use EXACTLY this date filter in your WHERE clause (already calculated for you):
  {date_ctx['filter']}
Do NOT calculate dates yourself — just use the filter above as-is.
"""

    prompt = f"""You are an expert Sales Analytics Assistant for a retail company. Convert plain English questions to valid PostgreSQL SQL queries.

Today's date is {datetime.now().strftime('%Y-%m-%d')}.

{metadata}

The orders table has data from 2022 to 2026.
The returns table is linked to orders via order_id.
The customers table has demographics: age, sex, prime_status, payment_mode, channel, device.
The customers table links to orders via customer_name.

MOST IMPORTANT RULE:
- If the question is unrelated to the database return exactly: UNRELATED
- NEVER return SQL comments (lines starting with --)
- NEVER return explanations instead of SQL

{date_instruction}

CUSTOMERS TABLE RULES:
- customers table has: customer_name, age, sex, prime_status, payment_mode, channel, device
- To join: JOIN customers c ON o.customer_name = c.customer_name
- For age group questions: WHERE c.age BETWEEN 20 AND 30
- For prime status: WHERE c.prime_status = 'Prime' or 'Non-Prime'
- For channel: WHERE c.channel = 'Web' or 'Mobile'
- For device: WHERE c.device = 'Android' or 'iOS' or 'Desktop'
- Always prefix customers columns with c. alias
- For "details of [name]" or "all details of [name]": 
  SELECT c.*, o.* FROM customers c 
  LEFT JOIN orders o ON c.customer_name = o.customer_name 
  WHERE LOWER(c.customer_name) = LOWER('name here') LIMIT 20
- NEVER return UNRELATED for questions containing a person's name
- If a name is mentioned, always look them up in customers or orders table
- customers table has NO customer_id column — always use COUNT(DISTINCT customer_name) to count customers
- Never use customer_id in any query
-- For age group questions, ALWAYS use this exact SQL structure:
  SELECT
    CASE 
      WHEN c.age BETWEEN 0 AND 10 THEN '0-10 (Child)'
      WHEN c.age BETWEEN 11 AND 20 THEN '11-20 (Teenager)'
      WHEN c.age BETWEEN 21 AND 30 THEN '21-30 (Young Adult)'
      WHEN c.age BETWEEN 31 AND 40 THEN '31-40 (Adult)'
      WHEN c.age BETWEEN 41 AND 50 THEN '41-50 (Middle Aged)'
      WHEN c.age BETWEEN 51 AND 60 THEN '51-60 (Senior)'
      WHEN c.age > 60 THEN '60+ (Elderly)'
      ELSE 'Unknown'
    END AS age_group,
    SUM(o.sales) AS total_sales
  FROM orders o
  JOIN customers c ON o.customer_name = c.customer_name
  GROUP BY age_group
  ORDER BY total_sales DESC
- The CASE statement goes inside SELECT, never after JOIN or WHERE
- GROUP BY age_group refers to the alias, not the full CASE expression
- When question asks for customer details WITH age and prime status, always include c.age and c.prime_status in SELECT
- Always add c.age, c.prime_status to GROUP BY when they are in SELECT
- Example: SELECT o.customer_name, ROUND(SUM(o.profit)::numeric, 2) AS total_profit, c.age, c.prime_status
  FROM orders o JOIN customers c ON o.customer_name = c.customer_name
  GROUP BY o.customer_name, c.age, c.prime_status
  ORDER BY total_profit DESC LIMIT 5

RELATIONSHIP BETWEEN TABLES:
- orders and returns are linked via order_id
- Always read the DATABASE STRUCTURE above carefully before writing SQL
- Only use columns that actually exist in each table as listed in DATABASE STRUCTURE
- For return rate by any group always use this exact pattern:
  SELECT o.group_column, ROUND((COUNT(DISTINCT r.return_id)::numeric / NULLIF(COUNT(DISTINCT o.order_id)::numeric, 0)) * 100, 2) AS return_rate FROM orders o LEFT JOIN returns r ON r.order_id = o.order_id GROUP BY o.group_column ORDER BY return_rate DESC
- Always prefix ambiguous columns with table alias
- Never write correlated subqueries for return rates — always use LEFT JOIN with COUNT
- The returns table does NOT have region, state, city, segment, category, customer_name, product_name, quantity_returned or product columns — it only has return_id and order_id
- To find who returned a specific product: JOIN orders o ON r.order_id = o.order_id then WHERE o.product_name = 'name'
- To get returns by region: JOIN orders ON returns.order_id = orders.order_id and GROUP BY orders.region
- Any time user asks for returns breakdown BY any group — always JOIN returns with orders

BASIC SQL RULES:
- Return ONLY the raw SQL query, no backticks, no markdown, no explanations, no semicolon
- Always include FROM clause in every query
- Always include every non-aggregate column in GROUP BY
- Always use ROUND(SUM(column)::numeric, 2) for decimal results
- Never nest aggregate functions like SUM(SUM()) or ROUND(ROUND())
- Always alias columns with meaningful names using AS
- Always include the GROUP BY column in SELECT
- highest maximum top most all mean ORDER BY DESC
- lowest minimum bottom least all mean ORDER BY ASC
- When user asks for sales always use SUM(sales) never SUM(profit)
- When user asks for profit always use SUM(profit) never SUM(sales)
- Always use lowercase table names
- Never use COUNT(*) always use COUNT(DISTINCT column_name)
- When using JOIN always prefix ALL column references with table alias
- For division inside ROUND, ALWAYS cast the numerator: ROUND(SUM(o.sales)::numeric / NULLIF(COUNT(DISTINCT o.order_id), 0), 2)
- Never cast only the denominator — always cast the numerator or the whole expression
- Correct: ROUND(SUM(sales)::numeric / NULLIF(COUNT(DISTINCT order_id), 0), 2)
- Wrong:   ROUND(SUM(sales) / COUNT(DISTINCT order_id)::numeric, 2)

AGGREGATION RULES:
- For top N: ORDER BY value DESC LIMIT N
- For bottom N: ORDER BY value ASC LIMIT N
- For profit margin: ROUND((SUM(profit)::numeric / NULLIF(SUM(sales)::numeric, 0)) * 100, 2)
- For average order value: ROUND((SUM(sales)::numeric / NULLIF(COUNT(DISTINCT order_id), 0)), 2)
- For overall return rate use: SELECT ROUND(((SELECT COUNT(DISTINCT return_id) FROM returns)::numeric / NULLIF((SELECT COUNT(DISTINCT order_id) FROM orders)::numeric, 0)) * 100, 2) AS return_percentage
- For loss making orders: WHERE profit < 0
- For high discount orders: WHERE discount > 0.30
- For cumulative/running total by month use window function:
  SELECT TO_CHAR(order_date, 'YYYY-MM') AS year_month,
  ROUND(SUM(SUM(sales)) OVER (ORDER BY TO_CHAR(order_date, 'YYYY-MM'))::numeric, 2) AS cumulative_sales
  FROM orders GROUP BY TO_CHAR(order_date, 'YYYY-MM') ORDER BY year_month
- NEVER use regular SUM for cumulative — always use SUM(SUM()) OVER() window function
- For "highest discount" or "most discount" by city/region/state — always use AVG(discount) not SUM(discount)
- SUM(discount) is meaningless for discount analysis — always use AVG or MAX

DELIVERY RULES:
- For average delivery time: SELECT ROUND(AVG(EXTRACT(EPOCH FROM (ship_date - order_date)) / 86400)::numeric, 2) AS avg_delivery_days FROM orders
- For late deliveries: WHERE EXTRACT(EPOCH FROM (ship_date - order_date)) / 86400 > 5

DATE RULES:
- Today is {datetime.now().strftime('%Y-%m-%d')}. You ALWAYS know the current date.
- Always use DATE_TRUNC for date range comparisons
- For last month: WHERE order_date >= DATE_TRUNC('month', CURRENT_DATE - INTERVAL '1 months') AND order_date < DATE_TRUNC('month', CURRENT_DATE)
- For comparing two months always use CASE WHEN with DATE_TRUNC ranges
- For 2023 vs 2024 monthly comparison, use: SELECT TO_CHAR(order_date, 'Mon') AS month, SUM(CASE WHEN EXTRACT(YEAR FROM order_date) = 2023 THEN sales ELSE 0 END) AS sales_2023, SUM(CASE WHEN EXTRACT(YEAR FROM order_date) = 2024 THEN sales ELSE 0 END) AS sales_2024 FROM orders WHERE EXTRACT(YEAR FROM order_date) IN (2023, 2024) GROUP BY TO_CHAR(order_date, 'Mon'), EXTRACT(MONTH FROM order_date) ORDER BY EXTRACT(MONTH FROM order_date)
- NEVER use BETWEEN for date comparisons always use >= and <
- When using CASE alias in ORDER BY, always repeat the full CASE expression in ORDER BY, never reference the alias name
- Always use INTERVAL '1 months' not INTERVAL '1 month'

LIMIT RULES:
- NEVER use LIMIT for breakdown questions: by category, by region, by segment, by state, by city, by year, by month, by quarter
- Only use LIMIT when user says: top N, bottom N, highest, lowest, most, least

RETURNS TABLE RULES:
- For refund analysis always use SUM(refund_amount) from returns table
- For return reasons use GROUP BY reason from returns table
- Always use LEFT JOIN returns with orders when calculating return rates by group

CTE RULES:
- When using CTEs with WITH clause, always reference the SUMMARY CTE not the base CTE
- If you have SeanReturns and SeanReturnSummary, use SeanReturnSummary in the final SELECT
- Never reference a base CTE in the final SELECT if a summary CTE exists for it
- Always include all CTEs in the FROM clause of the final SELECT

RETURN RATE RULES:
- For return rate by prime_status or any customer demographic, always use this exact pattern:
  SELECT 
    c.prime_status,
    ROUND((COUNT(DISTINCT r.return_id)::numeric / NULLIF(COUNT(DISTINCT o.order_id)::numeric, 0)) * 100, 2) AS return_rate
  FROM orders o
  JOIN customers c ON o.customer_name = c.customer_name
  LEFT JOIN returns r ON o.order_id = r.order_id
  GROUP BY c.prime_status
  ORDER BY return_rate DESC
- Always use COUNT(DISTINCT) for both return_id and order_id — never COUNT(*)
- Always LEFT JOIN returns so non-returned orders are included in the denominator

{context_str}

Current question: {question}"""

    # Use SQL_MODEL (gemma3) for SQL generation
    sql = call_model(prompt, model=SQL_MODEL, temperature=0.1,
                    num_ctx=4096, timeout=120)
    sql = sql.replace("```sql", "").replace("```", "").strip()
    if sql.endswith(";"): sql = sql[:-1]
    
    # Fix CASE alias in GROUP BY — replace with positional reference
    alias_match = re.search(r'END\s+AS\s+(\w+)', sql, re.IGNORECASE)
    if alias_match:
        alias = alias_match.group(1)
        sql = re.sub(rf'\bGROUP BY\s+{alias}\b', 'GROUP BY 1', sql, flags=re.IGNORECASE)
        sql = re.sub(rf'\bORDER BY\s*\n\s*CASE\s+WHEN\s+{alias}.*', '', sql, flags=re.IGNORECASE | re.DOTALL)
    
    return sql


# ══════════════════════════════════════════════════════════════
# SUMMARISE — uses Mistral (FAST_MODEL)
# ══════════════════════════════════════════════════════════════

def summarise(question, result):
    prompt = f"""You are an expert Sales Analytics Assistant for a retail company.
The user asked: '{question}'
The database returned: {result.to_string()}

Rules:
- Respond in strictly 2 sentences like a business analyst presenting findings
- Be insightful — mention what the number means for the business
- Use exact names and numbers from the data
- Never add labels like 'row 0' or 'index 0' before names
- Never make up numbers not in the data
- Do not start with 'I' or 'According to'
- The leftmost column is the INDEX not the answer
- If result shows a single number use it directly
- If result has multiple rows present all clearly one by one
- If result has two columns that look like a comparison say: first period had X, second period had Y
- If discount value is between 0 and 1 multiply by 100 and show as percentage
- Never say 'round' as a column name always use the alias name.
- If the question asks "what products" or "which products" — list them simply, no insights needed
- If the question is a simple listing question — just list the items clearly, no recommendations needed
- If result contains age column, always mention the age for each person
- If result contains prime_status column, always mention prime or non-prime for each person
- Present ALL columns in the result, not just the first two
-If result contains region, state, city columns — just state which region/state/city the person is from, nothing else
- Never give overall business analysis when the question is about a specific person's location
- If result contains product_name column — always list each product name clearly one by one
- If the question asks which products were returned or bought — list the product names, never just say a count
"""

    # Use FAST_MODEL (mistral) for summarisation
    return call_model(prompt, model=FAST_MODEL, temperature=0.4,
                     num_ctx=4096, timeout=120)


# ══════════════════════════════════════════════════════════════
# RUN QUERY — SQL fixes and execution
# ══════════════════════════════════════════════════════════════

def run_query(sql, question=""):
    try:
        sql = re.sub(r'ROUND\(([^,]+?)(?<!::numeric),\s*2\)', lambda m: f'ROUND(({m.group(1)})::numeric, 2)', sql, flags=re.IGNORECASE)
        sql = re.sub(r'ROUND\((.+?\* 100\))::numeric,\s*2\)', lambda m: f'ROUND(({m.group(1)})::numeric, 2)', sql, flags=re.IGNORECASE)
        sql = re.sub(r'\bOrders\b', 'orders', sql)
        sql = re.sub(r'\bReturns\b', 'returns', sql)
        sql = re.sub(r'\bCustomers\b', 'customers', sql)
        sql = re.sub(r"INTERVAL '(\d+) (month|year|day|week)'", lambda m: f"INTERVAL '{m.group(1)} {m.group(2)}s'", sql, flags=re.IGNORECASE)

        if re.search(r'EXTRACT\s*\(\s*MONTH\s+FROM', sql, re.IGNORECASE):
            if not re.search(r'EXTRACT\s*\(\s*YEAR\s+FROM', sql, re.IGNORECASE):
                if not re.search(r'DATE_TRUNC', sql, re.IGNORECASE):
                    current_year = datetime.now().year
                    sql = re.sub(r'\bWHERE\b', f'WHERE EXTRACT(YEAR FROM order_date) = {current_year} AND', sql, count=1, flags=re.IGNORECASE)

        def fix_extract_month(m):
            col = m.group(1)
            month_num = int(m.group(2))
            current_month = datetime.now().month
            months_ago = (current_month - month_num) % 12
            if months_ago == 0: months_ago = 12
            return (f"{col} >= DATE_TRUNC('month', CURRENT_DATE - INTERVAL '{months_ago} months') "
                   f"AND {col} < DATE_TRUNC('month', CURRENT_DATE - INTERVAL '{months_ago - 1} months')")

        sql = re.sub(r"EXTRACT\s*\(\s*MONTH\s+FROM\s+(\w+)\s*\)\s*=\s*(\d+)", fix_extract_month, sql, flags=re.IGNORECASE)
        sql = re.sub(r"DATE_TRUNC\(('(?:quarter|month|year|week|day)'),\s*'(\d{4}-\d{2}-\d{2})'\)", lambda m: f"DATE_TRUNC({m.group(1)}, '{m.group(2)}'::date)", sql, flags=re.IGNORECASE)

        if any(word in question.lower() for word in ["cumulative", "running total", "accumulated"]):
            sql = re.sub(r'\bSUM\((\w+)\)\s+OVER\s*\(', r'SUM(SUM(\1)) OVER (', sql, flags=re.IGNORECASE)
            print(f"  🔧 Fixed: Corrected cumulative window function")

        if re.search(r'GROUP BY\s+\d+\s*,\s*\d+', sql, re.IGNORECASE):
            if re.search(r'SUM\s*\(', sql, re.IGNORECASE):
                sql = re.sub(r'GROUP BY[\s\d,]+', '', sql, flags=re.IGNORECASE)

        if re.search(r'\bJOIN\b', sql, re.IGNORECASE) and table_columns:
            group_by_match = re.search(r'GROUP BY\s+(\w+)', sql, re.IGNORECASE)
            if group_by_match:
                col = group_by_match.group(1).lower()
                tables_with_col = [t for t, cols in table_columns.items() if col in cols]
                if len(tables_with_col) > 1:
                    main_table_match = re.search(r'FROM\s+(\w+)\s+(\w+)', sql, re.IGNORECASE)
                    if main_table_match:
                        main_alias = main_table_match.group(2)
                        sql = re.sub(rf'\bSELECT\s+{col}\b', f'SELECT {main_alias}.{col}', sql, flags=re.IGNORECASE)
                        sql = re.sub(rf'\bGROUP BY\s+{col}\b', f'GROUP BY {main_alias}.{col}', sql, flags=re.IGNORECASE)

        if table_columns:
            sql_keywords = {'where','from','join','on','and','or','not','in','is','as','by','order','group','having','select','distinct','count','sum','avg','max','min','round','extract','case','when','then','else','end','null','true','false','limit','offset','union','all','inner','outer','left','right','full','cross','natural','with'}
            for alias_match in re.finditer(r'\b([a-zA-Z_]\w*)\.(\w+)', sql):
                alias = alias_match.group(1)
                col = alias_match.group(2).lower()
                if alias.lower() in sql_keywords: continue
                from_match = re.search(rf'(?:FROM|JOIN)\s+(\w+)\s+{re.escape(alias)}\b', sql, re.IGNORECASE)
                if from_match:
                    table = from_match.group(1).lower()
                    if table in table_columns and col not in table_columns[table]:
                        for other_table, cols in table_columns.items():
                            if col in cols and other_table != table:
                                other_alias_match = re.search(rf'(?:FROM|JOIN)\s+{other_table}\s+([a-zA-Z_]\w*)\b', sql, re.IGNORECASE)
                                if other_alias_match:
                                    correct_alias = other_alias_match.group(1)
                                    sql = sql.replace(f'{alias}.{col}', f'{correct_alias}.{col}')
                                    break

        for normalized, quoted in special_columns.items():
            pattern = r'(?<!")(?<!\w)' + re.escape(normalized) + r'(?!\w)(?!")'
            sql = re.sub(pattern, quoted, sql, flags=re.IGNORECASE)

        sql = re.sub(r'(?<!\w)ROUND\(SUM\((\w+)\)::numeric,\s*2\)(?!\s+AS\b)', lambda m: f'ROUND(SUM({m.group(1)})::numeric, 2) AS total_{m.group(1)}', sql, flags=re.IGNORECASE)
        
        breakdown_words = [" by ", "by category", "by region", "by segment", "by state", "by city", "by year", "by month", "by quarter", "by ship mode", "trends", "monthly", "quarterly"]
        limit_words = ["highest", "lowest", "most", "least", "top", "bottom", "best", "worst", "maximum", "minimum"]
        is_breakdown = any(word in question.lower() for word in breakdown_words)
        is_specific = any(word in question.lower() for word in limit_words)
        if is_breakdown and not is_specific:
            sql = re.sub(r'LIMIT\s+1', '', sql, flags=re.IGNORECASE)

        group_match = re.search(r'GROUP BY (\w+\.\w+|\w+)', sql, re.IGNORECASE)
        select_match = re.search(r'SELECT (.+?) FROM', sql, re.IGNORECASE)
        if group_match and select_match:
            group_col = group_match.group(1)
            select_cols = select_match.group(1)
            # Only add if it's a simple column name (not alias.column) and not already in SELECT
            if '.' not in group_col and group_col.lower() not in select_cols.lower():
                sql = sql.replace('SELECT ', f'SELECT {group_col}, ', 1)

       
       # Fix %% back to % (Python f-string escaping)
       # Fix %% back to % (Python f-string escaping)
# Fix %% back to % (Python f-string escaping)
        sql = sql.replace('%%', '%')

        # Fix CASE alias in GROUP BY/ORDER BY
        case_alias_match = re.search(r'END\s+AS\s+(\w+)', sql, re.IGNORECASE)
        if case_alias_match:
            alias = case_alias_match.group(1)
            sql = re.sub(rf'\bGROUP BY\s+{alias}\b', 'GROUP BY 1', sql, flags=re.IGNORECASE)
            sql = re.sub(rf'\bORDER BY\s*\n\s*CASE\s+WHEN\s+{alias}.*', '', sql, flags=re.IGNORECASE | re.DOTALL)

        sql = re.sub(r'\br\.product_name\b', 'o.product_name', sql, flags=re.IGNORECASE)
        sql = re.sub(r'\br\.customer_name\b', 'o.customer_name', sql, flags=re.IGNORECASE)
        sql = re.sub(r'\br\.quantity_returned\b', 'r.return_id', sql, flags=re.IGNORECASE)
        sql = re.sub(r'\br\.category\b', 'o.category', sql, flags=re.IGNORECASE)
        sql = re.sub(r'\br\.region\b', 'o.region', sql, flags=re.IGNORECASE)
        sql = re.sub(r'SUM\(COUNT\(DISTINCT r\.return_id\)\)', 'COUNT(DISTINCT r.return_id)', sql, flags=re.IGNORECASE)
        sql = re.sub(r'ORDER BY SUM\(r\.return_id\)', 'ORDER BY COUNT(DISTINCT r.return_id)', sql, flags=re.IGNORECASE)

        with engine.connect() as conn:
            result = pd.read_sql(text(sql), conn)

            for col in result.columns:
                if 'discount' in col.lower():
                    try:
                        numeric_col = pd.to_numeric(result[col], errors='coerce')
                        if numeric_col.notna().all() and numeric_col.between(0, 1).all():
                            result[col] = (numeric_col * 100).round(2)
                            result = result.rename(columns={col: col + '_%'})
                    except Exception:
                        pass
            return result

    except Exception as e:
        print(f"  ❌ SQL Error: {e}")
        return None


# ══════════════════════════════════════════════════════════════
# VISUALISATION
# ══════════════════════════════════════════════════════════════

def needs_visualisation(question):
    no_chart_phrases = [
        "give me details", "show me details", "tell me about",
        "what is this database", "how can you help", "what are you",
        "what do you do", "how many unique", "what is the total",
        "who is the", "what is the name", "what is this",
        "which customer", "which product", "which state",
        "which city", "which region", "which category"
    ]
    question_lower = question.lower()
    if any(phrase in question_lower for phrase in no_chart_phrases):
        return False

    always_chart_phrases = [
        "by region", "by category", "by segment", "by state", "by city",
        "by month", "by quarter", "by year", "trend", "compare", "comparison",
        "breakdown", "distribution", "percentage", "percent", "chart", "graph",
        "plot", "visualise", "visualize", "monthly", "quarterly", "yearly",
        "growth", "pattern", "refund", "sales by", "profit by", "orders by",
        "return rate by", "return reasons", "top ", "bottom ", "show chart",
        "with chart", "vs", "versus", "cumulative", "running total",
        "by age", "by sex", "by channel", "by device", "by payment"
    ]
    if any(phrase in question_lower for phrase in always_chart_phrases):
        return True

    return False


def create_visualisation(result, question):
    global latest_chart_buf
    try:
        if len(result.columns) < 2:
            return None

        if len(result) == 1 and len(result.columns) == 2:
            col1, col2 = result.columns[0], result.columns[1]
            val1 = pd.to_numeric(result[col1].iloc[0], errors='coerce')
            val2 = pd.to_numeric(result[col2].iloc[0], errors='coerce')
            if pd.notna(val1) and pd.notna(val2):
                result = pd.DataFrame({
                    'period': [col1.replace('_', ' ').title(), col2.replace('_', ' ').title()],
                    'value': [val1, val2]
                })

        if len(result) < 2:
            return None

        fig, ax = plt.subplots(figsize=(12, 6))
        fig.patch.set_facecolor('#1c1915')
        ax.set_facecolor('#1c1915')

        x_col = result.columns[0]
        y_col = result.columns[1]
        x_values = result[x_col].astype(str)
        y_values = pd.to_numeric(result[y_col], errors='coerce')
        question_lower = question.lower()

        def format_x_label(val):
            try:
                if re.match(r'^\d{4}-\d{2}$', val):
                    dt = datetime.strptime(val, '%Y-%m')
                    return dt.strftime('%b %Y')
                return val
            except Exception:
                return val

        x_values = x_values.apply(format_x_label)

        # PIE CHART
        if any(word in question_lower for word in ["pie", "percentage", "percent",  "contribution"]):
            numeric_col = None
            for col in result.columns:
                if pd.to_numeric(result[col], errors='coerce').notna().all():
                    numeric_col = col
                    break
            if numeric_col is None:
                return None
            if len(result) == 1:
                pct = min(float(result[numeric_col].iloc[0]), 100)
                values = [pct, 100 - pct]
                labels = [f'Yes ({pct:.1f}%)', f'No ({100-pct:.1f}%)']
                colors = ['#E85D04', '#1D9E75']
            else:
                values = pd.to_numeric(result[numeric_col], errors='coerce')
                labels = result[result.columns[0]].astype(str)
                colors = ['#2E75B6', '#1D9E75', '#E85D04', '#534AB7', '#BA7517',
                         '#C62828', '#00838F', '#558B2F', '#6A1B9A', '#2E7D32']
            wedges, texts, autotexts = ax.pie(
                values, labels=labels, colors=colors[:len(values)],
                autopct='%1.1f%%', startangle=90,
                wedgeprops={'edgecolor': 'white', 'linewidth': 2}
            )
            for text in texts: text.set_color('#e2e8f0')
            for autotext in autotexts: autotext.set_color('#ffffff')
            ax.set_title(question.title(), fontsize=14, fontweight='bold', pad=20, color='#e2e8f0')
            plt.tight_layout()
            buf = io.BytesIO()
            plt.savefig(buf, format='png', dpi=150, bbox_inches='tight')
            plt.close()
            buf.seek(0)
            latest_chart_buf = buf
            print(f"\n  📊 Pie chart generated!")
            return True

        # AREA CHART — before line chart
        if any(word in question_lower for word in ["cumulative", "running total", "accumulated"]):
            ax.fill_between(range(len(x_values)), y_values, alpha=0.4, color='#2E75B6')
            ax.plot(range(len(x_values)), y_values, color='#2E75B6', linewidth=2)
            ax.set_xticks(range(len(x_values)))
            ax.set_xticklabels(x_values, rotation=45, ha='right', color='#aaaaaa')
            chart_type = "Area"

        # LINE CHART
        elif any(word in question_lower for word in ["trend", "monthly", "quarterly", "yearly", "over time", "by month", "by year", "by quarter"]):
            ax.plot(x_values, y_values, marker='o', linewidth=2.5,
                   color='#2E75B6', markersize=8, markerfacecolor='white', markeredgewidth=2)
            ax.fill_between(range(len(x_values)), y_values, alpha=0.1, color='#2E75B6')
            chart_type = "Line"

        # SCATTER CHART
        elif any(word in question_lower for word in ["correlation", "relationship", "scatter"]):
            ax.scatter(x_values, y_values, color='#2E75B6', alpha=0.6, s=100, edgecolors='white')
            chart_type = "Scatter"

        # GROUPED BAR
        elif any(word in question_lower for word in ["vs", "versus", "compare", "comparison"]) and len(result.columns) >= 3:
            x = range(len(result))
            width = 0.35
            col_colors = ['#2E75B6', '#E85D04', '#1D9E75', '#534AB7']
            for i, col in enumerate(result.columns[1:3]):
                vals = pd.to_numeric(result[col], errors='coerce')
                offset = (i - 0.5) * width
                bars = ax.bar([xi + offset for xi in x], vals, width,
                             label=col.replace('_', ' ').title(),
                             color=col_colors[i], edgecolor='white')
                for bar, val in zip(bars, vals):
                    ax.text(bar.get_x() + bar.get_width()/2., bar.get_height(),
                           f'{val:,.0f}', ha='center', va='bottom', fontsize=8, fontweight='bold', color='#ffffff')
            ax.set_xticks(list(x))
            ax.set_xticklabels(result.iloc[:, 0].astype(str), rotation=45, ha='right', color='#aaaaaa')
            ax.legend(labelcolor='#aaaaaa', facecolor='#1c1915', edgecolor='#333333')
            chart_type = "Grouped Bar"

        # HORIZONTAL BAR
        elif any(word in question_lower for word in ["top", "bottom"]) and len(result) > 3:
            sorted_df = result.copy()
            sorted_df[y_col] = pd.to_numeric(sorted_df[y_col], errors='coerce')
            sorted_df = sorted_df.sort_values(by=y_col, ascending=True)
            x_values = sorted_df[x_col].astype(str).apply(format_x_label)
            y_values = sorted_df[y_col]
            colors = plt.cm.Blues([0.3 + 0.7 * i / len(sorted_df) for i in range(len(sorted_df))])
            bars = ax.barh(x_values, y_values, color=colors, edgecolor='white')
            for bar, val in zip(bars, y_values):
                ax.text(bar.get_width(), bar.get_y() + bar.get_height()/2.,
                       f' {val:,.2f}', ha='left', va='center', fontsize=9, fontweight='bold', color='#ffffff')
            ax.tick_params(axis='y', colors='#aaaaaa')
            chart_type = "Horizontal Bar"

        # VERTICAL BAR (default)
        else:
            colors = ['#2E75B6', '#1D9E75', '#E85D04', '#534AB7', '#BA7517',
                     '#C62828', '#00838F', '#558B2F', '#6A1B9A', '#2E7D32']
            bars = ax.bar(x_values, y_values, color=colors[:len(result)],
                         edgecolor='white', linewidth=0.5, width=0.6)
            for bar, val in zip(bars, y_values):
                ax.text(bar.get_x() + bar.get_width()/2., bar.get_height(),
                       f'{val:,.2f}', ha='center', va='bottom', fontsize=9, fontweight='bold', color='#ffffff')
            chart_type = "Bar"

        ax.set_title(question.title(), fontsize=14, fontweight='bold', pad=20, color='#e2e8f0')

        if chart_type == "Horizontal Bar":
            ax.set_xlabel(y_col.replace('_', ' ').title(), fontsize=11, color='#aaaaaa')
            ax.set_ylabel(x_col.replace('_', ' ').title(), fontsize=11, color='#aaaaaa')
            ax.grid(axis='x', alpha=0.15, linestyle='--', color='white')
        else:
            ax.set_xlabel(x_col.replace('_', ' ').title(), fontsize=11, color='#aaaaaa')
            ax.set_ylabel(y_col.replace('_', ' ').title(), fontsize=11, color='#aaaaaa')
            ax.grid(axis='y', alpha=0.15, linestyle='--', color='white')

        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
        ax.spines['left'].set_color('#333333')
        ax.spines['bottom'].set_color('#333333')
        ax.tick_params(colors='#aaaaaa')
        plt.xticks(rotation=45, ha='right', color='#aaaaaa')
        plt.yticks(color='#aaaaaa')
        plt.tight_layout()

        buf = io.BytesIO()
        plt.savefig(buf, format='png', dpi=150, bbox_inches='tight')
        plt.close()
        buf.seek(0)
        latest_chart_buf = buf
        print(f"\n  📊 {chart_type} chart generated!")
        return True

    except Exception as e:
        print(f"  ⚠️ Could not generate chart: {e}")
        return None



# ══════════════════════════════════════════════════════════════
# PREDICTIVE ANALYTICS — handle_predictive()
# Uses 3-month moving average on historical monthly data
# ══════════════════════════════════════════════════════════════

def _get_monthly_data(metric='sales', group_by=None, specific_filter=None):
    """Fetch monthly historical data, excluding current incomplete month."""
    if metric == 'sales':
        agg = "ROUND(SUM(o.sales)::numeric, 2)"
        col = "total_sales"
    elif metric == 'profit':
        agg = "ROUND(SUM(o.profit)::numeric, 2)"
        col = "total_profit"
    elif metric == 'orders':
        agg = "COUNT(DISTINCT o.order_id)"
        col = "total_orders"
    elif metric == 'returns':
        agg = "COUNT(DISTINCT r.return_id)"
        col = "total_returns"
    elif metric == 'refund_amount':
        agg = "ROUND(SUM(r.refund_amount)::numeric, 2)"
        col = "total_refund_amount"
    elif metric == 'customers':
        agg = "COUNT(DISTINCT o.customer_name)"
        col = "total_customers"
    else:
        agg = "ROUND(SUM(o.sales)::numeric, 2)"
        col = "total_sales"

    join_clause = ""
    filter_clause = ""
    if specific_filter:
        fcol, fval = specific_filter
        fval_escaped = fval.replace("'", "''")
        if fcol == 'prime_status':
            join_clause = "JOIN customers c ON o.customer_name = c.customer_name"
            filter_clause = f"AND c.prime_status = '{fval_escaped}'"
        else:
            filter_clause = f"AND o.{fcol} = '{fval_escaped}'"

    # The actual DB column is "sub-category" (with a hyphen), needs quoting
    group_col_sql = 'o."sub-category"' if group_by == 'sub_category' else (f"o.{group_by}" if group_by else None)
    group_col_select = 'sub_category' if group_by == 'sub_category' else group_by
    if group_by:
        sql = f"""
            SELECT TO_CHAR(o.order_date, 'YYYY-MM') AS month,
                   {group_col_sql} AS {group_col_select},
                   {agg} AS {col}
            FROM orders o
            LEFT JOIN returns r ON o.order_id = r.order_id
            {join_clause}
            WHERE o.order_date < DATE_TRUNC('month', CURRENT_DATE) {filter_clause}
            GROUP BY TO_CHAR(o.order_date, 'YYYY-MM'), {group_col_sql}
            ORDER BY month
        """
    else:
        sql = f"""
            SELECT TO_CHAR(o.order_date, 'YYYY-MM') AS month,
                   {agg} AS {col}
            FROM orders o
            LEFT JOIN returns r ON o.order_id = r.order_id
            {join_clause}
            WHERE o.order_date < DATE_TRUNC('month', CURRENT_DATE) {filter_clause}
            GROUP BY TO_CHAR(o.order_date, 'YYYY-MM')
            ORDER BY month
        """
    try:
        with engine.connect() as conn:
            df = pd.read_sql(text(sql), conn)
        return df, col
    except Exception as e:
        print(f"  ❌ Predictive data fetch error: {e}")
        return None, col


def _moving_average_forecast(values, n_ahead=1, window=3):
    """
    3-month moving average forecast with capped trend percentage
    to avoid unrealistic swings on volatile/small datasets.
    Returns: predictions (list), ma (base average), trend_pct, std
    """
    if len(values) < window:
        window = max(1, len(values))

    recent = values[-window:]
    ma = float(np.mean(recent))

    if len(values) >= window * 2:
        prev = values[-(window * 2):-window]
        prev_ma = float(np.mean(prev))
        trend_pct = ((ma - prev_ma) / prev_ma * 100) if prev_ma != 0 else 0
        trend_pct = max(min(trend_pct, 30), -30)  # cap at +/-30% to stay realistic
    else:
        trend_pct = 0

    predictions = []
    monthly_growth = trend_pct / 100 / window
    for i in range(n_ahead):
        projected = ma * (1 + monthly_growth * (i + 1))
        predictions.append(projected)

    std = float(np.std(recent))
    return predictions, ma, trend_pct, std


def _get_next_month_label(n=1):
    """Get label(s) for next N months."""
    now = datetime.now()
    labels = []
    for i in range(1, n + 1):
        month = now.month + i
        year = now.year
        while month > 12:
            month -= 12
            year += 1
        labels.append(datetime(year, month, 1).strftime('%B %Y'))
    return labels


def _detect_prediction_intent(question):
    """
    Detect what metric, timeframe, grouping and scenario the user wants.
    Returns: metric, metric_name, n_months, timeframe, group_by, scenario, specific_filter
    """
    q = question.lower()
    now = datetime.now()

    # Metric detection
    if any(w in q for w in ['refund amount', 'refund']):
        metric, metric_name = 'refund_amount', 'refund amount'
    elif any(w in q for w in ['profit', 'profitable', 'profit margin']):
        metric, metric_name = 'profit', 'profit'
    elif any(w in q for w in ['return rate', 'returns']):
        metric, metric_name = 'returns', 'returns'
    elif any(w in q for w in ['order', 'orders']):
        metric, metric_name = 'orders', 'orders'
    elif any(w in q for w in ['customer', 'customers']) and any(w in q for w in ['new', 'how many']):
        metric, metric_name = 'customers', 'new customers'
    else:
        metric, metric_name = 'sales', 'sales'

    # Timeframe detection (priority order matters)
    n_months = 1
    timeframe = 'next month'
    found_specific = False

    month_names = {
        'january': 1, 'february': 2, 'march': 3, 'april': 4,
        'may': 5, 'june': 6, 'july': 7, 'august': 8,
        'september': 9, 'october': 10, 'november': 11, 'december': 12
    }

    # 1. Specific "Month YYYY"
    for month_name, month_num in month_names.items():
        year_match = re.search(rf'{month_name}\s+(\d{{4}})', q)
        if year_match:
            target_year = int(year_match.group(1))
            months_ahead = (target_year - now.year) * 12 + (month_num - now.month)
            if months_ahead > 0:
                n_months = months_ahead
                timeframe = f'{month_name.title()} {target_year}'
                found_specific = True
                break

    # 2. Specific quarter — e.g. "Q3 2026"
    if not found_specific:
        qm = re.search(r'q([1-4])\s*(\d{4})', q)
        if qm:
            qnum = int(qm.group(1))
            qyear = int(qm.group(2))
            quarter_start_month = {1: 1, 2: 4, 3: 7, 4: 10}[qnum]
            months_ahead = (qyear - now.year) * 12 + (quarter_start_month - now.month)
            n_months = 3
            timeframe = f'Q{qnum} {qyear}'
            found_specific = True

    # 2b. "growth/expected/outlook for YYYY" — rest of that year, growth-focused
    is_growth_question = False
    if not found_specific:
        year_growth_match = re.search(r'(?:growth|expected|outlook|forecast)\D{0,15}(20\d{2})\b', q)
        if not year_growth_match:
            year_growth_match = re.search(r'\bfor\s+(20\d{2})\b', q)
        if year_growth_match:
            target_year = int(year_growth_match.group(1))
            if target_year == now.year:
                months_ahead = 12 - now.month
                timeframe = f'rest of {target_year}'
            else:
                months_ahead = (target_year - now.year) * 12 + (12 - now.month)
                timeframe = f'{target_year}'
            if months_ahead > 0:
                n_months = months_ahead
                found_specific = True
                is_growth_question = 'growth' in q

    # 3. Relative timeframes
    if not found_specific:
        if any(w in q for w in ['next quarter', 'next 3 months']):
            n_months, timeframe = 3, 'next quarter (3 months)'
        elif any(w in q for w in ['next 6 months', 'next half year']):
            n_months, timeframe = 6, 'next 6 months'
        elif any(w in q for w in ['next year', 'next 12 months', 'this time next year']):
            n_months, timeframe = 12, 'next year'
        else:
            n_months, timeframe = 1, 'next month'

    # Group-by / specific-entity detection
    group_by = None
    specific_filter = None

    regions = ['west', 'east', 'south', 'central']
    categories = ['furniture', 'technology', 'office supplies']
    segments = ['consumer', 'corporate', 'home office']

    mentioned_region = next((r for r in regions if r in q), None)
    mentioned_category = next((c for c in categories if c in q), None)
    mentioned_segment = next((s for s in segments if s in q), None)
    mentioned_prime = None
    if 'non-prime' in q or 'non prime' in q:
        mentioned_prime = 'Non-Prime'
    elif 'prime' in q:
        mentioned_prime = 'Prime'

    if mentioned_region:
        specific_filter = ('region', mentioned_region.title())
    elif mentioned_category:
        specific_filter = ('category', mentioned_category.title())
    elif mentioned_segment:
        specific_filter = ('segment', mentioned_segment.title())
    elif mentioned_prime:
        specific_filter = ('prime_status', mentioned_prime)
    elif 'region' in q:
        group_by = 'region'
    elif any(w in q for w in ['sub-category', 'sub category', 'subcategory']):
        group_by = 'sub_category'
    elif 'category' in q:
        group_by = 'category'
    elif 'segment' in q:
        group_by = 'segment'
    elif any(w in q for w in ['bestseller', 'best selling', 'best-selling', 'which product']):
        group_by = 'product_name'

    # Scenario detection
    scenario = None
    discount_match = re.search(r'reduce discount[s]? by (\d+)', q)
    sales_growth_match = re.search(r'sales grow(?:s)? by (\d+)', q)
    return_drop_match = re.search(r'return rate drops? by (\d+)', q)
    if discount_match:
        scenario = {'type': 'discount_reduction', 'value': int(discount_match.group(1))}
    elif sales_growth_match:
        scenario = {'type': 'sales_growth', 'value': int(sales_growth_match.group(1))}
    elif return_drop_match:
        scenario = {'type': 'return_drop', 'value': int(return_drop_match.group(1))}

    return metric, metric_name, n_months, timeframe, group_by, scenario, specific_filter


def _handle_yoy_comparison(question):
    """Handle 'compared to last year' style questions directly with SQL."""
    try:
        with engine.connect() as conn:
            result = conn.execute(text("""
                SELECT EXTRACT(YEAR FROM order_date) AS year,
                       ROUND(SUM(sales)::numeric, 2) AS total_sales
                FROM orders
                WHERE EXTRACT(YEAR FROM order_date) IN (2023, 2024, 2025, 2026)
                GROUP BY year ORDER BY year
            """)).fetchall()
        rows = {int(r[0]): float(r[1]) for r in result}
        years = sorted(rows.keys())
        lines = []
        for i, yr in enumerate(years):
            if i == 0:
                lines.append(f"{yr}: ${rows[yr]:,.2f}")
            else:
                prev_yr = years[i - 1]
                growth = ((rows[yr] - rows[prev_yr]) / rows[prev_yr] * 100) if rows[prev_yr] else 0
                lines.append(f"{yr}: ${rows[yr]:,.2f} ({growth:+.1f}% vs {prev_yr})")
        answer = "Year over Year Sales:\n\n" + "\n".join(lines)
        context.append({"question": question, "sql": "", "answer": answer, "result": ""})
        if len(context) > 10:
            context.pop(0)
        return answer
    except Exception as e:
        print(f"  ❌ YoY error: {e}")
        return "Could not compute year over year comparison."


def _handle_scenario(question, scenario):
    """Handle what-if scenario questions using simple business math."""
    df_profit, _ = _get_monthly_data(metric='profit')
    df_sales, _ = _get_monthly_data(metric='sales')

    if df_profit is None or df_profit.empty:
        return "Could not fetch data for scenario analysis."

    avg_profit = float(df_profit['total_profit'].mean())
    avg_sales = float(df_sales['total_sales'].mean()) if df_sales is not None and not df_sales.empty else 0

    if scenario['type'] == 'discount_reduction':
        pct = scenario['value']
        lift = avg_profit * (pct / 100) * 0.6
        new_profit = avg_profit + lift
        answer = f"""Scenario Analysis: {pct}% Discount Reduction

Current average monthly profit: ${avg_profit:,.2f}
Estimated profit next month: ${new_profit:,.2f}
Estimated lift: +${lift:,.2f} ({pct * 0.6:.1f}% improvement)

Note: Assumes approximately 60% of the discount reduction converts directly to profit."""

    elif scenario['type'] == 'sales_growth':
        pct = scenario['value']
        new_sales = avg_sales * (1 + pct / 100)
        margin = avg_profit / avg_sales if avg_sales else 0
        new_profit = new_sales * margin
        answer = f"""Scenario Analysis: {pct}% Sales Growth

Current average monthly sales: ${avg_sales:,.2f}
Projected sales: ${new_sales:,.2f}
Projected profit (at {margin*100:.1f}% margin): ${new_profit:,.2f}"""

    elif scenario['type'] == 'return_drop':
        df_returns, _ = _get_monthly_data(metric='returns')
        avg_returns = float(df_returns['total_returns'].mean()) if df_returns is not None and not df_returns.empty else 0
        pct = scenario['value']
        avg_order_value = avg_sales / max(avg_returns, 1) if avg_returns else 0
        returns_avoided = avg_returns * (pct / 100)
        estimated_sales_recovery = returns_avoided * avg_order_value * 0.5
        new_sales = avg_sales + estimated_sales_recovery
        answer = f"""Scenario Analysis: {pct}% Drop in Return Rate

Current average monthly returns: {avg_returns:,.0f}
Estimated returns avoided: {returns_avoided:,.0f}
Current average monthly sales: ${avg_sales:,.2f}
Estimated recovered sales: ${new_sales:,.2f}

Note: This is a directional estimate based on average order value."""

    else:
        answer = "Could not process this scenario."

    context.append({"question": question, "sql": "", "answer": answer, "result": ""})
    if len(context) > 10:
        context.pop(0)
    return answer


def handle_predictive(question):
    """
    Handles predictive/forecasting questions using 3-month moving average.
    Routes to YoY comparison, scenario analysis, group-by forecast,
    or overall forecast depending on the question.
    """
    print(f"  🔮 Predictive intent detected")

    if not SKLEARN_AVAILABLE:
        return "Predictive analytics requires numpy. Run: pip install scikit-learn numpy"

    q_lower = question.lower()

    # Route 1: Year over year comparison
    if any(w in q_lower for w in ['compared to last year', 'vs last year', 'year over year', 'yoy', 'grow compared to last year']):
        return _handle_yoy_comparison(question)
    
    # Route: Goal-seeking — "when will X hit $Y"
    goal_match = re.search(r'hit\s*\$?([\d,]+)\s*(k|m|million|thousand)?\s*in\s*sales', q_lower)
    if not goal_match:
        goal_match = re.search(r'reach\s*\$?([\d,]+)\s*(k|m|million|thousand)?\s*in\s*sales', q_lower)
    if goal_match:
        target_value = float(goal_match.group(1).replace(',', ''))
        unit = goal_match.group(2) or ''
        if unit in ['k', 'thousand']:
            target_value *= 1000
        elif unit in ['m', 'million']:
            target_value *= 1000000

        regions = ['west', 'east', 'south', 'central']
        mentioned_region = next((r for r in regions if r in q_lower), None)
        specific_filter_goal = ('region', mentioned_region.title()) if mentioned_region else None

        df_goal, value_col_goal = _get_monthly_data(metric='sales', specific_filter=specific_filter_goal)
        if df_goal is None or df_goal.empty or len(df_goal) < 6:
            return "Not enough historical data to project when this target will be reached."

        values_goal = df_goal[value_col_goal].astype(float).values
        current_val = float(values_goal[-1])
        _, ma_goal, trend_pct_goal, _ = _moving_average_forecast(values_goal, 1)

        if trend_pct_goal <= 0:
            label = f" for {specific_filter_goal[1]} {specific_filter_goal[0]}" if specific_filter_goal else ""
            return f"At the current trend{label}, sales are flat or declining, so this target may not be reached on the current trajectory."

        monthly_growth_rate = trend_pct_goal / 100 / 3  # same monthly_growth logic as forecast
        months_needed = 0
        projected = ma_goal
        while projected < target_value and months_needed < 120:
            months_needed += 1
            projected = ma_goal * (1 + monthly_growth_rate * months_needed)

        label = f" for {specific_filter_goal[1]} {specific_filter_goal[0]}" if specific_filter_goal else ""
        if months_needed >= 120:
            answer = f"At the current growth rate{label}, the ${target_value:,.0f} sales target is not reachable within a 10-year projection window."
        else:
            target_date = _get_next_month_label(months_needed)[-1]
            answer = f"""Goal Projection{label}

Current monthly sales: ${current_val:,.2f}
Recent trend: {trend_pct_goal:+.1f}% per period
Projected to reach ${target_value:,.0f} by: {target_date} (approximately {months_needed} months from now)

Note: This is a linear projection based on the recent trend continuing unchanged."""

        context.append({"question": question, "sql": "", "answer": answer, "result": ""})
        if len(context) > 10:
            context.pop(0)
        return answer

    metric, metric_name, n_months, timeframe, group_by, scenario, specific_filter = _detect_prediction_intent(question)
    print(f"  📊 Predicting: {metric_name} | {timeframe} | group_by: {group_by} | filter: {specific_filter}")

    # Route 2: Scenario analysis
    if scenario:
        return _handle_scenario(question, scenario)

    # Fetch historical data
    df, value_col = _get_monthly_data(metric=metric, group_by=group_by, specific_filter=specific_filter)

    if df is None or df.empty:
        return f"Could not fetch historical {metric_name} data for prediction."

    if specific_filter:
        flabel = "Prime Members" if specific_filter == ('prime_status', 'Prime') else ("Non-Prime Members" if specific_filter == ('prime_status', 'Non-Prime') else f"{specific_filter[1]} {specific_filter[0]}")
        filter_label = f" for {flabel}"
    else:
        filter_label = ""

    # Route 3: Group-by forecast (region/category/segment, no specific filter)
    if group_by:
        groups = df[group_by].unique()
        results_by_group = {}
        for grp in groups:
            grp_df = df[df[group_by] == grp].copy().reset_index(drop=True)
            if len(grp_df) < 3:
                continue
            vals = grp_df[value_col].astype(float).values
            window_size = 6 if metric in ['orders', 'customers'] else 3
            preds, ma, trend_pct, std = _moving_average_forecast(vals, n_months, window=window_size)
            results_by_group[grp] = {
                'predictions': preds,
                'ma': ma,
                'trend_pct': trend_pct
            }

        if not results_by_group:
            return f"Not enough historical data to predict {metric_name} by {group_by}."

        # For sub-category "decline" questions, sort ascending (worst first); else descending
        is_decline_question = any(w in q_lower for w in ['decline', 'underperform', 'worst', 'risk'])
        sorted_groups = sorted(
            results_by_group.items(),
            key=lambda x: x[1]['trend_pct'],
            reverse=not is_decline_question
        ) if (group_by in ('product_name', 'sub_category') and is_decline_question) else sorted(
            results_by_group.items(), key=lambda x: x[1]['predictions'][-1], reverse=True
        )

        # Cap to top/bottom 5 for high-cardinality groupings
        is_product_grouping = group_by == 'product_name'
        is_subcategory_grouping = group_by == 'sub_category'
        if is_product_grouping or is_subcategory_grouping:
            sorted_groups = sorted_groups[:5]

        next_labels_grp = _get_next_month_label(n_months)
        unit_prefix = "$" if metric_name in ['sales', 'profit'] else ""
        unit_suffix = "" if metric_name in ['sales', 'profit'] else " orders"
        group_label = "Product" if is_product_grouping else ("Sub-Category" if is_subcategory_grouping else group_by.title())

        if is_product_grouping and not is_decline_question:
            top_grp, top_data = sorted_groups[0]
            top_total = sum(top_data['predictions']) if n_months > 1 else top_data['predictions'][0]
            header = f"Likely bestseller for {timeframe}: {top_grp} ({unit_prefix}{top_total:,.2f}{unit_suffix} predicted)"
            lines = [header, "", "Top 5 products by predicted sales:"]
            for grp, data in sorted_groups:
                total_pred = sum(data['predictions']) if n_months > 1 else data['predictions'][0]
                lines.append(f"{grp}: {unit_prefix}{total_pred:,.2f}{unit_suffix}")
        elif is_decline_question and (is_product_grouping or is_subcategory_grouping):
            worst_grp, worst_data = sorted_groups[0]
            header = f"Most likely to decline for {timeframe}: {worst_grp} ({worst_data['trend_pct']:+.1f}% trend)"
            lines = [header, "", f"Bottom 5 {group_label.lower()}s by trend:"]
            for grp, data in sorted_groups:
                total_pred = sum(data['predictions']) if n_months > 1 else data['predictions'][0]
                lines.append(f"{grp}: {unit_prefix}{total_pred:,.2f}{unit_suffix} (trend: {data['trend_pct']:+.1f}%)")
        elif n_months == 1:
            header = f"Predicted {metric_name.title()} by {group_label} for {next_labels_grp[0]}:"
            lines = [header, ""]
            for grp, data in sorted_groups:
                direction = "up" if data['trend_pct'] > 0 else "down"
                lines.append(f"{grp}: {unit_prefix}{data['predictions'][0]:,.2f}{unit_suffix} (trend: {direction} {abs(data['trend_pct']):.1f}%)")
        else:
            header = f"Predicted {metric_name.title()} by {group_label} for {timeframe}:"
            lines = [header, ""]
            for grp, data in sorted_groups:
                total_pred = sum(data['predictions'])
                lines.append(f"{grp}: {unit_prefix}{total_pred:,.2f}{unit_suffix} total (trend: {'up' if data['trend_pct'] > 0 else 'down'} {abs(data['trend_pct']):.1f}%)")
        forecast_text = "\n".join(lines)

        explain_prompt = f"""You are a Sales Analytics expert for Superstore.

Predicted {metric_name} by {group_label} for {timeframe}:
{chr(10).join([f"{g}: {unit_prefix}{(sum(d['predictions']) if n_months > 1 else d['predictions'][0]):,.2f}{unit_suffix} (trend: {d['trend_pct']:+.1f}%)" for g, d in sorted_groups])}

These are SALES trends, not price changes. Never mention "price" or "cost decreasing" — only discuss sales volume going up or down.
These are MONTHLY trends, not yearly. Never say "years" or "previous years" — only say "recent months".
Give exactly 3 one-line reasons based ONLY on these numbers. Do not invent reasons.
Format:
1. [reason]
2. [reason]
3. [reason]
No emojis. Never start with I."""

        explanation = call_model(explain_prompt, model=FAST_MODEL, temperature=0.1, num_ctx=1024, timeout=120)
        explanation = explanation.strip()
        answer = forecast_text + '\n\nWhy this prediction?\n' + explanation

    # Route 4: Overall or filtered (single region/category) forecast
    else:
        if len(df) < 3:
            return f"Not enough data to forecast{filter_label}. Need at least 3 months of {metric_name} history."

        values = df[value_col].astype(float).values
        next_labels = _get_next_month_label(n_months)
        window_size = 6 if metric in ['orders', 'customers'] else 3
        predictions, ma, trend_pct, std = _moving_average_forecast(values, n_months, window=window_size)
        pct_vs_last = ((predictions[0] - values[-1]) / values[-1] * 100) if values[-1] != 0 else 0
        direction = "increase" if pct_vs_last > 0 else "decrease"
        trend_word = "growing" if trend_pct > 0 else "declining"

        last_3 = df.tail(3)
        unit_prefix_for_str = "$" if metric in ['sales', 'profit','refund_amount'] else ""
        last_3_str = ", ".join([f"{row['month']}: {unit_prefix_for_str}{row[value_col]:,.0f}" for _, row in last_3.iterrows()])

        if n_months == 1:
            unit_prefix_main = "$" if metric in ['sales', 'profit','refund_amount'] else ""
            forecast_text = f"""Forecast{filter_label} -- {next_labels[0]}

Predicted {metric_name.title()}: {unit_prefix_main}{predictions[0]:,.2f}
Last month actual: {unit_prefix_main}{values[-1]:,.2f}
Expected {direction}: {abs(pct_vs_last):.1f}% vs last month
3-month average used: {unit_prefix_main}{ma:,.2f}
Trend direction: {metric_name.title()} is {trend_word} at {abs(trend_pct):.1f}% over recent months"""
        else:
            unit_prefix_main = "$" if metric in ['sales', 'profit','refund_amount'] else ""
            # If timeframe is a specific month (not "next quarter"/"next 6 months"/"next year"),
            # highlight that final target month clearly at the top
            is_specific_target = timeframe not in ['next quarter (3 months)', 'next 6 months', 'next year']

            if is_specific_target:
                target_pred = predictions[-1]
                target_label = next_labels[-1]
                header2 = f"Predicted {metric_name.title()}{filter_label} for {target_label}: {unit_prefix_main}{target_pred:,.2f}"
                lines = [header2, ""]
                lines.append("Month-by-month path to get there:")
                for pred, label in zip(predictions, next_labels):
                    marker = "  <-- target" if label == target_label else ""
                    lines.append(f'{label}: {unit_prefix_main}{pred:,.2f}{marker}')
            else:
                header2 = f"{metric_name.title()} Forecast{filter_label} for {timeframe}:"
                lines = [header2, ""]
                for pred, label in zip(predictions, next_labels):
                    d = 'up' if pred > ma else 'down'
                    lines.append(f'{label}: {unit_prefix_main}{pred:,.2f} ({d})')

            forecast_text = '\n'.join(lines)
            forecast_text += f'\n\nBase average (last 3 months): {unit_prefix_main}{ma:,.2f}'
        # Chart
        try:
            global latest_chart_buf
            fig, ax = plt.subplots(figsize=(12, 6))
            fig.patch.set_facecolor('#1c1915')
            ax.set_facecolor('#1c1915')

            hist_labels = df['month'].values
            show_from = max(0, len(values) - 12)
            ax.plot(range(len(values[show_from:])), values[show_from:],
                    color='#2E75B6', linewidth=2, marker='o', markersize=5, label='Historical')

            forecast_x = range(len(values[show_from:]), len(values[show_from:]) + n_months)
            ax.plot(list(forecast_x), predictions, color='#f59e0b', linewidth=2.5,
                    marker='o', markersize=7, linestyle='--', label='Forecast')

            ax.fill_between(list(forecast_x),
                             [p - std for p in predictions],
                             [p + std for p in predictions],
                             alpha=0.15, color='#f59e0b', label='Uncertainty band')

            all_labels = list(hist_labels[show_from:]) + next_labels
            tick_step = max(1, len(all_labels) // 10)
            ax.set_xticks(range(0, len(all_labels), tick_step))
            ax.set_xticklabels([all_labels[i] for i in range(0, len(all_labels), tick_step)],
                                rotation=45, ha='right', color='#aaaaaa', fontsize=9)

            title_suffix = filter_label if filter_label else ""
            ax.set_title(f'{metric_name.title()} Forecast{title_suffix} -- {timeframe}', fontsize=14,
                         fontweight='bold', color='#e2e8f0', pad=15)
            ax.set_ylabel(metric_name.title(), fontsize=11, color='#aaaaaa')
            ax.legend(labelcolor='#aaaaaa', facecolor='#1c1915', edgecolor='#333333')
            ax.grid(axis='y', alpha=0.1, linestyle='--', color='white')
            ax.spines['top'].set_visible(False)
            ax.spines['right'].set_visible(False)
            ax.spines['left'].set_color('#333333')
            ax.spines['bottom'].set_color('#333333')
            ax.tick_params(colors='#aaaaaa')
            plt.tight_layout()

            buf = io.BytesIO()
            plt.savefig(buf, format='png', dpi=150, bbox_inches='tight')
            plt.close()
            buf.seek(0)
            latest_chart_buf = buf
            print(f"  📊 Forecast chart generated!")
        except Exception as e:
            print(f"  ⚠️ Chart error: {e}")

        target_pred_explain = predictions[-1]
        target_label_explain = next_labels[-1]
        explain_direction = "increase" if target_pred_explain > ma else "decrease"

        explain_prompt = f"""You are a Sales Analytics expert for Superstore.

Predicted {metric_name}{filter_label} for {target_label_explain}: {unit_prefix_main}{target_pred_explain:,.2f}
Last 3 months actual: {last_3_str}
Base average (last 3 months): {unit_prefix_main}{ma:,.2f}
Overall trend: {trend_word} at {abs(trend_pct):.1f}% per month over recent months

Give exactly 3 short one-line business reasons why {metric_name} is expected to {explain_direction} by {target_label_explain}.
Base your reasons only on the numbers above. Do not invent reasons.
Do not mention moving average or technical terms.
No emojis. Never start with I."""

        explanation = call_model(explain_prompt, model=FAST_MODEL, temperature=0.1, num_ctx=1024, timeout=120)
        explanation = explanation.strip()
        answer = forecast_text + '\n\nWhy this prediction?\n' + explanation

    context.append({"question": question, "sql": "", "answer": answer, "result": ""})
    if len(context) > 10:
        context.pop(0)

    return answer



# ══════════════════════════════════════════════════════════════
# WORD DOC EXPORT
# ══════════════════════════════════════════════════════════════

def save_to_word(context_history):
    if not DOCX_AVAILABLE:
        return "  ⚠️  python-docx not installed. Run: pip install python-docx"
    if not context_history:
        return "  ⚠️  No conversation to save yet. Ask some questions first!"

    try:
        doc = Document()
        title = doc.add_heading('Superstore Sales Chat Report', 0)
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        date_para = doc.add_paragraph(f'Generated: {datetime.now().strftime("%Y-%m-%d %H:%M")}')
        date_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        total_para = doc.add_paragraph(f'Total Questions: {len(context_history)}')
        total_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph('')

        doc.add_heading('Chat Summary', level=1)
        summary_table = doc.add_table(rows=1, cols=2)
        summary_table.style = 'Light Grid Accent 1'
        hdr = summary_table.rows[0].cells
        hdr[0].text = 'Question'
        hdr[1].text = 'Answer'
        for cell in hdr:
            cell.paragraphs[0].runs[0].bold = True

        for exchange in context_history:
            row = summary_table.add_row().cells
            row[0].text = exchange['question']
            row[1].text = exchange['answer'].split('\n\n  💡')[0].strip()

        doc.add_paragraph('')
        doc.add_heading('Detailed Analysis', level=1)

        for i, exchange in enumerate(context_history):
            doc.add_heading(f'Q{i+1}: {exchange["question"]}', level=2)
            doc.add_paragraph(exchange['answer'].split('\n\n  💡')[0].strip())
            sql_para = doc.add_paragraph()
            sql_run = sql_para.add_run(f'SQL: {exchange.get("sql", "")}')
            sql_run.font.size = Pt(8)
            sql_run.font.color.rgb = RGBColor(0x88, 0x88, 0x88)
            doc.add_paragraph('─' * 60)

        filename = f"chat_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
        filepath = os.path.join(os.path.dirname(os.path.abspath(__file__)), filename)
        doc.save(filepath)
        return f"  📄 Report saved: {filename} ({len(context_history)} questions)"

    except Exception as e:
        return f"  ❌ Could not save Word doc: {e}"


# ══════════════════════════════════════════════════════════════
# MAIN ask() FUNCTION — Clean intent-based routing
# ══════════════════════════════════════════════════════════════

def ask(question):
    global context
    print(f"\n  ┌─ Question: {question}")

    # ── Database info ─────────────────────────────────────────
    database_info_phrases = [
        "what is this database", "what database", "about this database",
        "what data", "what does this contain", "what can you answer",
        "tell me about this database", "what do you know",
        "what is the database name", "database name",
        "what are you", "what do you do", "how can you help",
        "what is this", "which database"
    ]
    if any(phrase in question.lower() for phrase in database_info_phrases):
        return """This is the Superstore Sales Database (2022-2026):

  Orders     — 9,994 sales transactions
  Returns    — 1,999 product returns
  Customers  — 793 customers with demographics (age, sex, prime status, payment, channel, device)
  Geography  — 4 regions, 49 US states
  Products   — Furniture, Technology, Office Supplies

Ask me anything about sales, profit, returns, customer demographics, regions, products and trends!"""

    # ── Export trigger ────────────────────────────────────────
    export_phrases = ["download", "save", "export", "word doc", "word document", "generate report"]
    real_ctx = [ex for ex in context if ex.get("question") != "__summary__"]

    if any(phrase in question.lower() for phrase in export_phrases):
        return save_to_word(real_ctx) if real_ctx else "No conversation to save yet!"

    # ── SQL override — only when no context exists ────────────
    is_sql_question = False
    if not real_ctx:
        sql_override_phrases = [
        "what products", "which products", "list products",
        "what did", "what has", "what have", "which items",
        "how many", "how much", "what are", "show me",
        "what was", "who bought", "has bought", "have bought",
        "has purchased", "have purchased", "has ordered",
        "purchased by", "bought by", "ordered by",
        "give me all details", "all details", "details of",
        "tell me about", "info on", "information on"
        ]
        is_sql_question = any(phrase in question.lower() for phrase in sql_override_phrases)

    # ── Customer recommendation trigger ──────────────────────
    # Only fires if NOT a SQL/data question
    recommendation_triggers = [
        "recommend for", "recommendation for",
        "recommend products for", "recommend product for",
        "you may also", "you may buy",
        "suggest products for", "suggest for",
        "what would you recommend for",
        "what should we recommend for",
        "other customers also bought",
        "similar customers bought",
        "recommend him", "recommend her", "recommend them","recommend products for age",
        "recommend for age",
        "recommend products for prime",
        "recommend products for non-prime",
        "recommend products for mobile",
        "recommend products for web",
        "recommend products for android",
        "recommend products for ios",
        "recommend for age group",
        "recommend products for age group","will buy next", "buy next", "what will", "products will", "which product will", "which products will", "what product will","what will he buy next"
    ]

    if any(trigger in question.lower() for trigger in recommendation_triggers):
        try:
            with engine.connect() as conn:
                customers = conn.execute(text(
                    "SELECT customer_name FROM customers"
                )).fetchall()
                customer_names = [row[0] for row in customers]

            found_customer = None
            q_lower = question.lower()
            for name in customer_names:
                if name.lower() in q_lower:
                    found_customer = name
                    break

            if found_customer:
                print(f"  🛍️ Customer recommendation for: {found_customer}")
                return handle_customer_recommendation(question, found_customer)
        except Exception as e:
            print(f"  ⚠️ Customer lookup failed: {e}")

    # ── Predictive trigger — check before classify_intent ────
    predictive_triggers = [
        'predict', 'forecast', 'next month', 'next quarter', 'next year',
        'will sales', 'will profit', 'will return', 'will orders',
        'expected sales', 'expected profit',
        'future sales', 'future profit', 'projection',
        'if we reduce', 'if we increase', 'if sales grow',
        'when will', 'peak month', 'likely to grow', 'likely to decline',
        'compared to last year', 'vs last year', 'year over year',
        'this time next year', 'will it hit', 'at risk of',
        'churn next', 'will demand', 'how many new customers',
        'next 3 months', 'next 6 months', 'next 12 months', 'will improve',
        'will underperform', 'will decline', 'will grow', 'will increase',
        'will decrease', 'return rate drops'
    ]
    if any(t in question.lower() for t in predictive_triggers):
        print(f"  🔮 Intent: predictive")
        return handle_predictive(question)

    # ── Classify intent ───────────────────────────────────────
    intent = classify_intent(question)
    print(f"  🎯 Intent: {intent}")

    # ── Route to handler ──────────────────────────────────────
    if intent == "generic":
        return handle_generic(question)

    elif intent == "contextual":
        # Pass full context (includes summary entry) — handle_contextual extracts both
        return handle_contextual(question, context)

    elif intent == "recommendation":
        answer, relevant_data = handle_recommendation(question, context)
        context.append({"question": question, "sql": "", "answer": answer, "result": relevant_data})
        if len(context) > 10: context.pop(0)
        return answer

    elif intent == "predictive":
        return handle_predictive(question)

    elif intent == "subjective_sql":
        return handle_subjective_sql(question, context)

    else:  # sql (default)
        return handle_sql(question)


# ══════════════════════════════════════════════════════════════
# TERMINAL LOOP
# ══════════════════════════════════════════════════════════════

if __name__ == "__main__":
    print()
    print("  Welcome to the Superstore Sales Database Chatbot!")
    print()
    print("  Commands: 'exit' | 'clear' | 'context' | 'download'")
    print()
    print("=" * 60)

    while True:
        print()
        question = input("  You: ")

        if question.strip().lower() == "exit":
            print("  Goodbye!")
            break
        if question.strip().lower() == "clear":
            context = []
            print("  ✅ Conversation history cleared!")
            continue
        if question.strip().lower() == "context":
            if not context:
                print("  No conversation history yet!")
            else:
                for i, ex in enumerate(context):
                    print(f"\n  {i+1}. Q: {ex['question']}\n     A: {ex['answer']}")
            continue
        if question.strip() == "":
            print("  ⚠️  Please type a question!")
            continue

        print("  🤖 Thinking...")
        answer = ask(question)
        print(f"\n  └─ Answer:")
        wrapped = textwrap.fill(answer, width=70, initial_indent="     ", subsequent_indent="     ")
        print(wrapped)
        print()
        print("  " + "─" * 56)