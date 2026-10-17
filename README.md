# Superstore Sales Analytics Chatbot

An AI-powered conversational analytics assistant for retail sales data.

## Setup

1. Copy `env.example` to `.env` and fill in your values
2. Install dependencies: `pip install -r requirements.txt`
3. Run: `uvicorn api:app --host 0.0.0.0 --port 8000`

## Environment Variables Required

- `DB_URL` — PostgreSQL connection string
- `GROQ_API_KEY` — Groq API key
- `GROQ_SQL_MODEL` — e.g. `llama-3.3-70b-versatile`
- `GROQ_FAST_MODEL` — e.g. `llama-3.3-70b-versatile`
