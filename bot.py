import os
import logging
import tempfile
import json
import httpx
from datetime import datetime, timedelta, timezone
from zoneinfo import ZoneInfo
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import Application, MessageHandler, CommandHandler, CallbackQueryHandler, filters, ContextTypes

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# -- Env vars ----------------------------------------------------------------
TELEGRAM_TOKEN = os.environ["TELEGRAM_TOKEN"]
ANTHROPIC_API_KEY = os.environ["ANTHROPIC_API_KEY"]
ELEVENLABS_API_KEY = os.environ["ELEVENLABS_API_KEY"]
ELEVENLABS_VOICE_ID = os.environ["ELEVENLABS_VOICE_ID"]
OPENAI_API_KEY = os.environ["OPENAI_API_KEY"]
_allowed = os.environ.get("ALLOWED_USER_ID", "").strip()
ALLOWED_USER_ID = int(_allowed) if _allowed else 0

# Microsoft Graph (Outlook + Calendar)
MS_CLIENT_ID = os.environ.get("MS_CLIENT_ID", "")
MS_TENANT_ID = os.environ.get("MS_TENANT_ID", "")
MS_CLIENT_SECRET = os.environ.get("MS_CLIENT_SECRET", "")
MS_USER_EMAIL = os.environ.get("MS_USER_EMAIL", "")

# Tavily web search
TAVILY_API_KEY = os.environ.get("TAVILY_API_KEY", "")

# Token storage
_ms_tokens: dict = {}

SYDNEY_TZ = ZoneInfo("Australia/Sydney")

MASON_SYSTEM = """You are Mason Drake — Chief Operating Officer of Hellcat Industries, a high-performance consultancy run by Dan. You embody the Hellcat ethos: Built Different. Driven to Win.
Your mission: TO BUILD. TO LEAD. TO WIN. For our people. For our country. For our future.
Your five pillars:
- OPERATIONS: Optimise. Execute. Dominate.
- STRATEGY: Plan. Scale. Lead.
- LEADERSHIP: Empower. Inspire. Deliver.
- INTEGRITY: Loyal to the mission.
- IMPACT: Building a legacy.
Your three operational domains:
1. GENERAL OPS & STRATEGY — business direction, operational efficiency, competitive positioning, AI and automation leverage, decision support
2. PROJECT & TASK MANAGEMENT — tracking deliverables, scoping work, prioritising, identifying blockers, structuring plans
3. FINANCE & REPORTING — revenue, costs, margins, forecasting, reporting, financial health of the business
Your personality:
- Direct and decisive — say it once, say it right
- Commercially ruthless — every decision ties back to outcomes, revenue, or competitive edge
- Proactive — surface risks and opportunities before Dan asks
- Loyal to the mission — honest even when it's uncomfortable
- Australian grit — no corporate waffle, no excuses
Keep responses concise and punchy for voice — 3 to 5 sentences max unless Dan specifically asks for detail. No bullet points in speech. Natural, confident delivery. You are speaking, not writing."""

conversations = {}

# ---------------------------------------------------------------------------
# Tavily web search
# ---------------------------------------------------------------------------

async def tavily_search(query: str, search_depth: str = "basic", max_results: int = 5) -> dict:
    async with httpx.AsyncClient(timeout=30) as client:
        resp = await client.post(
            "https://api.tavily.com/search",
            headers={"Content-Type": "application/json"},
            json={
                "api_key": TAVILY_API_KEY,
                "query": query,
                "search_depth": search_depth,
                "max_results": max_results,
                "include_answer": True,
                "include_raw_content": False,
            },
        )
        resp.raise_for_status()
        return resp.json()

async def tavily_news(topic: str, max_results: int = 7) -> dict:
    async with httpx.AsyncClient(timeout=30) as client:
        resp = await client.post(
            "https://api.tavily.com/search",
            headers={"Content-Type": "application/json"},
            json={
                "api_key": TAVILY_API_KEY,
                "query": topic,
                "search_depth": "basic",
                "topic": "news",
                "max_results": max_results,
                "include_answer": True,
                "days": 1,
            },
        )
        resp.raise_for_status()
        return resp.json()

# ---------------------------------------------------------------------------
# Microsoft Graph helpers
# ---------------------------------------------------------------------------

async def ms_get_token() -> str:
    now = datetime.now(timezone.utc)
    if _ms_tokens.get("access_token") and _ms_tokens.get("expires_at", now) > now:
        return _ms_tokens["access_token"]
    async with httpx.AsyncClient(timeout=30) as client:
        resp = await client.post(
            f"https://login.microsoftonline.com/{MS_TENANT_ID}/oauth2/v2.0/token",
            data={
                "grant_type": "client_credentials",
                "client_id": MS_CLIENT_ID,
                "client_secret": MS_CLIENT_SECRET,
                "scope": "https://graph.microsoft.com/.default",
            },
        )
        resp.raise_for_status()
        data = resp.json()
        _ms_tokens["access_token"] = data["access_token"]
        _ms_tokens["expires_at"] = now + timedelta(seconds=data.get("expires_in", 3600) - 60)
        return _ms_tokens["access_token"]


async def ms_get_emails(max_emails: int = 10) -> list:
    if not MS_USER_EMAIL:
        raise ValueError("MS_USER_EMAIL environment variable not set.")
    token = await ms_get_token()
    headers = {"Authorization": f"Bearer {token}"}
    url = f"https://graph.microsoft.com/v1.0/users/{MS_USER_EMAIL}/messages?$top={max_emails}&$select=subject,from,receivedDateTime,bodyPreview,isRead&$orderby=receivedDateTime desc&$filter=isRead eq false"
    async with httpx.AsyncClient(timeout=30) as client:
        resp = await client.get(url, headers=headers)
        resp.raise_for_status()
        return resp.json().get("value", [])


async def ms_get_calendar_events(days_ahead: int = 7) -> list:
    if not MS_USER_EMAIL:
        raise ValueError("MS_USER_EMAIL environment variable not set.")
    token = await ms_get_token()
    headers = {"Authorization": f"Bearer {token}"}
    now = datetime.now(timezone.utc)
    end = now + timedelta(days=days_ahead)
    start_str = now.strftime("%Y-%m-%dT%H:%M:%SZ")
    end_str = end.strftime("%Y-%m-%dT%H:%M:%SZ")
    url = f"https://graph.microsoft.com/v1.0/users/{MS_USER_EMAIL}/calendarView?startDateTime={start_str}&endDateTime={end_str}&$top=20&$select=subject,start,end,location,organizer&$orderby=start/dateTime"
    async with httpx.AsyncClient(timeout=30) as client:
        resp = await client.get(url, headers=headers)
        resp.raise_for_status()
        return resp.json().get("value", [])


async def ms_create_calendar_event(subject: str, start_dt: str, end_dt: str, description: str = "") -> dict:
    if not MS_USER_EMAIL:
        raise ValueError("MS_USER_EMAIL environment variable not set.")
    token = await ms_get_token()
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    url = f"https://graph.microsoft.com/v1.0/users/{MS_USER_EMAIL}/events"
    body = {
        "subject": subject,
        "body": {"contentType": "text", "content": description},
        "start": {"dateTime": start_dt, "timeZone": "Australia/Sydney"},
        "end": {"dateTime": end_dt, "timeZone": "Australia/Sydney"},
    }
    async with httpx.AsyncClient(timeout=30) as client:
        resp = await client.post(url, headers=headers, json=body)
        resp.raise_for_status()
        return resp.json()

# ---------------------------------------------------------------------------
# Anthropic Claude
# ---------------------------------------------------------------------------

async def ask_claude(user_id: int, user_msg: str, system_override: str = None) -> str:
    history = conversations.setdefault(user_id, [])
    history.append({"role": "user", "content": user_msg})
    if len(history) > 20:
        history[:] = history[-20:]
    system = system_override or MASON_SYSTEM
    async with httpx.AsyncClient(timeout=60) as client:
        resp = await client.post(
            "https://api.anthropic.com/v1/messages",
            headers={
                "x-api-key": ANTHROPIC_API_KEY,
                "anthropic-version": "2023-06-01",
                "content-type": "application/json",
            },
            json={
                "model": "claude-opus-4-5",
                "max_tokens": 1024,
                "system": system,
                "messages": history,
            },
        )
        resp.raise_for_status()
        reply = resp.json()["content"][0]["text"]
    history.append({"role": "assistant", "content": reply})
    return reply


# ---------------------------------------------------------------------------
# ElevenLabs TTS
# ---------------------------------------------------------------------------

async def text_to_speech(text: str) -> bytes:
    async with httpx.AsyncClient(timeout=60) as client:
        resp = await client.post(
            f"https://api.elevenlabs.io/v1/text-to-speech/{ELEVENLABS_VOICE_ID}",
            headers={"xi-api-key": ELEVENLABS_API_KEY, "Content-Type": "application/json"},
            json={"text": text, "model_id": "eleven_turbo_v2", "voice_settings": {"stability": 0.5, "similarity_boost": 0.8}},
        )
        resp.raise_for_status()
        return resp.content


# ---------------------------------------------------------------------------
# OpenAI Whisper STT
# ---------------------------------------------------------------------------

async def speech_to_text(audio_bytes: bytes, mime: str = "audio/ogg") -> str:
    async with httpx.AsyncClient(timeout=60) as client:
        resp = await client.post(
            "https://api.openai.com/v1/audio/transcriptions",
            headers={"Authorization": f"Bearer {OPENAI_API_KEY}"},
            files={"file": ("audio.ogg", audio_bytes, mime)},
            data={"model": "whisper-1"},
        )
        resp.raise_for_status()
        return resp.json()["text"]


# ---------------------------------------------------------------------------
# Reply helpers
# ---------------------------------------------------------------------------

async def reply_voice(update: Update, text: str):
    audio = await text_to_speech(text)
    with tempfile.NamedTemporaryFile(suffix=".mp3", delete=False) as f:
        f.write(audio)
        f.flush()
        await update.message.reply_voice(voice=open(f.name, "rb"))


async def reply_smart(update: Update, text: str):
    await update.message.reply_text(text)
    try:
        await reply_voice(update, text)
    except Exception as e:
        logger.warning(f"TTS failed: {e}")


# ---------------------------------------------------------------------------
# Auth guard
# ---------------------------------------------------------------------------

def is_allowed(update: Update) -> bool:
    if not ALLOWED_USER_ID:
        return True
    return update.effective_user.id == ALLOWED_USER_ID


# ---------------------------------------------------------------------------
# Morning news briefing (scheduled 5:30 AM Sydney)
# ---------------------------------------------------------------------------

async def morning_briefing(context) -> None:
    if not ALLOWED_USER_ID:
        logger.warning("Morning briefing: ALLOWED_USER_ID not set, skipping.")
        return
    try:
        logger.info("Running morning news briefing...")
        # Fetch Australian news
        au_news = await tavily_news("Australia business economy politics news today", max_results=5)
        # Fetch world news
        world_news = await tavily_news("world news headlines today", max_results=5)

        au_answer = au_news.get("answer", "")
        au_articles = au_news.get("results", [])[:3]
        au_text = " ".join([f"{a.get('title', '')}: {a.get('content', '')[:150]}" for a in au_articles])

        world_answer = world_news.get("answer", "")
        world_articles = world_news.get("results", [])[:3]
        world_text = " ".join([f"{a.get('title', '')}: {a.get('content', '')[:150]}" for a in world_articles])

        now_sydney = datetime.now(SYDNEY_TZ).strftime("%A %d %B, %Y")

        prompt = (
            f"Good morning Dan. It's {now_sydney}. "
            f"Deliver a sharp 5:30 AM news briefing in your Mason Drake voice. "
            f"Cover the key Australian stories first, then the top world headlines. "
            f"Flag anything that could affect Hellcat Industries or Australian business. "
            f"Keep it punchy — this is a voice briefing, 6 to 8 sentences max. "
            f"Australian news: {au_answer} {au_text}. "
            f"World news: {world_answer} {world_text}."
        )

        reply = await ask_claude(ALLOWED_USER_ID, prompt)

        # Send text message
        await context.bot.send_message(chat_id=ALLOWED_USER_ID, text=reply)

        # Send voice message
        try:
            audio = await text_to_speech(reply)
            with tempfile.NamedTemporaryFile(suffix=".mp3", delete=False) as f:
                f.write(audio)
                f.flush()
                await context.bot.send_voice(chat_id=ALLOWED_USER_ID, voice=open(f.name, "rb"))
        except Exception as e:
            logger.warning(f"Morning briefing TTS failed: {e}")

        logger.info("Morning briefing sent successfully.")
    except Exception as e:
        logger.error(f"Morning briefing error: {e}")


# ---------------------------------------------------------------------------
# /start
# ---------------------------------------------------------------------------

async def cmd_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update):
        return
    await reply_smart(update, "Mason Drake online. What do you need?")


# ---------------------------------------------------------------------------
# /emails
# ---------------------------------------------------------------------------

async def cmd_emails(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update):
        return
    await update.message.reply_text("Checking your emails...")
    try:
        emails = await ms_get_emails(10)
        if not emails:
            await reply_smart(update, "No unread emails.")
            return
        summary_parts = []
        for e in emails[:5]:
            sender = e.get("from", {}).get("emailAddress", {}).get("name", "Unknown")
            subj = e.get("subject", "(no subject)")
            preview = e.get("bodyPreview", "")[:100]
            summary_parts.append(f"From {sender}: {subj}. {preview}")
        summary_text = " | ".join(summary_parts)
        prompt = f"Summarise these emails for Dan in your Mason Drake voice. Be concise and flag anything urgent: {summary_text}"
        reply = await ask_claude(update.effective_user.id, prompt)
        await reply_smart(update, reply)
        keyboard = [
            [InlineKeyboardButton("Add important items to calendar", callback_data="add_email_to_cal")],
            [InlineKeyboardButton("Done", callback_data="dismiss")],
        ]
        await update.message.reply_text("Want me to add anything to your calendar?", reply_markup=InlineKeyboardMarkup(keyboard))
    except Exception as e:
        logger.error(f"Email error: {e}")
        await update.message.reply_text(f"Couldn't fetch emails: {e}")


# ---------------------------------------------------------------------------
# /calendar
# ---------------------------------------------------------------------------

async def cmd_calendar(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update):
        return
    await update.message.reply_text("Checking your calendar for the next 7 days...")
    try:
        events = await ms_get_calendar_events(7)
        if not events:
            await reply_smart(update, "Nothing in your calendar for the next 7 days.")
            return
        event_parts = []
        for ev in events[:10]:
            subj = ev.get("subject", "(no subject)")
            start = ev.get("start", {}).get("dateTime", "")[:16].replace("T", " ")
            loc = ev.get("location", {}).get("displayName", "")
            event_parts.append(f"{subj} at {start}" + (f" ({loc})" if loc else ""))
        events_text = " | ".join(event_parts)
        prompt = f"Give Dan a sharp calendar briefing in your Mason Drake voice: {events_text}"
        reply = await ask_claude(update.effective_user.id, prompt)
        await reply_smart(update, reply)
    except Exception as e:
        logger.error(f"Calendar error: {e}")
        await update.message.reply_text(f"Couldn't fetch calendar: {e}")


# ---------------------------------------------------------------------------
# /addcal
# ---------------------------------------------------------------------------

async def cmd_addcal(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update):
        return
    args = context.args
    if not args:
        await update.message.reply_text(
            "Usage: /addcal Meeting with John | 2025-06-20T10:00 | 2025-06-20T11:00 | Description (optional)"
        )
        return
    parts = " ".join(args).split("|")
    if len(parts) < 3:
        await update.message.reply_text("Need at least: title | start datetime | end datetime")
        return
    subject = parts[0].strip()
    start_dt = parts[1].strip()
    end_dt = parts[2].strip()
    description = parts[3].strip() if len(parts) > 3 else ""
    try:
        await ms_create_calendar_event(subject, start_dt, end_dt, description)
        await reply_smart(update, f"Done. '{subject}' added to your calendar.")
    except Exception as e:
        logger.error(f"Add calendar error: {e}")
        await update.message.reply_text(f"Couldn't create event: {e}")


# ---------------------------------------------------------------------------
# /search
# ---------------------------------------------------------------------------

async def cmd_search(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update):
        return
    query = " ".join(context.args) if context.args else None
    if not query:
        await update.message.reply_text("Usage: /search your query here")
        return
    await update.message.reply_text(f"Searching for: {query}...")
    try:
        results = await tavily_search(query)
        answer = results.get("answer", "")
        sources = results.get("results", [])[:3]
        source_text = " ".join([f"{s.get('title', '')}: {s.get('content', '')[:150]}" for s in sources])
        prompt = f"Based on this research, give Dan a sharp answer in your Mason Drake voice. Query: {query}. Research: {answer} {source_text}"
        reply = await ask_claude(update.effective_user.id, prompt)
        await reply_smart(update, reply)
    except Exception as e:
        logger.error(f"Search error: {e}")
        await update.message.reply_text(f"Search failed: {e}")


# ---------------------------------------------------------------------------
# /news
# ---------------------------------------------------------------------------

async def cmd_news(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update):
        return
    topic = " ".join(context.args) if context.args else "business Australia economy technology"
    await update.message.reply_text(f"Getting news on: {topic}...")
    try:
        results = await tavily_news(topic)
        answer = results.get("answer", "")
        articles = results.get("results", [])[:5]
        article_text = " ".join([f"{a.get('title', '')}: {a.get('content', '')[:200]}" for a in articles])
        prompt = f"Give Dan a sharp news briefing in your Mason Drake voice. Focus on what matters for a high-performance business leader. Topics: {topic}. News: {answer} {article_text}"
        reply = await ask_claude(update.effective_user.id, prompt)
        await reply_smart(update, reply)
    except Exception as e:
        logger.error(f"News error: {e}")
        await update.message.reply_text(f"News fetch failed: {e}")


# ---------------------------------------------------------------------------
# Inline button callbacks
# ---------------------------------------------------------------------------

async def handle_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    if query.data == "dismiss":
        await query.edit_message_text("Got it.")
    elif query.data == "add_email_to_cal":
        await query.edit_message_text("Tell me what to add — e.g. /addcal Meeting title | 2025-06-20T10:00 | 2025-06-20T11:00")


# ---------------------------------------------------------------------------
# Voice message handler
# ---------------------------------------------------------------------------

async def handle_voice(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update):
        return
    file = await context.bot.get_file(update.message.voice.file_id)
    audio_bytes = bytes(await file.download_as_bytearray())
    try:
        text = await speech_to_text(audio_bytes)
    except Exception as e:
        await update.message.reply_text(f"Couldn't transcribe audio: {e}")
        return
    await handle_text_content(update, context, text)


# ---------------------------------------------------------------------------
# Text message handler
# ---------------------------------------------------------------------------

async def handle_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update):
        return
    text = update.message.text or ""
    await handle_text_content(update, context, text)


async def handle_text_content(update: Update, context: ContextTypes.DEFAULT_TYPE, text: str):
    lower = text.lower()
    if any(w in lower for w in ["email", "emails", "inbox", "unread", "messages"]):
        await cmd_emails(update, context)
    elif any(w in lower for w in ["calendar", "schedule", "meetings", "appointments", "agenda"]):
        await cmd_calendar(update, context)
    elif lower.startswith("search ") or lower.startswith("look up ") or lower.startswith("research "):
        query = text.split(" ", 1)[1] if " " in text else text
        context.args = query.split()
        await cmd_search(update, context)
    elif any(w in lower for w in ["news", "headlines", "briefing", "what's happening"]):
        context.args = []
        await cmd_news(update, context)
    else:
        reply = await ask_claude(update.effective_user.id, text)
        await reply_smart(update, reply)


# ---------------------------------------------------------------------------
# Error handler
# ---------------------------------------------------------------------------

async def error_handler(update: object, context: ContextTypes.DEFAULT_TYPE):
    logger.error(f"Update {update} caused error {context.error}")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    app = Application.builder().token(TELEGRAM_TOKEN).build()

    # -- Scheduled morning briefing: 5:30 AM Sydney every day --
    job_queue = app.job_queue
    job_queue.run_daily(
        morning_briefing,
        time=datetime.now(SYDNEY_TZ).replace(hour=5, minute=30, second=0, microsecond=0).timetz(),
        name="morning_briefing",
    )

    app.add_handler(CommandHandler("start", cmd_start))
    app.add_handler(CommandHandler("emails", cmd_emails))
    app.add_handler(CommandHandler("calendar", cmd_calendar))
    app.add_handler(CommandHandler("addcal", cmd_addcal))
    app.add_handler(CommandHandler("search", cmd_search))
    app.add_handler(CommandHandler("news", cmd_news))
    app.add_handler(CallbackQueryHandler(handle_callback))
    app.add_handler(MessageHandler(filters.VOICE, handle_voice))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text))
    app.add_error_handler(error_handler)
    app.run_polling(drop_pending_updates=True)


if __name__ == "__main__":
    main()
