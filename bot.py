import os
import logging
import tempfile
import json
import httpx
from datetime import datetime, timedelta, timezone
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

# Tavily web search
TAVILY_API_KEY = os.environ.get("TAVILY_API_KEY", "")

# Token storage
_ms_tokens: dict = {}

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
                "days": 3,
            },
        )
        resp.raise_for_status()
        return resp.json()


# ---------------------------------------------------------------------------
# Microsoft Graph helpers — NOTE: all OData params embedded in URL directly
# to avoid httpx percent-encoding the $ sign
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


async def ms_get_user_email() -> str:
    """Get the primary mailbox email for the org."""
    token = await ms_get_token()
    async with httpx.AsyncClient(timeout=30) as client:
        resp = await client.get(
            "https://graph.microsoft.com/v1.0/users?$select=mail,displayName&$top=1",
            headers={"Authorization": f"Bearer {token}"},
        )
        resp.raise_for_status()
        users = resp.json().get("value", [])
        if not users:
            raise ValueError("No users found in directory")
        return users[0]["mail"]


async def get_outlook_emails(max_emails: int = 10) -> list:
    token = await ms_get_token()
    user_email = await ms_get_user_email()
    async with httpx.AsyncClient(timeout=30) as client:
        url = (
            f"https://graph.microsoft.com/v1.0/users/{user_email}/mailFolders/inbox/messages"
            f"?$top={max_emails}&$select=subject,from,receivedDateTime,isRead,bodyPreview,importance"
            f"&$orderby=receivedDateTime desc&$filter=isRead eq false"
        )
        resp = await client.get(url, headers={"Authorization": f"Bearer {token}"})
        resp.raise_for_status()
        return resp.json().get("value", [])


async def get_calendar_events(days_ahead: int = 7) -> list:
    token = await ms_get_token()
    user_email = await ms_get_user_email()
    now = datetime.now(timezone.utc)
    end = now + timedelta(days=days_ahead)
    start_str = now.strftime("%Y-%m-%dT%H:%M:%SZ")
    end_str = end.strftime("%Y-%m-%dT%H:%M:%SZ")
    async with httpx.AsyncClient(timeout=30) as client:
        url = (
            f"https://graph.microsoft.com/v1.0/users/{user_email}/calendarView"
            f"?startDateTime={start_str}&endDateTime={end_str}"
            f"&$select=subject,start,end,location,organizer&$orderby=start/dateTime&$top=20"
        )
        resp = await client.get(url, headers={"Authorization": f"Bearer {token}"})
        resp.raise_for_status()
        return resp.json().get("value", [])


async def add_calendar_event(subject: str, start_dt: str, end_dt: str, description: str = "") -> dict:
    token = await ms_get_token()
    user_email = await ms_get_user_email()
    async with httpx.AsyncClient(timeout=30) as client:
        resp = await client.post(
            f"https://graph.microsoft.com/v1.0/users/{user_email}/events",
            headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
            json={
                "subject": subject,
                "body": {"contentType": "Text", "content": description},
                "start": {"dateTime": start_dt, "timeZone": "Australia/Brisbane"},
                "end": {"dateTime": end_dt, "timeZone": "Australia/Brisbane"},
                "isReminderOn": True,
                "reminderMinutesBeforeStart": 15,
            },
        )
        resp.raise_for_status()
        return resp.json()


# ---------------------------------------------------------------------------
# AI + Voice helpers
# ---------------------------------------------------------------------------

async def transcribe_voice(file_bytes: bytes) -> str:
    async with httpx.AsyncClient(timeout=60) as client:
        resp = await client.post(
            "https://api.openai.com/v1/audio/transcriptions",
            headers={"Authorization": f"Bearer {OPENAI_API_KEY}"},
            files={"file": ("voice.ogg", file_bytes, "audio/ogg")},
            data={"model": "whisper-1"},
        )
        resp.raise_for_status()
        return resp.json().get("text", "").strip()


async def call_mason(user_id: int, user_message: str) -> str:
    if user_id not in conversations:
        conversations[user_id] = []
    conversations[user_id].append({"role": "user", "content": user_message})
    history = conversations[user_id][-20:]
    async with httpx.AsyncClient(timeout=30) as client:
        resp = await client.post(
            "https://api.anthropic.com/v1/messages",
            headers={
                "x-api-key": ANTHROPIC_API_KEY,
                "anthropic-version": "2023-06-01",
                "content-type": "application/json",
            },
            json={
                "model": "claude-sonnet-4-5",
                "max_tokens": 500,
                "system": MASON_SYSTEM,
                "messages": history,
            },
        )
        resp.raise_for_status()
        data = resp.json()
    reply = " ".join(
        block["text"] for block in data.get("content", [])
        if block.get("type") == "text"
    ).strip()
    conversations[user_id].append({"role": "assistant", "content": reply})
    return reply


async def call_mason_with_context(user_id: int, user_message: str, context_text: str) -> str:
    enriched = f"{user_message}\n\n[Real-time data retrieved]:\n{context_text}"
    return await call_mason(user_id, enriched)


async def text_to_voice(text: str) -> bytes:
    clean = text.replace("**", "").replace("*", "").replace("#", "").strip()
    async with httpx.AsyncClient(timeout=30) as client:
        resp = await client.post(
            f"https://api.elevenlabs.io/v1/text-to-speech/{ELEVENLABS_VOICE_ID}",
            headers={"xi-api-key": ELEVENLABS_API_KEY, "Content-Type": "application/json"},
            json={
                "text": clean,
                "model_id": "eleven_turbo_v2_5",
                "voice_settings": {
                    "stability": 0.55,
                    "similarity_boost": 0.80,
                    "style": 0.25,
                    "use_speaker_boost": True,
                },
            },
        )
        resp.raise_for_status()
        return resp.content


# ---------------------------------------------------------------------------
# Intent detection
# ---------------------------------------------------------------------------

def detect_search_intent(text: str) -> bool:
    lower = text.lower()
    triggers = ["search", "look up", "find out", "what's happening", "latest news",
                "news on", "news about", "current", "today", "this week", "recent",
                "update on", "research", "who is", "what happened", "price of",
                "stock", "market", "weather", "asx", "bitcoin", "crypto"]
    return any(t in lower for t in triggers)


def detect_news_intent(text: str) -> bool:
    lower = text.lower()
    triggers = ["news", "headlines", "what's happening", "latest", "today's news", "briefing"]
    return any(t in lower for t in triggers)


def detect_email_intent(text: str) -> bool:
    lower = text.lower()
    return any(w in lower for w in ["check my email", "my inbox", "unread email", "read my email",
                                     "email update", "email briefing", "emails", "check email"])


def detect_calendar_intent(text: str) -> bool:
    lower = text.lower()
    return any(w in lower for w in ["my calendar", "my schedule", "what's on", "upcoming meetings",
                                     "meetings today", "meetings this week", "diary", "agenda", "calendar"])


# ---------------------------------------------------------------------------
# Telegram handlers
# ---------------------------------------------------------------------------

async def is_authorised(update: Update) -> bool:
    uid = update.effective_user.id if update.effective_user else 0
    if ALLOWED_USER_ID and uid != ALLOWED_USER_ID:
        await update.effective_message.reply_text("Blocked.")
        return False
    return True


async def cmd_search(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not await is_authorised(update):
        return
    if not context.args:
        await update.message.reply_text("Usage: /search <topic>")
        return
    query = " ".join(context.args)
    await update.message.reply_text(f"Searching: {query}...")
    try:
        data = await tavily_search(query, search_depth="advanced", max_results=5)
        answer = data.get("answer", "")
        results = data.get("results", [])
        ctx = answer + "\n\n" + "\n".join(f"- {r.get('title','')}: {r.get('content','')[:200]}" for r in results[:5])
        mason_prompt = f"Dan asked you to research: '{query}'. Here's what the web says — give him a sharp, actionable briefing:\n\n{ctx}"
        reply = await call_mason(update.effective_user.id, mason_prompt)
    except Exception as e:
        await update.message.reply_text(f"Search failed: {e}")
        return
    await update.message.reply_text(f"*{query}*\n\n{reply}", parse_mode="Markdown")
    try:
        audio = await text_to_voice(reply)
        with tempfile.NamedTemporaryFile(suffix=".mp3", delete=False) as f:
            f.write(audio)
            f.flush()
            with open(f.name, "rb") as af:
                await update.message.reply_voice(voice=af)
    except Exception:
        pass


async def cmd_news(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not await is_authorised(update):
        return
    topic = " ".join(context.args) if context.args else "Australia business technology news today"
    await update.message.reply_text(f"Pulling news: {topic}...")
    try:
        data = await tavily_news(topic, max_results=7)
        answer = data.get("answer", "")
        results = data.get("results", [])
        ctx = answer + "\n\n"
        for r in results[:7]:
            title = r.get("title", "")
            content = r.get("content", "")[:200]
            published = r.get("published_date", "")[:10] if r.get("published_date") else ""
            ctx += f"- [{published}] {title}: {content}\n"
        mason_prompt = f"Dan wants a news briefing on: '{topic}'. Here are the latest headlines — give him a punchy executive summary, flag anything that matters for business or strategy:\n\n{ctx}"
        reply = await call_mason(update.effective_user.id, mason_prompt)
    except Exception as e:
        await update.message.reply_text(f"News fetch failed: {e}")
        return
    await update.message.reply_text(f"*News: {topic}*\n\n{reply}", parse_mode="Markdown")
    try:
        audio = await text_to_voice(reply)
        with tempfile.NamedTemporaryFile(suffix=".mp3", delete=False) as f:
            f.write(audio)
            f.flush()
            with open(f.name, "rb") as af:
                await update.message.reply_voice(voice=af)
    except Exception:
        pass


async def cmd_emails(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not await is_authorised(update):
        return
    if not MS_CLIENT_ID:
        await update.message.reply_text("Outlook not configured. Add MS env vars to Railway.")
        return
    await update.message.reply_text("Checking your inbox...")
    try:
        emails = await get_outlook_emails(max_emails=10)
    except Exception as e:
        await update.message.reply_text(f"Couldn't fetch emails: {e}")
        return
    if not emails:
        await update.message.reply_text("No unread emails. Inbox clear, boss.")
        return
    email_text = ""
    for i, em in enumerate(emails, 1):
        sender = em.get("from", {}).get("emailAddress", {}).get("name", "Unknown")
        subject = em.get("subject", "(no subject)")
        preview = em.get("bodyPreview", "")[:150]
        received = em.get("receivedDateTime", "")[:10]
        importance = em.get("importance", "normal")
        star = " URGENT" if importance == "high" else ""
        email_text += f"{i}. From: {sender}{star}. Subject: {subject}. Date: {received}. Preview: {preview}\n\n"
    mason_prompt = f"Here are Dan's {len(emails)} unread emails. Give a sharp briefing — flag anything urgent, group by theme, tell me what needs action:\n\n{email_text}"
    summary = await call_mason(update.effective_user.id, mason_prompt)
    keyboard = []
    for i, em in enumerate(emails[:5]):
        subject = em.get("subject", "(no subject)")[:40]
        keyboard.append([InlineKeyboardButton(f"Add to calendar: {subject}", callback_data=f"cal_{i}")])
    reply_markup = InlineKeyboardMarkup(keyboard) if keyboard else None
    await update.message.reply_text(f"*Email Briefing*\n\n{summary}", parse_mode="Markdown", reply_markup=reply_markup)
    try:
        audio = await text_to_voice(summary)
        with tempfile.NamedTemporaryFile(suffix=".mp3", delete=False) as f:
            f.write(audio)
            f.flush()
            with open(f.name, "rb") as af:
                await update.message.reply_voice(voice=af)
    except Exception:
        pass
    context.bot_data["last_emails"] = emails


async def cmd_calendar_view(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not await is_authorised(update):
        return
    if not MS_CLIENT_ID:
        await update.message.reply_text("Outlook not configured. Add MS env vars to Railway.")
        return
    days = int(context.args[0]) if context.args and context.args[0].isdigit() else 7
    await update.message.reply_text(f"Checking your calendar for the next {days} days...")
    try:
        events = await get_calendar_events(days_ahead=days)
    except Exception as e:
        await update.message.reply_text(f"Couldn't fetch calendar: {e}")
        return
    if not events:
        await update.message.reply_text("Nothing on the calendar. Schedule's clear.")
        return
    cal_text = ""
    for ev in events:
        subject = ev.get("subject", "(no title)")
        start = ev.get("start", {}).get("dateTime", "")[:16].replace("T", " ")
        end = ev.get("end", {}).get("dateTime", "")[:16].replace("T", " ")
        location = ev.get("location", {}).get("displayName", "")
        loc_str = f" @ {location}" if location else ""
        cal_text += f"- {subject}: {start} to {end}{loc_str}\n"
    mason_prompt = f"Here's Dan's calendar for the next {days} days. Give him a quick briefing — what's coming up, anything he needs to prep for:\n\n{cal_text}"
    reply = await call_mason(update.effective_user.id, mason_prompt)
    await update.message.reply_text(f"*Calendar — Next {days} days*\n\n{reply}", parse_mode="Markdown")
    try:
        audio = await text_to_voice(reply)
        with tempfile.NamedTemporaryFile(suffix=".mp3", delete=False) as f:
            f.write(audio)
            f.flush()
            with open(f.name, "rb") as af:
                await update.message.reply_voice(voice=af)
    except Exception:
        pass


async def cmd_addcal(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not await is_authorised(update):
        return
    if not MS_CLIENT_ID:
        await update.message.reply_text("Outlook not configured.")
        return
    if not context.args:
        await update.message.reply_text("Usage: /addcal Title | YYYY-MM-DD HH:MM | duration_mins")
        return
    text = " ".join(context.args)
    parts = [p.strip() for p in text.split("|")]
    if len(parts) < 2:
        await update.message.reply_text("Format: /addcal Title | YYYY-MM-DD HH:MM | duration_minutes")
        return
    title = parts[0]
    dt_str = parts[1]
    duration = int(parts[2]) if len(parts) > 2 and parts[2].isdigit() else 60
    try:
        start = datetime.strptime(dt_str, "%Y-%m-%d %H:%M")
        end = start + timedelta(minutes=duration)
    except ValueError:
        await update.message.reply_text("Date format must be YYYY-MM-DD HH:MM")
        return
    try:
        await add_calendar_event(title, start.isoformat(), end.isoformat())
        await update.message.reply_text(f"Added: *{title}*\nStart: {dt_str} | Duration: {duration} mins", parse_mode="Markdown")
    except Exception as e:
        await update.message.reply_text(f"Failed to add event: {e}")


async def calendar_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    if not query.data.startswith("cal_"):
        return
    idx = int(query.data.split("_")[1])
    emails = context.bot_data.get("last_emails", [])
    if idx >= len(emails):
        await query.edit_message_text("Email not found.")
        return
    em = emails[idx]
    subject = em.get("subject", "Meeting")
    tomorrow = datetime.now() + timedelta(days=1)
    start = tomorrow.replace(hour=9, minute=0, second=0, microsecond=0)
    end = start + timedelta(hours=1)
    try:
        await add_calendar_event(subject, start.isoformat(), end.isoformat(), f"From email: {subject}")
        await query.edit_message_text(f"Added: {subject}\nScheduled: {start.strftime('%Y-%m-%d 09:00')} (1hr)")
    except Exception as e:
        await query.edit_message_text(f"Failed: {e}")


async def handle_voice(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not await is_authorised(update):
        return
    file = await context.bot.get_file(update.message.voice.file_id)
    file_bytes = await file.download_as_bytearray()
    try:
        text = await transcribe_voice(bytes(file_bytes))
    except Exception as e:
        await update.message.reply_text(f"Couldn't transcribe: {e}")
        return
    if not text:
        await update.message.reply_text("Couldn't make out what you said.")
        return
    await _process_message(update, context, text)


async def handle_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not await is_authorised(update):
        return
    text = update.message.text.strip()
    await _process_message(update, context, text)


async def _process_message(update: Update, context: ContextTypes.DEFAULT_TYPE, text: str):
    if detect_email_intent(text):
        context.args = []
        await cmd_emails(update, context)
        return
    if detect_calendar_intent(text):
        context.args = []
        await cmd_calendar_view(update, context)
        return
    if detect_news_intent(text):
        context.args = text.split()
        await cmd_news(update, context)
        return
    if detect_search_intent(text) and TAVILY_API_KEY:
        try:
            data = await tavily_search(text, search_depth="basic", max_results=4)
            answer = data.get("answer", "")
            results = data.get("results", [])
            ctx = answer + "\n" + "\n".join(f"- {r.get('title','')}: {r.get('content','')[:150]}" for r in results[:4])
            reply = await call_mason_with_context(update.effective_user.id, text, ctx)
        except Exception:
            reply = await call_mason(update.effective_user.id, text)
    else:
        reply = await call_mason(update.effective_user.id, text)
    try:
        audio = await text_to_voice(reply)
        with tempfile.NamedTemporaryFile(suffix=".mp3", delete=False) as f:
            f.write(audio)
            f.flush()
            with open(f.name, "rb") as af:
                await update.message.reply_voice(voice=af)
    except Exception:
        await update.message.reply_text(reply)


async def error_handler(update: object, context: ContextTypes.DEFAULT_TYPE):
    logger.error("Exception:", exc_info=context.error)
    if isinstance(update, Update) and update.effective_message:
        await update.effective_message.reply_text("Mason's offline briefly — back in a sec.")


def main():
    logger.info("Mason Drake is online.")
    app = Application.builder().token(TELEGRAM_TOKEN).build()
    app.add_handler(CommandHandler("emails", cmd_emails))
    app.add_handler(CommandHandler("calendar", cmd_calendar_view))
    app.add_handler(CommandHandler("addcal", cmd_addcal))
    app.add_handler(CommandHandler("search", cmd_search))
    app.add_handler(CommandHandler("news", cmd_news))
    app.add_handler(CallbackQueryHandler(calendar_callback, pattern="^cal_"))
    app.add_handler(MessageHandler(filters.VOICE, handle_voice))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text))
    app.add_error_handler(error_handler)
    app.run_polling(drop_pending_updates=True)


if __name__ == "__main__":
    main()
