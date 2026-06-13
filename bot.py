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

# Token storage (in-memory; persists while bot runs)
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
# Microsoft Graph helpers
# ---------------------------------------------------------------------------

async def ms_get_token() -> str:
    """Get a valid Microsoft Graph access token (client credentials flow for org accounts)."""
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


async def get_outlook_emails(max_emails: int = 10) -> list[dict]:
    """Fetch recent unread emails from Outlook inbox."""
    token = await ms_get_token()
    async with httpx.AsyncClient(timeout=30) as client:
        # Get the first mailbox user in the org (Dan's account)
        users_resp = await client.get(
            "https://graph.microsoft.com/v1.0/users",
            headers={"Authorization": f"Bearer {token}"},
            params={"$select": "mail,displayName", "$top": "1"},
        )
        users_resp.raise_for_status()
        users = users_resp.json().get("value", [])
        if not users:
            return []
        user_email = users[0]["mail"]

        resp = await client.get(
            f"https://graph.microsoft.com/v1.0/users/{user_email}/mailFolders/inbox/messages",
            headers={"Authorization": f"Bearer {token}"},
            params={
                "\$top": str(max_emails),
                "\$select": "subject,from,receivedDateTime,isRead,bodyPreview,importance",
                "\$orderby": "receivedDateTime desc",
                "\$filter": "isRead eq false",
            },
        )
        resp.raise_for_status()
        return resp.json().get("value", [])


async def add_calendar_event(subject: str, start_dt: str, end_dt: str, description: str = "") -> dict:
    """Add an event to Outlook Calendar. start_dt and end_dt are ISO 8601 strings."""
    token = await ms_get_token()
    async with httpx.AsyncClient(timeout=30) as client:
        users_resp = await client.get(
            "https://graph.microsoft.com/v1.0/users",
            headers={"Authorization": f"Bearer {token}"},
            params={"\$select": "mail", "\$top": "1"},
        )
        users_resp.raise_for_status()
        users = users_resp.json().get("value", [])
        if not users:
            return {"error": "No users found"}
        user_email = users[0]["mail"]

        resp = await client.post(
            f"https://graph.microsoft.com/v1.0/users/{user_email}/events",
            headers={
                "Authorization": f"Bearer {token}",
                "Content-Type": "application/json",
            },
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


async def text_to_voice(text: str) -> bytes:
    clean = text.replace("**", "").replace("*", "").replace("#", "").strip()
    async with httpx.AsyncClient(timeout=30) as client:
        resp = await client.post(
            f"https://api.elevenlabs.io/v1/text-to-speech/{ELEVENLABS_VOICE_ID}",
            headers={
                "xi-api-key": ELEVENLABS_API_KEY,
                "Content-Type": "application/json",
            },
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
# Telegram handlers
# ---------------------------------------------------------------------------

async def is_authorised(update: Update) -> bool:
    uid = update.effective_user.id if update.effective_user else 0
    if ALLOWED_USER_ID and uid != ALLOWED_USER_ID:
        await update.effective_message.reply_text("Blocked.")
        return False
    return True


async def cmd_emails(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Handle /emails command — fetch and summarise unread Outlook emails."""
    if not await is_authorised(update):
        return

    if not MS_CLIENT_ID:
        await update.message.reply_text("Outlook not configured yet. Add MS_CLIENT_ID, MS_TENANT_ID, MS_CLIENT_SECRET to Railway.")
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

    # Build a summary and ask Mason to analyse it
    email_text = ""
    for i, em in enumerate(emails, 1):
        sender = em.get("from", {}).get("emailAddress", {}).get("name", "Unknown")
        subject = em.get("subject", "(no subject)")
        preview = em.get("bodyPreview", "")[:150]
        received = em.get("receivedDateTime", "")[:10]
        importance = em.get("importance", "normal")
        star = " ⭐" if importance == "high" else ""
        email_text += f"{i}. From: {sender}{star}\nSubject: {subject}\nDate: {received}\nPreview: {preview}\n\n"

    mason_prompt = f"Here are Dan's {len(emails)} unread emails. Give a sharp briefing — flag anything urgent or important, group by theme if needed, and tell me what needs action:\n\n{email_text}"
    summary = await call_mason(update.effective_user.id, mason_prompt)

    # Build inline keyboard for calendar actions
    keyboard = []
    for i, em in enumerate(emails[:5]):  # Show buttons for first 5
        subject = em.get("subject", "(no subject)")[:40]
        keyboard.append([InlineKeyboardButton(f"📅 Add to calendar: {subject}", callback_data=f"cal_{i}")])

    reply_markup = InlineKeyboardMarkup(keyboard) if keyboard else None

    await update.message.reply_text(f"📬 **Email Briefing**\n\n{summary}", parse_mode="Markdown", reply_markup=reply_markup)

    # Also send as voice
    try:
        audio = await text_to_voice(summary)
        with tempfile.NamedTemporaryFile(suffix=".mp3", delete=False) as f:
            f.write(audio)
            f.flush()
            with open(f.name, "rb") as af:
                await update.message.reply_voice(voice=af)
    except Exception:
        pass  # Voice optional

    # Store emails for callback handler
    context.bot_data["last_emails"] = emails


async def cmd_calendar(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Handle /addcal command — add a custom event to calendar."""
    if not await is_authorised(update):
        return

    if not MS_CLIENT_ID:
        await update.message.reply_text("Outlook not configured. Add MS env vars to Railway.")
        return

    args = context.args
    if not args:
        await update.message.reply_text(
            "Usage: /addcal <title> | <YYYY-MM-DD HH:MM> | <duration mins>\n"
            "Example: /addcal Strategy meeting | 2026-06-14 09:00 | 60"
        )
        return

    text = " ".join(args)
    parts = [p.strip() for p in text.split("|")]
    if len(parts) < 2:
        await update.message.reply_text("Format: /addcal Title | YYYY-MM-DD HH:MM | duration_minutes")
        return

    title = parts[0]
    dt_str = parts[1]
    duration = int(parts[2]) if len(parts) > 2 else 60

    try:
        start = datetime.strptime(dt_str, "%Y-%m-%d %H:%M")
        end = start + timedelta(minutes=duration)
        start_iso = start.isoformat()
        end_iso = end.isoformat()
    except ValueError:
        await update.message.reply_text("Date format must be YYYY-MM-DD HH:MM")
        return

    try:
        result = await add_calendar_event(title, start_iso, end_iso)
        await update.message.reply_text(f"✅ Added to calendar: **{title}**\nStart: {dt_str}\nDuration: {duration} mins", parse_mode="Markdown")
    except Exception as e:
        await update.message.reply_text(f"Failed to add event: {e}")


async def calendar_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Handle inline button press to add email to calendar."""
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
    received_str = em.get("receivedDateTime", "")

    # Default: schedule for next day at 9am
    tomorrow = datetime.now() + timedelta(days=1)
    start = tomorrow.replace(hour=9, minute=0, second=0, microsecond=0)
    end = start + timedelta(hours=1)

    try:
        await add_calendar_event(subject, start.isoformat(), end.isoformat(), f"From email: {subject}")
        await query.edit_message_text(f"✅ Added to calendar: {subject}\nScheduled: {start.strftime('%Y-%m-%d 09:00')} (1hr)\nAdjust time with /addcal if needed.")
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

    # Check for email/calendar intents in voice
    lower = text.lower()
    if any(w in lower for w in ["email", "inbox", "mail", "messages"]):
        context.args = []
        await cmd_emails(update, context)
        return

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


async def handle_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not await is_authorised(update):
        return

    text = update.message.text.strip()

    # Check for email/calendar intents in text
    lower = text.lower()
    if any(w in lower for w in ["check my email", "my inbox", "unread email", "read my email", "email update", "email briefing"]):
        context.args = []
        await cmd_emails(update, context)
        return

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
    app.add_handler(CommandHandler("addcal", cmd_calendar))
    app.add_handler(CallbackQueryHandler(calendar_callback, pattern="^cal_"))
    app.add_handler(MessageHandler(filters.VOICE, handle_voice))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text))
    app.add_error_handler(error_handler)
    app.run_polling(drop_pending_updates=True)


if __name__ == "__main__":
    main()
