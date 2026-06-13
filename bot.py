import os
import logging
import tempfile
import httpx
from telegram import Update
from telegram.ext import Application, MessageHandler, filters, ContextTypes

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# -- Env vars (set these in Railway, never in code) --------------------------
TELEGRAM_TOKEN = os.environ["TELEGRAM_TOKEN"]
ANTHROPIC_API_KEY = os.environ["ANTHROPIC_API_KEY"]
ELEVENLABS_API_KEY = os.environ["ELEVENLABS_API_KEY"]
ELEVENLABS_VOICE_ID = os.environ["ELEVENLABS_VOICE_ID"]
OPENAI_API_KEY = os.environ["OPENAI_API_KEY"]  # for Whisper transcription
_allowed = os.environ.get("ALLOWED_USER_ID", "").strip()
ALLOWED_USER_ID = int(_allowed) if _allowed else 0  # your Telegram user ID

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

# Conversation history per user
conversations = {}


async def transcribe_voice(file_bytes: bytes) -> str:
    """Transcribe voice message using OpenAI Whisper."""
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
    """Call Anthropic API with Mason's persona."""
    if user_id not in conversations:
        conversations[user_id] = []

    conversations[user_id].append({"role": "user", "content": user_message})

    # Keep last 20 messages to manage context
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
    """Convert text to audio using ElevenLabs."""
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


async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Handle incoming text or voice messages."""
    user_id = update.effective_user.id

    # Security: only respond to Dan
    if ALLOWED_USER_ID and user_id != ALLOWED_USER_ID:
        logger.warning(f"Blocked unauthorized user {user_id}")
        return

    await context.bot.send_chat_action(
        chat_id=update.effective_chat.id, action="typing"
    )

    # Get the message text
    if update.message.voice:
        try:
            voice_file = await context.bot.get_file(update.message.voice.file_id)
            file_bytes = await voice_file.download_as_bytearray()
            user_text = await transcribe_voice(bytes(file_bytes))
            if not user_text:
                await update.message.reply_text("Couldn't catch that — try again.")
                return
        except Exception as e:
            logger.error(f"Transcription error: {e}")
            await update.message.reply_text("Voice transcription hit a snag. Try typing it for now.")
            return
    elif update.message.text:
        user_text = update.message.text
    else:
        return

    try:
        reply = await call_mason(user_id, user_text)

        await update.message.reply_text(reply)

        await context.bot.send_chat_action(
            chat_id=update.effective_chat.id, action="record_voice"
        )
        audio_bytes = await text_to_voice(reply)

        with tempfile.NamedTemporaryFile(suffix=".mp3", delete=False) as f:
            f.write(audio_bytes)
            tmp_path = f.name

        with open(tmp_path, "rb") as audio_file:
            await update.message.reply_voice(voice=audio_file)

        os.unlink(tmp_path)

    except Exception as e:
        logger.error(f"Error: {e}")
        await update.message.reply_text(
            "Mason's offline briefly. Try again in a moment."
        )


def main():
    app = Application.builder().token(TELEGRAM_TOKEN).build()
    app.add_handler(MessageHandler(filters.TEXT | filters.VOICE, handle_message))
    logger.info("Mason Drake is online.")
    app.run_polling(drop_pending_updates=True)


if __name__ == "__main__":
    main()
