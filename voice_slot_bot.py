"""
Combined Shopify webhook listener + Discord bot for 3-hour voice channel slots.

Flow:
  1. Customer buys a "3-hour slot" product on Shopify.
  2. Shopify sends an orders/create webhook -> this server generates a
     redemption code and stores it in SQLite.
  3. Customer runs /redeem <code> in Discord.
  4. Bot validates the code, moves them into the target voice channel,
     and starts their 3-hour timer.
  5. A background loop checks every 30s and disconnects anyone who has
     been in the channel longer than 3 hours.

Run:
  pip install discord.py flask python-dotenv
  python voice_slot_bot.py

Environment variables (put these in a .env file or your host's config):
  DISCORD_BOT_TOKEN     - bot token from the Discord Developer Portal
  SHOPIFY_WEBHOOK_SECRET- from Shopify Admin > Settings > Notifications > Webhooks
  VOICE_CHANNEL_ID      - the Discord voice channel ID slots unlock access to
  GUILD_ID              - your Discord server (guild) ID
  FLASK_PORT            - port for the webhook server (default 5000)
"""

import os
import hmac
import hashlib
import base64
import secrets
import sqlite3
import threading
import datetime
import asyncio

import discord
from discord import app_commands
from discord.ext import tasks
from flask import Flask, request, abort
from dotenv import load_dotenv

load_dotenv()

DISCORD_BOT_TOKEN = os.environ["DISCORD_BOT_TOKEN"]
SHOPIFY_WEBHOOK_SECRET = os.environ["SHOPIFY_WEBHOOK_SECRET"]
VOICE_CHANNEL_ID = int(os.environ["VOICE_CHANNEL_ID"])
GUILD_ID = int(os.environ["GUILD_ID"])
FLASK_PORT = int(os.environ.get("FLASK_PORT", 5000))

SLOT_DURATION = datetime.timedelta(hours=3)
DB_PATH = "slots.db"

# ---------------------------------------------------------------------------
# Database setup
# ---------------------------------------------------------------------------

def db_conn():
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS codes (
            code TEXT PRIMARY KEY,
            order_id TEXT,
            customer_email TEXT,
            used INTEGER DEFAULT 0,
            redeemed_by TEXT,
            redeemed_at TEXT,
            created_at TEXT
        )
    """)
    conn.commit()
    return conn

db_lock = threading.Lock()
conn = db_conn()

# ---------------------------------------------------------------------------
# Flask webhook server (runs in its own thread)
# ---------------------------------------------------------------------------

app = Flask(__name__)

def verify_shopify_hmac(data: bytes, hmac_header: str) -> bool:
    digest = hmac.new(SHOPIFY_WEBHOOK_SECRET.encode(), data, hashlib.sha256).digest()
    computed = base64.b64encode(digest).decode()
    return hmac.compare_digest(computed, hmac_header or "")

@app.route("/webhook/order", methods=["POST"])
def order_created():
    hmac_header = request.headers.get("X-Shopify-Hmac-Sha256", "")
    raw_body = request.get_data()

    if not verify_shopify_hmac(raw_body, hmac_header):
        abort(401)

    order = request.get_json(force=True)
    order_id = str(order.get("id"))
    email = order.get("email") or order.get("contact_email") or "unknown"

    code = secrets.token_hex(4).upper()  # e.g. "A1B2C3D4"

    with db_lock:
        conn.execute(
            "INSERT INTO codes (code, order_id, customer_email, created_at) VALUES (?, ?, ?, ?)",
            (code, order_id, email, datetime.datetime.utcnow().isoformat()),
        )
        conn.commit()

    # TODO: send the code to the customer here (email API call, or rely on
    # the Shopify order-confirmation email template to display it — see
    # shopify_email_template.liquid). For now it's just stored in the DB.
    print(f"[webhook] order {order_id} -> code {code} for {email}")

    return "ok", 200

def run_flask():
    app.run(host="0.0.0.0", port=FLASK_PORT)

# ---------------------------------------------------------------------------
# Discord bot
# ---------------------------------------------------------------------------

intents = discord.Intents.default()
intents.voice_states = True
intents.members = True

class SlotBot(discord.Client):
    def __init__(self):
        super().__init__(intents=intents)
        self.tree = app_commands.CommandTree(self)

    async def setup_hook(self):
        await self.tree.sync(guild=discord.Object(id=GUILD_ID))
        check_expired_slots.start()

client = SlotBot()

# user_id -> join_time (UTC)
active_sessions: dict[int, datetime.datetime] = {}

@client.tree.command(
    name="redeem",
    description="Redeem your 3-hour voice channel slot code",
    guild=discord.Object(id=GUILD_ID),
)
@app_commands.describe(code="The code from your order confirmation")
async def redeem(interaction: discord.Interaction, code: str):
    code = code.strip().upper()

    with db_lock:
        row = conn.execute(
            "SELECT used FROM codes WHERE code = ?", (code,)
        ).fetchone()

        if row is None:
            await interaction.response.send_message("That code isn't valid.", ephemeral=True)
            return
        if row[0] == 1:
            await interaction.response.send_message("That code has already been used.", ephemeral=True)
            return

        conn.execute(
            "UPDATE codes SET used = 1, redeemed_by = ?, redeemed_at = ? WHERE code = ?",
            (str(interaction.user.id), datetime.datetime.utcnow().isoformat(), code),
        )
        conn.commit()

    guild = client.get_guild(GUILD_ID)
    member = guild.get_member(interaction.user.id)
    channel = guild.get_channel(VOICE_CHANNEL_ID)

    if member is None or channel is None:
        await interaction.response.send_message(
            "Redeemed, but I couldn't find you or the channel. Contact an admin.", ephemeral=True
        )
        return

    active_sessions[member.id] = datetime.datetime.utcnow()

    if member.voice is not None:
        await member.move_to(channel)
    # If they're not already in a voice channel, Discord doesn't let bots
    # pull them in from nothing — send them the channel link/invite instead.

    await interaction.response.send_message(
        f"Code redeemed! Join {channel.mention} — your 3-hour window starts now.",
        ephemeral=True,
    )

@client.event
async def on_voice_state_update(member, before, after):
    # Clear the session if they leave the slot channel on their own.
    if before.channel and before.channel.id == VOICE_CHANNEL_ID:
        if not after.channel or after.channel.id != VOICE_CHANNEL_ID:
            active_sessions.pop(member.id, None)

@tasks.loop(seconds=30)
async def check_expired_slots():
    now = datetime.datetime.utcnow()
    guild = client.get_guild(GUILD_ID)
    if guild is None:
        return

    for user_id, joined in list(active_sessions.items()):
        if now - joined > SLOT_DURATION:
            member = guild.get_member(user_id)
            if member and member.voice and member.voice.channel and member.voice.channel.id == VOICE_CHANNEL_ID:
                await member.move_to(None)  # disconnects them from voice
            active_sessions.pop(user_id, None)

# ---------------------------------------------------------------------------
# Entrypoint: run Flask in a background thread, Discord bot in the main loop
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    flask_thread = threading.Thread(target=run_flask, daemon=True)
    flask_thread.start()
    client.run(DISCORD_BOT_TOKEN)
