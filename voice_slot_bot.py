"""
Combined Shopify webhook listener + Discord bot for token-based 3-hour
voice channel booking slots.

Flow:
  1. Customer buys N "tokens" on Shopify (quantity = number of tokens).
  2. Shopify sends an orders/create webhook -> this server generates one
     redemption code per token and emails them to the customer.
  3. Customer runs /redeem <code> in Discord -> adds 1 token to their balance.
  4. Customer runs /book <date> <slot> to spend 1 token reserving a fixed
     daily 3-hour window: 9-12, 12-3, or 3-6 (America/Toronto time).
  5. A background loop creates a private voice channel automatically when
     a booked slot starts, and deletes it automatically when it ends.

Run:
  pip install discord.py flask python-dotenv openpyxl
  python voice_slot_bot.py

Environment variables (put these in a .env file or your host's config):
  DISCORD_BOT_TOKEN     - bot token from the Discord Developer Portal
  SHOPIFY_WEBHOOK_SECRET- from Shopify Admin > Settings > Notifications > Webhooks
  GUILD_ID              - your Discord server (guild) ID
  FLASK_PORT            - port for the webhook server (default 5000, ignored on Railway)
  DB_PATH               - path to the SQLite file (default "slots.db"; on Railway, point this at a mounted Volume, e.g. /data/slots.db, or the database resets on every deploy)
  GMAIL_ADDRESS         - Gmail address to send redemption codes from
  GMAIL_APP_PASSWORD    - 16-character Gmail App Password (not your normal password)
  TOKEN_PRODUCT_ID      - Shopify product ID for the token product (only this product's line items generate codes)
  TUTOR_ROLE_ID         - Discord role ID for tutors; they get access to every student's private text channel
"""

import os
import hmac
import hashlib
import base64
import secrets
import sqlite3
import threading
import datetime
import smtplib
from email.mime.text import MIMEText
from zoneinfo import ZoneInfo

import discord
from discord import app_commands
from discord.ext import tasks
from flask import Flask, request, abort
from dotenv import load_dotenv
from openpyxl import Workbook

load_dotenv()

DISCORD_BOT_TOKEN = os.environ["DISCORD_BOT_TOKEN"]
SHOPIFY_WEBHOOK_SECRET = os.environ["SHOPIFY_WEBHOOK_SECRET"]
GUILD_ID = int(os.environ["GUILD_ID"])
FLASK_PORT = int(os.environ.get("FLASK_PORT", 5000))
GMAIL_ADDRESS = os.environ["GMAIL_ADDRESS"]
GMAIL_APP_PASSWORD = os.environ["GMAIL_APP_PASSWORD"]
TOKEN_PRODUCT_ID = os.environ["TOKEN_PRODUCT_ID"]
TUTOR_ROLE_ID = int(os.environ["TUTOR_ROLE_ID"])

TZ = ZoneInfo("America/Toronto")
DB_PATH = os.environ.get("DB_PATH", "slots.db")

# Default daily slots, seeded into the DB on first run. After that, slots
# live in the slot_definitions table and admins manage them with
# /addslot, /removeslot, /listslots.
DEFAULT_SLOTS = {
    "9-12": (9, 12, "9am - 12pm"),
    "12-3": (12, 15, "12pm - 3pm"),
    "3-6": (15, 18, "3pm - 6pm"),
}

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
            token_value INTEGER DEFAULT 1,
            used INTEGER DEFAULT 0,
            redeemed_by TEXT,
            redeemed_at TEXT,
            created_at TEXT
        )
    """)
    # Migration: add token_value to any pre-existing codes table that
    # predates this column.
    existing_cols = [row[1] for row in conn.execute("PRAGMA table_info(codes)").fetchall()]
    if "token_value" not in existing_cols:
        conn.execute("ALTER TABLE codes ADD COLUMN token_value INTEGER DEFAULT 1")
        conn.commit()
    conn.execute("""
        CREATE TABLE IF NOT EXISTS balances (
            discord_user_id TEXT PRIMARY KEY,
            tokens INTEGER DEFAULT 0
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS bookings (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            discord_user_id TEXT,
            booking_date TEXT,
            slot TEXT,
            start_ts TEXT,
            end_ts TEXT,
            channel_id TEXT,
            status TEXT DEFAULT 'booked'
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS student_channels (
            discord_user_id TEXT PRIMARY KEY,
            category_id TEXT,
            text_channel_id TEXT
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS slot_definitions (
            slot_key TEXT PRIMARY KEY,
            start_hour INTEGER,
            end_hour INTEGER,
            label TEXT
        )
    """)
    conn.commit()

    existing = conn.execute("SELECT COUNT(*) FROM slot_definitions").fetchone()[0]
    if existing == 0:
        for key, (start_h, end_h, label) in DEFAULT_SLOTS.items():
            conn.execute(
                "INSERT INTO slot_definitions (slot_key, start_hour, end_hour, label) VALUES (?, ?, ?, ?)",
                (key, start_h, end_h, label),
            )
        conn.commit()

    return conn

db_lock = threading.Lock()
conn = db_conn()

# ---------------------------------------------------------------------------
# Email
# ---------------------------------------------------------------------------

def send_codes_email(to_email, codes):
    # codes is a list of (code, value) tuples
    total_tokens = sum(value for _, value in codes)

    if len(codes) == 1:
        code, value = codes[0]
        body = (
            f"Your code is: {code} (worth {value} token{'s' if value != 1 else ''})\n\n"
            f"Join our Discord and run /redeem {code} to add {value} token(s) to your "
            f"balance, then /book to reserve a slot."
        )
    else:
        code_lines = "\n".join(f"  {c} — {v} token{'s' if v != 1 else ''}" for c, v in codes)
        body = (
            f"You purchased {total_tokens} token(s) total. Your codes are:\n\n{code_lines}\n\n"
            f"Join our Discord and run /redeem <code> for each one to add them to "
            f"your balance, then /book to reserve slots."
        )

    msg = MIMEText(body)
    msg["Subject"] = "Your voice channel tokens"
    msg["From"] = GMAIL_ADDRESS
    msg["To"] = to_email

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(GMAIL_ADDRESS, GMAIL_APP_PASSWORD)
        server.sendmail(GMAIL_ADDRESS, to_email, msg.as_string())

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

    # One code per matching line item, worth that item's quantity in
    # tokens. Only line items for the configured token product are
    # counted — other products in the store (if any) are ignored.
    all_line_items = order.get("line_items", [])
    line_items = [
        item for item in all_line_items
        if str(item.get("product_id")) == TOKEN_PRODUCT_ID
    ]

    if not line_items:
        print(f"[webhook] order {order_id} had no matching token product line items, skipping")
        return "ok", 200

    codes = []  # list of (code, value) tuples
    with db_lock:
        for item in line_items:
            value = item.get("quantity", 1)
            code = secrets.token_hex(4).upper()
            conn.execute(
                "INSERT INTO codes (code, order_id, customer_email, token_value, created_at) VALUES (?, ?, ?, ?, ?)",
                (code, order_id, email, value, datetime.datetime.utcnow().isoformat()),
            )
            codes.append((code, value))
        conn.commit()

    if email != "unknown":
        try:
            send_codes_email(email, codes)
        except Exception as e:
            print(f"[webhook] failed to email codes to {email}: {e}")

    print(f"[webhook] order {order_id} -> codes {codes} for {email}")

    return "ok", 200

def run_flask():
    port = int(os.environ.get("PORT", FLASK_PORT))
    app.run(host="0.0.0.0", port=port)

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
        slot_scheduler.start()

client = SlotBot()

def get_balance(discord_user_id: str) -> int:
    row = conn.execute(
        "SELECT tokens FROM balances WHERE discord_user_id = ?", (discord_user_id,)
    ).fetchone()
    return row[0] if row else 0

def get_slots() -> dict:
    """Returns {slot_key: (start_hour, end_hour, label)} from the DB."""
    rows = conn.execute("SELECT slot_key, start_hour, end_hour, label FROM slot_definitions").fetchall()
    return {key: (start_h, end_h, label) for key, start_h, end_h, label in rows}

async def ensure_student_channels(guild: discord.Guild, member: discord.Member):
    """Returns (category, text_channel), creating them if this student doesn't have them yet."""
    row = conn.execute(
        "SELECT category_id, text_channel_id FROM student_channels WHERE discord_user_id = ?",
        (str(member.id),),
    ).fetchone()

    if row:
        category = guild.get_channel(int(row[0])) if row[0] else None
        text_channel = guild.get_channel(int(row[1])) if row[1] else None
        if category is not None and text_channel is not None:
            return category, text_channel

    tutor_role = guild.get_role(TUTOR_ROLE_ID)
    overwrites = {
        guild.default_role: discord.PermissionOverwrite(view_channel=False, connect=False),
        member: discord.PermissionOverwrite(view_channel=True, connect=True, send_messages=True, read_message_history=True),
        guild.me: discord.PermissionOverwrite(view_channel=True, connect=True, send_messages=True, manage_channels=True),
    }
    if tutor_role is not None:
        overwrites[tutor_role] = discord.PermissionOverwrite(
            view_channel=True, connect=True, send_messages=True, read_message_history=True
        )

    category = await guild.create_category(name=member.display_name[:100], overwrites=overwrites)
    text_channel = await guild.create_text_channel(
        name="chat", overwrites=overwrites, category=category
    )

    with db_lock:
        conn.execute(
            "INSERT INTO student_channels (discord_user_id, category_id, text_channel_id) VALUES (?, ?, ?) "
            "ON CONFLICT(discord_user_id) DO UPDATE SET category_id = ?, text_channel_id = ?",
            (str(member.id), str(category.id), str(text_channel.id), str(category.id), str(text_channel.id)),
        )
        conn.commit()

    await text_channel.send(
        f"{member.mention} welcome! This is your private space — only you and tutors can see it. "
        f"Use this channel for commands and questions."
    )

    return category, text_channel

def is_admin(interaction: discord.Interaction) -> bool:
    return interaction.user.guild_permissions.administrator

def get_student_text_channel(discord_user_id: str):
    row = conn.execute(
        "SELECT text_channel_id FROM student_channels WHERE discord_user_id = ?",
        (discord_user_id,),
    ).fetchone()
    return row[0] if row else None

async def reply(interaction: discord.Interaction, content: str, ephemeral: bool = True):
    """Sends the bot's reply to the command user, and also posts a copy into
    their private student channel so tutors can see every response the bot
    gives, not just the command that triggered it."""
    await interaction.response.send_message(content, ephemeral=ephemeral)

    text_channel_id = get_student_text_channel(str(interaction.user.id))
    if text_channel_id:
        channel = client.get_channel(int(text_channel_id))
        if channel is not None:
            try:
                await channel.send(f"🤖 Bot replied to {interaction.user.mention}:\n{content}")
            except discord.HTTPException:
                pass

async def admin_reply(interaction: discord.Interaction, content: str, ephemeral: bool = True):
    """Admin-only command reply — stays private to the admin, never logged
    into any student's channel."""
    await interaction.response.send_message(content, ephemeral=ephemeral)

ADMIN_COMMANDS = {"addslot", "removeslot", "createcode", "addbalance", "removebalance", "exportbalances"}

@client.tree.interaction_check
async def log_commands_to_student_channel(interaction: discord.Interaction) -> bool:
    if interaction.type != discord.InteractionType.application_command:
        return True

    cmd_name = interaction.command.name if interaction.command else "unknown"
    if cmd_name in ADMIN_COMMANDS:
        return True

    text_channel_id = get_student_text_channel(str(interaction.user.id))
    if text_channel_id:
        channel = client.get_channel(int(text_channel_id))
        if channel is not None:
            params = " ".join(f"{k}:{v}" for k, v in interaction.namespace.__dict__.items())
            try:
                await channel.send(f"📋 {interaction.user.mention} ran `/{cmd_name} {params}`")
            except discord.HTTPException:
                pass

    return True

@client.event
async def on_member_join(member: discord.Member):
    if member.bot:
        return
    await ensure_student_channels(member.guild, member)

@client.tree.command(
    name="redeem",
    description="Redeem a purchased code to add tokens to your balance",
    guild=discord.Object(id=GUILD_ID),
)
@app_commands.describe(code="The code from your order confirmation email")
async def redeem(interaction: discord.Interaction, code: str):
    code = code.strip().upper()
    user_id = str(interaction.user.id)

    with db_lock:
        row = conn.execute("SELECT used, token_value FROM codes WHERE code = ?", (code,)).fetchone()

        if row is None:
            await reply(interaction, "That code isn't valid.", ephemeral=True)
            return
        if row[0] == 1:
            await reply(interaction, "That code has already been used.", ephemeral=True)
            return

        value = row[1] or 1

        conn.execute(
            "UPDATE codes SET used = 1, redeemed_by = ?, redeemed_at = ? WHERE code = ?",
            (user_id, datetime.datetime.utcnow().isoformat(), code),
        )
        conn.execute(
            "INSERT INTO balances (discord_user_id, tokens) VALUES (?, ?) "
            "ON CONFLICT(discord_user_id) DO UPDATE SET tokens = tokens + ?",
            (user_id, value, value),
        )
        conn.commit()
        new_balance = get_balance(user_id)

    await reply(interaction, 
        f"{value} token(s) added! You now have {new_balance} token(s). Use /book to reserve a slot.",
        ephemeral=True,
    )

async def slot_autocomplete(interaction: discord.Interaction, current: str):
    slots = get_slots()
    choices = [
        app_commands.Choice(name=label, value=key)
        for key, (_, _, label) in slots.items()
        if current.lower() in label.lower() or current.lower() in key.lower()
    ]
    return choices[:25]

@client.tree.command(
    name="book",
    description="Book a slot using one of your tokens",
    guild=discord.Object(id=GUILD_ID),
)
@app_commands.describe(
    date="Date in YYYY-MM-DD format (Eastern time)",
    slot="Which time slot (use /listslots to see options)",
)
@app_commands.autocomplete(slot=slot_autocomplete)
async def book(interaction: discord.Interaction, date: str, slot: str):
    user_id = str(interaction.user.id)
    slot_key = slot

    slots = get_slots()
    if slot_key not in slots:
        await reply(interaction, 
            "That slot doesn't exist. Use /listslots to see current options.", ephemeral=True
        )
        return

    try:
        booking_date = datetime.datetime.strptime(date, "%Y-%m-%d").date()
    except ValueError:
        await reply(interaction, 
            "Invalid date format. Use YYYY-MM-DD, e.g. 2026-07-15.", ephemeral=True
        )
        return

    start_hour, end_hour, label = slots[slot_key]
    start_dt = datetime.datetime.combine(booking_date, datetime.time(hour=start_hour), tzinfo=TZ)
    end_dt = datetime.datetime.combine(booking_date, datetime.time(hour=end_hour), tzinfo=TZ)
    now = datetime.datetime.now(TZ)

    if end_dt <= now:
        await reply(interaction, 
            "That slot has already ended. Pick a future date/slot.", ephemeral=True
        )
        return

    with db_lock:
        balance = get_balance(user_id)
        if balance < 1:
            await reply(interaction, 
                "You don't have any tokens. Redeem a code first with /redeem.", ephemeral=True
            )
            return

        # Prevent the same user double-booking the identical slot.
        dupe = conn.execute(
            "SELECT id FROM bookings WHERE discord_user_id = ? AND booking_date = ? AND slot = ? AND status != 'cancelled'",
            (user_id, date, slot_key),
        ).fetchone()
        if dupe:
            await reply(interaction, 
                "You've already booked that exact slot.", ephemeral=True
            )
            return

        conn.execute(
            "UPDATE balances SET tokens = tokens - 1 WHERE discord_user_id = ?", (user_id,)
        )
        conn.execute(
            "INSERT INTO bookings (discord_user_id, booking_date, slot, start_ts, end_ts, status) "
            "VALUES (?, ?, ?, ?, ?, 'booked')",
            (user_id, date, slot_key, start_dt.isoformat(), end_dt.isoformat()),
        )
        conn.commit()

    joined_late = start_dt <= now
    timing_note = (
        "The slot has already started, so your channel will appear within a minute "
        f"and still end at the scheduled time ({end_dt.strftime('%-I:%M%p')} Eastern) — "
        "you won't get extra time for joining late."
        if joined_late else
        "Your private voice channel will appear automatically when the slot starts, "
        "and close automatically when it ends."
    )

    await reply(interaction, 
        f"Booked! {date} {label} (Eastern). {timing_note}",
        ephemeral=True,
    )

@client.tree.command(
    name="listslots",
    description="Show the current bookable time slots",
    guild=discord.Object(id=GUILD_ID),
)
async def listslots(interaction: discord.Interaction):
    slots = get_slots()
    if not slots:
        await reply(interaction, "No slots are configured.", ephemeral=True)
        return
    text = "\n".join(f"  {key} — {label}" for key, (_, _, label) in sorted(slots.items()))
    await reply(interaction, f"Current slots:\n{text}", ephemeral=True)

@client.tree.command(
    name="addslot",
    description="[Admin] Add or update a bookable time slot",
    guild=discord.Object(id=GUILD_ID),
)
@app_commands.describe(
    key="Short identifier, e.g. '9-12'",
    start_hour="Start hour, 24h format, e.g. 9",
    end_hour="End hour, 24h format, e.g. 12",
    label="Display label, e.g. '9am - 12pm'",
)
async def addslot(interaction: discord.Interaction, key: str, start_hour: int, end_hour: int, label: str):
    if not is_admin(interaction):
        await admin_reply(interaction, "Admins only.", ephemeral=True)
        return
    if not (0 <= start_hour < 24 and 0 < end_hour <= 24 and start_hour < end_hour):
        await admin_reply(interaction, "Hours must be 0-24 and start before end.", ephemeral=True)
        return

    with db_lock:
        conn.execute(
            "INSERT INTO slot_definitions (slot_key, start_hour, end_hour, label) VALUES (?, ?, ?, ?) "
            "ON CONFLICT(slot_key) DO UPDATE SET start_hour = ?, end_hour = ?, label = ?",
            (key, start_hour, end_hour, label, start_hour, end_hour, label),
        )
        conn.commit()

    await admin_reply(interaction, f"Slot '{key}' set to {label}.", ephemeral=True)

@client.tree.command(
    name="removeslot",
    description="[Admin] Remove a bookable time slot",
    guild=discord.Object(id=GUILD_ID),
)
@app_commands.describe(key="The slot identifier to remove, e.g. '9-12'")
async def removeslot(interaction: discord.Interaction, key: str):
    if not is_admin(interaction):
        await admin_reply(interaction, "Admins only.", ephemeral=True)
        return

    with db_lock:
        cur = conn.execute("DELETE FROM slot_definitions WHERE slot_key = ?", (key,))
        conn.commit()

    if cur.rowcount == 0:
        await admin_reply(interaction, f"No slot found with key '{key}'.", ephemeral=True)
    else:
        await admin_reply(interaction, f"Slot '{key}' removed.", ephemeral=True)

@client.tree.command(
    name="createcode",
    description="[Admin] Manually generate a redemption code worth N tokens",
    guild=discord.Object(id=GUILD_ID),
)
@app_commands.describe(value="How many tokens this code is worth", note="Optional note, e.g. customer email or reason")
async def createcode(interaction: discord.Interaction, value: int, note: str = ""):
    if not is_admin(interaction):
        await admin_reply(interaction, "Admins only.", ephemeral=True)
        return
    if value <= 0:
        await admin_reply(interaction, "Value must be positive.", ephemeral=True)
        return

    code = secrets.token_hex(4).upper()
    with db_lock:
        conn.execute(
            "INSERT INTO codes (code, order_id, customer_email, token_value, created_at) VALUES (?, ?, ?, ?, ?)",
            (code, "manual", note or "manual", value, datetime.datetime.utcnow().isoformat()),
        )
        conn.commit()

    await admin_reply(interaction, 
        f"Created code `{code}` worth {value} token(s). Give this to whoever should redeem it with /redeem.",
        ephemeral=True,
    )

@client.tree.command(
    name="addbalance",
    description="[Admin] Increase a user's token balance",
    guild=discord.Object(id=GUILD_ID),
)
@app_commands.describe(user="The user to credit", amount="How many tokens to add")
async def addbalance(interaction: discord.Interaction, user: discord.Member, amount: int):
    if not is_admin(interaction):
        await admin_reply(interaction, "Admins only.", ephemeral=True)
        return
    if amount <= 0:
        await admin_reply(interaction, "Amount must be positive.", ephemeral=True)
        return

    user_id = str(user.id)
    with db_lock:
        conn.execute(
            "INSERT INTO balances (discord_user_id, tokens) VALUES (?, ?) "
            "ON CONFLICT(discord_user_id) DO UPDATE SET tokens = tokens + ?",
            (user_id, amount, amount),
        )
        conn.commit()
        new_balance = get_balance(user_id)

    await admin_reply(interaction, 
        f"Added {amount} token(s) to {user.mention}. New balance: {new_balance}.", ephemeral=True
    )

@client.tree.command(
    name="removebalance",
    description="[Admin] Decrease a user's token balance",
    guild=discord.Object(id=GUILD_ID),
)
@app_commands.describe(user="The user to debit", amount="How many tokens to remove")
async def removebalance(interaction: discord.Interaction, user: discord.Member, amount: int):
    if not is_admin(interaction):
        await admin_reply(interaction, "Admins only.", ephemeral=True)
        return
    if amount <= 0:
        await admin_reply(interaction, "Amount must be positive.", ephemeral=True)
        return

    user_id = str(user.id)
    with db_lock:
        current = get_balance(user_id)
        new_amount = max(0, current - amount)
        conn.execute(
            "INSERT INTO balances (discord_user_id, tokens) VALUES (?, ?) "
            "ON CONFLICT(discord_user_id) DO UPDATE SET tokens = ?",
            (user_id, new_amount, new_amount),
        )
        conn.commit()

    await admin_reply(interaction, 
        f"Removed {amount} token(s) from {user.mention}. New balance: {new_amount}.", ephemeral=True
    )

@client.tree.command(
    name="exportbalances",
    description="[Admin] Download all token balances as an Excel file",
    guild=discord.Object(id=GUILD_ID),
)
async def exportbalances(interaction: discord.Interaction):
    if not is_admin(interaction):
        await reply(interaction, "Admins only.", ephemeral=True)
        return

    await interaction.response.defer(ephemeral=True)

    guild = interaction.guild
    rows = conn.execute(
        "SELECT discord_user_id, tokens FROM balances ORDER BY tokens DESC"
    ).fetchall()

    wb = Workbook()
    ws = wb.active
    ws.title = "Balances"
    ws.append(["Discord User ID", "Username", "Token Balance"])

    for user_id, tokens in rows:
        member = guild.get_member(int(user_id)) if guild else None
        username = member.display_name if member else "(not in server)"
        ws.append([user_id, username, tokens])

    file_path = "/tmp/balances.xlsx"
    wb.save(file_path)

    await interaction.followup.send(
        "Current token balances:",
        file=discord.File(file_path, filename="balances.xlsx"),
        ephemeral=True,
    )

@client.tree.command(
    name="mybookings",
    description="Show your upcoming bookings and token balance",
    guild=discord.Object(id=GUILD_ID),
)
async def mybookings(interaction: discord.Interaction):
    user_id = str(interaction.user.id)
    balance = get_balance(user_id)

    rows = conn.execute(
        "SELECT booking_date, slot, status FROM bookings WHERE discord_user_id = ? "
        "AND status != 'cancelled' ORDER BY booking_date, slot",
        (user_id,),
    ).fetchall()

    if not rows:
        booking_text = "No upcoming bookings."
    else:
        booking_text = "\n".join(f"  {d} — {s} ({st})" for d, s, st in rows)

    await reply(interaction, 
        f"Token balance: {balance}\n\nBookings:\n{booking_text}", ephemeral=True
    )

@client.tree.command(
    name="checkbalance",
    description="Check your current token balance",
    guild=discord.Object(id=GUILD_ID),
)
async def checkbalance(interaction: discord.Interaction):
    balance = get_balance(str(interaction.user.id))
    await reply(interaction, 
        f"Your token balance: {balance}", ephemeral=True
    )

# ---------------------------------------------------------------------------
# Scheduler: create channels when slots start, delete when they end
# ---------------------------------------------------------------------------

@tasks.loop(seconds=60)
async def slot_scheduler():
    guild = client.get_guild(GUILD_ID)
    if guild is None:
        return

    now = datetime.datetime.now(TZ)

    # Start slots that have just begun and don't have channels yet.
    to_start = conn.execute(
        "SELECT id, discord_user_id, start_ts, end_ts FROM bookings "
        "WHERE status = 'booked' AND channel_id IS NULL"
    ).fetchall()

    for booking_id, user_id, start_ts, end_ts in to_start:
        start_dt = datetime.datetime.fromisoformat(start_ts)
        if start_dt <= now:
            member = guild.get_member(int(user_id))
            if member is None:
                continue

            category, text_channel = await ensure_student_channels(guild, member)

            tutor_role = guild.get_role(TUTOR_ROLE_ID)
            voice_overwrites = {
                guild.default_role: discord.PermissionOverwrite(view_channel=False, connect=False),
                member: discord.PermissionOverwrite(view_channel=True, connect=True, speak=True),
                guild.me: discord.PermissionOverwrite(view_channel=True, connect=True, manage_channels=True),
            }
            if tutor_role is not None:
                voice_overwrites[tutor_role] = discord.PermissionOverwrite(view_channel=True, connect=True, speak=True)

            voice_channel = await guild.create_voice_channel(
                name="voice",
                overwrites=voice_overwrites,
                category=category,
            )

            with db_lock:
                conn.execute(
                    "UPDATE bookings SET channel_id = ? WHERE id = ?",
                    (str(voice_channel.id), booking_id),
                )
                conn.commit()

            await text_channel.send(
                f"{member.mention} your slot has started — join your voice room: {voice_channel.mention}"
            )

            try:
                await member.send(f"Your slot has started! Join your voice room: {voice_channel.mention}")
            except discord.Forbidden:
                pass  # user has DMs disabled

    # End slots that are over: delete only the voice channel, mark completed.
    # The student's text channel and category are permanent and untouched.
    to_end = conn.execute(
        "SELECT id, channel_id, end_ts FROM bookings "
        "WHERE status = 'booked' AND channel_id IS NOT NULL"
    ).fetchall()

    for booking_id, channel_id, end_ts in to_end:
        end_dt = datetime.datetime.fromisoformat(end_ts)
        if end_dt <= now:
            voice_channel = guild.get_channel(int(channel_id))
            if voice_channel is not None:
                for member in list(voice_channel.members):
                    try:
                        await member.move_to(None)
                    except discord.HTTPException:
                        pass
                try:
                    await voice_channel.delete(reason="Slot time expired")
                except discord.HTTPException:
                    pass

            with db_lock:
                conn.execute(
                    "UPDATE bookings SET status = 'completed' WHERE id = ?", (booking_id,)
                )
                conn.commit()

# ---------------------------------------------------------------------------
# Entrypoint: run Flask in a background thread, Discord bot in the main loop
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    flask_thread = threading.Thread(target=run_flask, daemon=True)
    flask_thread.start()
    client.run(DISCORD_BOT_TOKEN)