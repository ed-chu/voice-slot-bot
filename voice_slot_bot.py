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
import asyncio
import smtplib
from email.mime.text import MIMEText
from zoneinfo import ZoneInfo

import discord
from discord import app_commands
from discord.ext import tasks
from flask import Flask, request, abort
from dotenv import load_dotenv
from openpyxl import Workbook
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

load_dotenv()

DISCORD_BOT_TOKEN = os.environ["DISCORD_BOT_TOKEN"]
SHOPIFY_WEBHOOK_SECRET = os.environ["SHOPIFY_WEBHOOK_SECRET"]
GUILD_ID = int(os.environ["GUILD_ID"])
FLASK_PORT = int(os.environ.get("FLASK_PORT", 5000))
GMAIL_ADDRESS = os.environ["GMAIL_ADDRESS"]
GMAIL_APP_PASSWORD = os.environ["GMAIL_APP_PASSWORD"]
TOKEN_PRODUCT_ID = os.environ["TOKEN_PRODUCT_ID"]
TUTOR_ROLE_ID = int(os.environ["TUTOR_ROLE_ID"])
GOOGLE_CLIENT_ID = os.environ.get("GOOGLE_CLIENT_ID")
GOOGLE_CLIENT_SECRET = os.environ.get("GOOGLE_CLIENT_SECRET")
GOOGLE_REFRESH_TOKEN = os.environ.get("GOOGLE_REFRESH_TOKEN")
GOOGLE_TEMPLATE_SHEET_ID = os.environ.get("GOOGLE_TEMPLATE_SHEET_ID", "1LiKujXIT7FIkcYmOnrfsEySq96tzg84IoYwWceLFpRo")
GOOGLE_PARENT_FOLDER_ID = os.environ.get("GOOGLE_PARENT_FOLDER_ID")
GOOGLE_SHEETS_ENABLED = bool(GOOGLE_CLIENT_ID and GOOGLE_CLIENT_SECRET and GOOGLE_REFRESH_TOKEN and GOOGLE_PARENT_FOLDER_ID)

TZ = ZoneInfo("America/Toronto")

# TODO: replace with your real sign-in link once you have one.
SIGNIN_LINK = os.environ.get("SIGNIN_LINK", "https://example.com/sign-in-placeholder")
FEEDBACK_FORM_LINK = "https://docs.google.com/forms/d/e/1FAIpQLSdpe7V3wpKN_k_l7BzQlmSqLBsnckgBz6BgkOaPW1RMjvpjpg/viewform?usp=sharing&ouid=100164859931306672567"
DB_PATH = os.environ.get("DB_PATH", "slots.db")

SESSION_HOURS = 3  # every booking is a fixed 3-hour block
MIN_ADVANCE_HOURS = 24  # bookings must be made at least this far ahead

# ---------------------------------------------------------------------------
# Database setup
# ---------------------------------------------------------------------------

def db_conn():
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    # The Flask webhook thread and the Discord bot's asyncio thread both hit
    # this same connection. WAL mode lets reads and writes coexist better,
    # and busy_timeout makes SQLite wait/retry briefly on a lock instead of
    # raising "database is locked" immediately, which was causing some
    # slash commands to fail silently before they could reply to Discord.
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA busy_timeout=5000")
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
            status TEXT DEFAULT 'booked',
            checkout_reminder_sent INTEGER DEFAULT 0,
            signin_reminder_sent INTEGER DEFAULT 0
        )
    """)
    booking_cols = [row[1] for row in conn.execute("PRAGMA table_info(bookings)").fetchall()]
    if "checkout_reminder_sent" not in booking_cols:
        conn.execute("ALTER TABLE bookings ADD COLUMN checkout_reminder_sent INTEGER DEFAULT 0")
        conn.commit()
    if "signin_reminder_sent" not in booking_cols:
        conn.execute("ALTER TABLE bookings ADD COLUMN signin_reminder_sent INTEGER DEFAULT 0")
        conn.commit()
    conn.execute("""
        CREATE TABLE IF NOT EXISTS student_channels (
            discord_user_id TEXT PRIMARY KEY,
            category_id TEXT,
            text_channel_id TEXT
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS student_emails (
            discord_user_id TEXT PRIMARY KEY,
            email TEXT
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS student_sheets (
            discord_user_id TEXT PRIMARY KEY,
            folder_id TEXT,
            spreadsheet_id TEXT,
            spreadsheet_url TEXT
        )
    """)
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

    # DB has no record (e.g. it was reset) but a folder may already exist
    # from before — look for a category that's private to exactly this
    # member before creating a new (duplicate) one.
    existing_category = None
    for cat in guild.categories:
        overwrite = cat.overwrites_for(member)
        everyone_overwrite = cat.overwrites_for(guild.default_role)
        if overwrite.view_channel is True and everyone_overwrite.view_channel is False:
            existing_category = cat
            break

    if existing_category is not None:
        existing_text = discord.utils.get(existing_category.text_channels, name="chat")
        if existing_text is not None:
            with db_lock:
                conn.execute(
                    "INSERT INTO student_channels (discord_user_id, category_id, text_channel_id) VALUES (?, ?, ?) "
                    "ON CONFLICT(discord_user_id) DO UPDATE SET category_id = ?, text_channel_id = ?",
                    (str(member.id), str(existing_category.id), str(existing_text.id),
                     str(existing_category.id), str(existing_text.id)),
                )
                conn.commit()
            return existing_category, existing_text

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

    category_name = f"{member.display_name}-{str(member.id)[-4:]}"[:100]
    category = await guild.create_category(name=category_name, overwrites=overwrites)
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

    welcome_message = (
        f"👋 **Welcome, {member.mention}!**\n\n"
        f"This is your private space — only you and your tutors can see it. All your sessions, "
        f"bookings, and your study log will happen here.\n\n"
        f"**How it works:**\n"
        f"1️⃣ After you purchase tokens, you'll get code(s) by email.\n"
        f"2️⃣ Redeem a code here with `/redeem code:YOUR_CODE` — this adds tokens to your balance.\n"
        f"3️⃣ Book a {SESSION_HOURS}-hour session with `/book date:07-15 time:2pm` (any date/time, "
        f"at least {MIN_ADVANCE_HOURS} hours in advance).\n"
        f"4️⃣ You'll get a reminder 3 minutes before your session starts, and check-in instructions "
        f"right when it begins — a private voice room appears automatically.\n"
        f"5️⃣ About 5 minutes before your session ends, you'll get a check-out reminder with a quick "
        f"feedback form.\n\n"
        f"**Common commands:**\n"
        f"`/redeem code:XXXX` — add tokens to your balance\n"
        f"`/checkbalance` — see how many tokens you have\n"
        f"`/book date:MM-DD time:H:MMam/pm` — book a session\n"
        f"`/mybookings` — see your upcoming sessions and balance\n"
        f"`/setemail email:you@example.com` — link your Google account so your study log gets shared with you\n\n"
        f"If anything isn't working, just ask here — a tutor can see this channel and help directly."
    )

    await text_channel.send(welcome_message)

    return category, text_channel

# ---------------------------------------------------------------------------
# Google Sheets: per-student study log, created from a template on their
# first check-in. Runs the (blocking) Google API calls in a thread executor
# so it doesn't freeze the bot's event loop.
# ---------------------------------------------------------------------------

def get_student_email(discord_user_id: str):
    row = conn.execute(
        "SELECT email FROM student_emails WHERE discord_user_id = ?", (discord_user_id,)
    ).fetchone()
    if row and row[0]:
        return row[0]
    # Fall back to the email from their most recent redeemed Shopify code —
    # may belong to a parent/payer rather than the student, so /setemail
    # should be preferred when the student has their own Google account.
    row = conn.execute(
        "SELECT customer_email FROM codes WHERE redeemed_by = ? ORDER BY redeemed_at DESC LIMIT 1",
        (discord_user_id,),
    ).fetchone()
    return row[0] if row and row[0] and row[0] != "unknown" else None

def _get_drive_service():
    creds = Credentials(
        None,
        refresh_token=GOOGLE_REFRESH_TOKEN,
        client_id=GOOGLE_CLIENT_ID,
        client_secret=GOOGLE_CLIENT_SECRET,
        token_uri="https://oauth2.googleapis.com/token",
    )
    return build("drive", "v3", credentials=creds)

def _create_student_sheet_sync(student_name: str, student_email):
    """Blocking Google API work: create a subfolder, copy the template sheet
    into it, and share the subfolder with the student. Must be called via
    run_in_executor, never awaited directly."""
    drive = _get_drive_service()

    folder = drive.files().create(
        body={
            "name": student_name,
            "mimeType": "application/vnd.google-apps.folder",
            "parents": [GOOGLE_PARENT_FOLDER_ID],
        },
        fields="id",
    ).execute()
    folder_id = folder["id"]

    sheet = drive.files().copy(
        fileId=GOOGLE_TEMPLATE_SHEET_ID,
        body={"name": f"{student_name} - Study Log", "parents": [folder_id]},
        fields="id, webViewLink",
    ).execute()

    if student_email:
        drive.permissions().create(
            fileId=folder_id,
            body={"type": "user", "role": "writer", "emailAddress": student_email},
            sendNotificationEmail=False,
        ).execute()

    return folder_id, sheet["id"], sheet.get("webViewLink")

async def ensure_student_sheet(member: discord.Member):
    """Returns the study log URL, creating it on first call for this
    student. Returns None if Google Sheets isn't configured or creation
    fails — callers should treat that as 'skip this, not fatal'."""
    if not GOOGLE_SHEETS_ENABLED:
        return None

    user_id = str(member.id)
    row = conn.execute(
        "SELECT spreadsheet_url FROM student_sheets WHERE discord_user_id = ?", (user_id,)
    ).fetchone()
    if row:
        return row[0]

    email = get_student_email(user_id)
    loop = asyncio.get_event_loop()
    try:
        folder_id, sheet_id, url = await loop.run_in_executor(
            None, _create_student_sheet_sync, member.display_name, email
        )
    except Exception as e:
        print(f"[sheets] failed to create study log for {user_id}: {e}")
        return None

    with db_lock:
        conn.execute(
            "INSERT INTO student_sheets (discord_user_id, folder_id, spreadsheet_id, spreadsheet_url) "
            "VALUES (?, ?, ?, ?)",
            (user_id, folder_id, sheet_id, url),
        )
        conn.commit()

    return url

def is_admin(interaction: discord.Interaction) -> bool:
    return interaction.user.guild_permissions.administrator

def is_tutor_or_admin(interaction: discord.Interaction) -> bool:
    if is_admin(interaction):
        return True
    return any(role.id == TUTOR_ROLE_ID for role in interaction.user.roles)

def get_student_text_channel(discord_user_id: str):
    row = conn.execute(
        "SELECT text_channel_id FROM student_channels WHERE discord_user_id = ?",
        (discord_user_id,),
    ).fetchone()
    return row[0] if row else None

async def reply(interaction: discord.Interaction, content: str, ephemeral: bool = True):
    """Replies directly in the channel as a normal message, so tutors see it
    naturally (Discord also shows the /command that triggered it automatically
    for non-ephemeral responses — no separate logging needed). Only falls
    back to a private ephemeral reply if the command wasn't run inside the
    student's own private channel, to avoid leaking things like redemption
    codes into a channel other people can see."""
    text_channel_id = get_student_text_channel(str(interaction.user.id))
    in_own_channel = (
        text_channel_id is not None
        and interaction.channel is not None
        and str(interaction.channel.id) == text_channel_id
    )

    if in_own_channel:
        await interaction.response.send_message(content)
    else:
        await interaction.response.send_message(
            f"{content}\n\n(Tip: run commands in your private channel so tutors can see them.)",
            ephemeral=True,
        )

async def admin_reply(interaction: discord.Interaction, content: str, ephemeral: bool = True):
    """Admin-only command reply — stays private to the admin, never shown
    in any student's channel."""
    await interaction.response.send_message(content, ephemeral=ephemeral)

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

def parse_month_day(date_str: str):
    """Parses a date string with no year, e.g. '07-15', '7/15', 'July 15',
    'Jul 15'. Returns (month, day) or None if it doesn't match anything."""
    date_str = date_str.strip()
    formats = ["%m-%d", "%m/%d", "%B %d", "%b %d", "%d %B", "%d %b"]
    for fmt in formats:
        try:
            parsed = datetime.datetime.strptime(date_str, fmt)
            return parsed.month, parsed.day
        except ValueError:
            continue
    return None

def parse_time_str(time_str: str):
    """Parses a start time, e.g. '2pm', '2:30pm', '14:00', '14'. Returns
    (hour, minute) or None if it doesn't match anything."""
    time_str = time_str.strip().upper().replace(" ", "")
    formats = ["%I:%M%p", "%I%p", "%H:%M", "%H"]
    for fmt in formats:
        try:
            parsed = datetime.datetime.strptime(time_str, fmt)
            return parsed.hour, parsed.minute
        except ValueError:
            continue
    return None

@client.tree.command(
    name="book",
    description=f"Book a {SESSION_HOURS}-hour study session at any time, at least {MIN_ADVANCE_HOURS}h in advance",
    guild=discord.Object(id=GUILD_ID),
)
@app_commands.describe(
    date="Date without year, e.g. 07-15, 7/15, or 'July 15' — nearest upcoming occurrence is used",
    time="Start time, e.g. 2pm, 2:30pm, or 14:00",
)
async def book(interaction: discord.Interaction, date: str, time: str):
    # Defer immediately so Discord always gets an ack within its 3-second
    # window, regardless of how long the logic below takes or whether it
    # errors.
    await interaction.response.defer()

    try:
        user_id = str(interaction.user.id)

        text_channel_id = get_student_text_channel(user_id)
        in_own_channel = (
            text_channel_id is not None
            and interaction.channel is not None
            and str(interaction.channel.id) == text_channel_id
        )
        if not in_own_channel:
            await interaction.followup.send(
                "Please run /book inside your own private channel.", ephemeral=True
            )
            return

        parsed_date = parse_month_day(date)
        if parsed_date is None:
            await interaction.followup.send(
                "Couldn't understand that date. Try 07-15, 7/15, or 'July 15'.", ephemeral=True
            )
            return

        parsed_time = parse_time_str(time)
        if parsed_time is None:
            await interaction.followup.send(
                "Couldn't understand that time. Try 2pm, 2:30pm, or 14:00.", ephemeral=True
            )
            return

        month, day = parsed_date
        hour, minute = parsed_time
        now = datetime.datetime.now(TZ)
        today = now.date()

        try:
            candidate_date = datetime.date(today.year, month, day)
        except ValueError:
            await interaction.followup.send("That's not a valid date.", ephemeral=True)
            return

        start_dt = datetime.datetime.combine(candidate_date, datetime.time(hour=hour, minute=minute), tzinfo=TZ)
        if start_dt < now:
            try:
                candidate_date = datetime.date(today.year + 1, month, day)
            except ValueError:
                await interaction.followup.send(
                    "That date doesn't exist next year either (e.g. Feb 29).", ephemeral=True
                )
                return
            start_dt = datetime.datetime.combine(candidate_date, datetime.time(hour=hour, minute=minute), tzinfo=TZ)

        end_dt = start_dt + datetime.timedelta(hours=SESSION_HOURS)

        if start_dt - now < datetime.timedelta(hours=MIN_ADVANCE_HOURS):
            await interaction.followup.send(
                f"Sessions must be booked at least {MIN_ADVANCE_HOURS} hours in advance. "
                f"The earliest you can book is {(now + datetime.timedelta(hours=MIN_ADVANCE_HOURS)).strftime('%b %-d, %-I:%M%p')} Eastern.",
                ephemeral=True,
            )
            return

        with db_lock:
            balance = get_balance(user_id)
            if balance < 1:
                await interaction.followup.send(
                    "You don't have any tokens. Redeem a code first with /redeem.", ephemeral=True
                )
                return

            # Prevent this student from booking two overlapping sessions.
            overlap = conn.execute(
                "SELECT id FROM bookings WHERE discord_user_id = ? AND status = 'booked' "
                "AND start_ts < ? AND end_ts > ?",
                (user_id, end_dt.isoformat(), start_dt.isoformat()),
            ).fetchone()
            if overlap:
                await interaction.followup.send(
                    "You already have a session that overlaps this time.", ephemeral=True
                )
                return

            conn.execute("UPDATE balances SET tokens = tokens - 1 WHERE discord_user_id = ?", (user_id,))
            label = f"{start_dt.strftime('%-I:%M%p')} - {end_dt.strftime('%-I:%M%p')}"
            conn.execute(
                "INSERT INTO bookings (discord_user_id, booking_date, slot, start_ts, end_ts, status) "
                "VALUES (?, ?, ?, ?, ?, 'booked')",
                (user_id, start_dt.date().isoformat(), label, start_dt.isoformat(), end_dt.isoformat()),
            )
            conn.commit()

        await interaction.followup.send(
            f"✅ Booked! {start_dt.strftime('%A, %B %-d')} from {label} (Eastern). "
            f"Your voice channel will appear automatically when the session starts, "
            f"and close automatically when it ends."
        )
    except Exception as e:
        print(f"[book] unexpected error for user {interaction.user.id}: {e}")
        try:
            await interaction.followup.send(
                "Something went wrong processing that — please try again.", ephemeral=True
            )
        except discord.HTTPException:
            pass

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
        await admin_reply(interaction, "Admins only.", ephemeral=True)
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
    name="createroom",
    description="[Tutor] Recreate a student's private room if it was deleted",
    guild=discord.Object(id=GUILD_ID),
)
@app_commands.describe(user="The student to create/restore a room for")
async def createroom(interaction: discord.Interaction, user: discord.Member):
    if not is_tutor_or_admin(interaction):
        await admin_reply(interaction, "Tutors/admins only.", ephemeral=True)
        return

    category, text_channel = await ensure_student_channels(interaction.guild, user)
    await admin_reply(
        interaction,
        f"Room ready for {user.mention}: {text_channel.mention}",
        ephemeral=True,
    )

@client.tree.command(
    name="viewbookings",
    description="[Tutor] List a student's upcoming bookings with their IDs",
    guild=discord.Object(id=GUILD_ID),
)
@app_commands.describe(user="The student to look up")
async def viewbookings(interaction: discord.Interaction, user: discord.Member):
    if not is_tutor_or_admin(interaction):
        await admin_reply(interaction, "Tutors/admins only.", ephemeral=True)
        return

    rows = conn.execute(
        "SELECT id, booking_date, slot, status, start_ts FROM bookings "
        "WHERE discord_user_id = ? AND status != 'cancelled' ORDER BY start_ts",
        (str(user.id),),
    ).fetchall()

    if not rows:
        await admin_reply(interaction, f"{user.mention} has no bookings.", ephemeral=True)
        return

    now = datetime.datetime.now(TZ)
    lines = []
    for booking_id, booking_date, slot, status, start_ts in rows:
        start_dt = datetime.datetime.fromisoformat(start_ts)
        started_note = " (in progress/past)" if start_dt <= now else ""
        lines.append(f"ID {booking_id} — {booking_date} {slot} ({status}){started_note}")

    await admin_reply(
        interaction,
        f"Bookings for {user.mention}:\n" + "\n".join(lines),
        ephemeral=True,
    )

@client.tree.command(
    name="cancelbooking",
    description="[Tutor] Cancel a student's booking, only if it hasn't started yet",
    guild=discord.Object(id=GUILD_ID),
)
@app_commands.describe(booking_id="The booking ID from /viewbookings")
async def cancelbooking(interaction: discord.Interaction, booking_id: int):
    if not is_tutor_or_admin(interaction):
        await admin_reply(interaction, "Tutors/admins only.", ephemeral=True)
        return

    with db_lock:
        row = conn.execute(
            "SELECT discord_user_id, booking_date, slot, start_ts, status FROM bookings WHERE id = ?",
            (booking_id,),
        ).fetchone()

        if row is None:
            await admin_reply(interaction, "No booking found with that ID.", ephemeral=True)
            return

        user_id, booking_date, slot, start_ts, status = row

        if status != "booked":
            await admin_reply(interaction, f"That booking is already '{status}', nothing to cancel.", ephemeral=True)
            return

        start_dt = datetime.datetime.fromisoformat(start_ts)
        now = datetime.datetime.now(TZ)
        if start_dt <= now:
            await admin_reply(
                interaction,
                "That booking has already started, so it can't be cancelled this way.",
                ephemeral=True,
            )
            return

        conn.execute("UPDATE bookings SET status = 'cancelled' WHERE id = ?", (booking_id,))
        conn.execute(
            "INSERT INTO balances (discord_user_id, tokens) VALUES (?, 1) "
            "ON CONFLICT(discord_user_id) DO UPDATE SET tokens = tokens + 1",
            (user_id,),
        )
        conn.commit()

    await admin_reply(
        interaction,
        f"Cancelled booking {booking_id} ({booking_date} {slot}) and refunded 1 token to <@{user_id}>.",
        ephemeral=True,
    )

@client.tree.command(
    name="startroom",
    description="[Tutor] Start a room for a student right now for N minutes",
    guild=discord.Object(id=GUILD_ID),
)
@app_commands.describe(user="The student to start a room for", minutes="How many minutes the room should stay open")
async def startroom(interaction: discord.Interaction, user: discord.Member, minutes: int):
    if not is_tutor_or_admin(interaction):
        await admin_reply(interaction, "Tutors/admins only.", ephemeral=True)
        return
    if minutes <= 0:
        await admin_reply(interaction, "Minutes must be positive.", ephemeral=True)
        return

    await interaction.response.defer(ephemeral=True)

    now = datetime.datetime.now(TZ)
    end_dt = now + datetime.timedelta(minutes=minutes)

    with db_lock:
        cur = conn.execute(
            "INSERT INTO bookings (discord_user_id, booking_date, slot, start_ts, end_ts, status) "
            "VALUES (?, ?, ?, ?, ?, 'booked')",
            (str(user.id), now.date().isoformat(), "manual", now.isoformat(), end_dt.isoformat()),
        )
        conn.commit()
        booking_id = cur.lastrowid

    voice_channel = await create_voice_room_for_booking(interaction.guild, user, booking_id, end_dt)

    await interaction.followup.send(
        f"Started a {minutes}-minute room for {user.mention}: {voice_channel.mention}. "
        f"It will close automatically at {end_dt.strftime('%-I:%M%p')} Eastern.",
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

@client.tree.command(
    name="setemail",
    description="Set the Google account email your study log should be shared with",
    guild=discord.Object(id=GUILD_ID),
)
@app_commands.describe(email="The Gmail/Google account email you use")
async def setemail(interaction: discord.Interaction, email: str):
    if "@" not in email or "." not in email:
        await reply(interaction, "That doesn't look like a valid email address.", ephemeral=True)
        return

    user_id = str(interaction.user.id)
    with db_lock:
        conn.execute(
            "INSERT INTO student_emails (discord_user_id, email) VALUES (?, ?) "
            "ON CONFLICT(discord_user_id) DO UPDATE SET email = ?",
            (user_id, email, email),
        )
        conn.commit()

    await reply(
        interaction,
        f"Email set to {email}. If your study log hasn't been created yet, it'll be shared with this "
        f"address automatically at your next check-in. If it already exists, ask a tutor to share it "
        f"with you directly in Google Drive.",
        ephemeral=True,
    )

# ---------------------------------------------------------------------------
# Scheduler: create channels when slots start, delete when they end
# ---------------------------------------------------------------------------

async def create_voice_room_for_booking(guild: discord.Guild, member: discord.Member, booking_id: int, end_dt: datetime.datetime):
    """Creates the voice channel for a booking that's ready to start, updates
    the booking's channel_id, and sends check-in instructions. Shared by the
    scheduler (automatic slot starts) and /startroom (manual instant start)."""
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

    end_time_str = end_dt.strftime("%-I:%M%p")
    checkin_message = (
        f"✅ **Your session has started, {member.mention}!** Join your voice room: {voice_channel.mention}\n"
        f"If you haven't signed in yet, please do so here: {SIGNIN_LINK}\n"
        f"Please turn on your camera for the session.\n"
        f"Your session ends at {end_time_str} Eastern — you'll get a check-out reminder 5 minutes before it ends."
    )

    await text_channel.send(checkin_message)

    sheet_url = await ensure_student_sheet(member)
    if sheet_url:
        await text_channel.send(f"📊 Your study log: {sheet_url}")

    return voice_channel

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
            end_dt = datetime.datetime.fromisoformat(end_ts)
            await create_voice_room_for_booking(guild, member, booking_id, end_dt)

    # Sign-in reminders: 3 minutes before a booking starts, give a heads-up
    # in the student's permanent channel (the voice room doesn't exist yet).
    to_signin = conn.execute(
        "SELECT id, discord_user_id, start_ts FROM bookings "
        "WHERE status = 'booked' AND signin_reminder_sent = 0"
    ).fetchall()

    for booking_id, user_id, start_ts in to_signin:
        start_dt = datetime.datetime.fromisoformat(start_ts)
        minutes_until = (start_dt - now).total_seconds() / 60
        if 0 < minutes_until <= 3:
            member = guild.get_member(int(user_id))
            text_channel_id = get_student_text_channel(user_id)
            if member is not None and text_channel_id is not None:
                text_channel = guild.get_channel(int(text_channel_id))
                if text_channel is not None:
                    signin_message = (
                        f"🔔 **Heads up, {member.mention}** — your session starts in about 3 minutes, "
                        f"at {start_dt.strftime('%-I:%M%p')} Eastern.\n"
                        f"Please sign in here: {SIGNIN_LINK}\n"
                        f"Find a quiet place to work before your session begins."
                    )
                    try:
                        await text_channel.send(signin_message)
                    except discord.HTTPException:
                        pass

            with db_lock:
                conn.execute(
                    "UPDATE bookings SET signin_reminder_sent = 1 WHERE id = ?", (booking_id,)
                )
                conn.commit()

    # Check-out reminders: 5 minutes before a booking ends, send a one-time
    # heads-up so the student knows to wrap up.
    to_remind = conn.execute(
        "SELECT id, discord_user_id, end_ts FROM bookings "
        "WHERE status = 'booked' AND channel_id IS NOT NULL AND checkout_reminder_sent = 0"
    ).fetchall()

    for booking_id, user_id, end_ts in to_remind:
        end_dt = datetime.datetime.fromisoformat(end_ts)
        minutes_left = (end_dt - now).total_seconds() / 60
        if 0 < minutes_left <= 5:
            member = guild.get_member(int(user_id))
            text_channel_id = get_student_text_channel(user_id)
            if member is not None and text_channel_id is not None:
                text_channel = guild.get_channel(int(text_channel_id))
                if text_channel is not None:
                    checkout_message = (
                        f"⏰ **Check-out reminder, {member.mention}:** your session ends in about "
                        f"5 minutes, at {end_dt.strftime('%-I:%M%p')} Eastern. Please wrap up and "
                        f"leave the voice channel when you're done — it will close automatically.\n\n"
                        f"Before you go, please fill out this quick feedback form: {FEEDBACK_FORM_LINK}"
                    )
                    try:
                        await text_channel.send(checkout_message)
                    except discord.HTTPException:
                        pass

            with db_lock:
                conn.execute(
                    "UPDATE bookings SET checkout_reminder_sent = 1 WHERE id = ?", (booking_id,)
                )
                conn.commit()

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