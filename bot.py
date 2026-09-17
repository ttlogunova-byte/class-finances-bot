"""
Бот учёта финансов класса 1-К
"""

import os
import json
os.makedirs("/data", exist_ok=True)
import logging
from datetime import datetime
from telegram import Update, ReplyKeyboardMarkup
from telegram.ext import (
    Application, CommandHandler, MessageHandler, filters,
    ConversationHandler, ContextTypes
)

BOT_TOKEN = os.environ.get("BOT_TOKEN")
BASE_FUND = 41000
COMMITTEE_IDS = []
DATA_FILE = "/data/data.json"
CATEGORIES = ["📚 Материалы", "🍕 Питание", "🎉 Мероприятие", "🏠 Организационные", "🚗 Доставка"]

AMOUNT, CATEGORY, DESCRIPTION, WHO = range(4)

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

def load_data():
    if os.path.exists(DATA_FILE):
        with open(DATA_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {"expenses": [], "contributions": []}

def save_data(data):
    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def is_committee(user_id):
    return True if not COMMITTEE_IDS else user_id in COMMITTEE_IDS

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    await update.message.reply_text(
        f"👋 Привет! Я бот учёта финансов класса 1-К.\n"
        f"Твой ID: {user_id}\n\n"
        f"📊 /report — финансы\n"
        f"📋 /history — история\n\n"
        f"Комитет:\n"
        f"💰 /expense — расход\n"
        f"💸 /contribution — взнос"
    )

async def report(update: Update, context: ContextTypes.DEFAULT_TYPE):
    data = load_data()
    expenses = sum(e["amount"] for e in data["expenses"])
    contributions = sum(c["amount"] for c in data["contributions"])
    current = BASE_FUND + contributions - expenses
    
    msg = f"📊 ФИНАНСЫ\n\n💰 Фонд: {BASE_FUND:,} ₽\n➕ Взносы: {contributions:,} ₽\n➖ Расходы: {expenses:,} ₽\n📈 Остаток: {current:,} ₽"
    await update.message.reply_text(msg)

async def history(update: Update, context: ContextTypes.DEFAULT_TYPE):
    data = load_data()
    if not data["expenses"]:
