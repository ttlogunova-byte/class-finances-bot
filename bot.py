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
        f"💸 /contribution — взнос\n"
        f"🗑 /delete_expense — удалить последний расход"
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
        await update.message.reply_text("Расходов нет")
        return
    msg = "📋 Последние расходы:\n\n"
    for i, e in enumerate(data["expenses"][-10:], 1):
        msg += f"{i}. {e['date']} | {e['category']} | {e['amount']:,} ₽ | {e['description']} ({e['who']})\n"
    await update.message.reply_text(msg)

async def delete_expense(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_committee(update.effective_user.id):
        await update.message.reply_text("❌ Нет прав")
        return
    
    data = load_data()
    if not data["expenses"]:
        await update.message.reply_text("❌ Расходов нет")
        return
    
    deleted = data["expenses"].pop()
    save_data(data)
    await update.message.reply_text(
        f"✅ Удалён последний расход:\n"
        f"{deleted['date']} | {deleted['category']} | {deleted['amount']:,} ₽ | {deleted['description']}"
    )

async def expense_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_committee(update.effective_user.id):
        await update.message.reply_text("❌ Нет прав")
        return ConversationHandler.END
    await update.message.reply_text("Сумма расхода?")
    return AMOUNT

async def amount_expense(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        context.user_data["amount"] = float(update.message.text)
    except:
        await update.message.reply_text("Введи число")
        return AMOUNT
    buttons = [[c] for c in CATEGORIES]
    await update.message.reply_text("Категория?", reply_markup=ReplyKeyboardMarkup(buttons, one_time_keyboard=True))
    return CATEGORY

async def category_expense(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["category"] = update.message.text
    await update.message.reply_text("Описание?")
    return DESCRIPTION

async def description_expense(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["description"] = update.message.text
    await update.message.reply_text("От кого?")
    return WHO

async def who_expense(update: Update, context: ContextTypes.DEFAULT_TYPE):
    data = load_data()
    expense = {
        "date": datetime.now().strftime("%d.%m.%Y"),
        "amount": context.user_data["amount"],
        "category": context.user_data["category"],
        "description": context.user_data["description"],
        "who": update.message.text
    }
    data["expenses"].append(expense)
    save_data(data)
    await update.message.reply_text(f"✅ Расход {expense['amount']:,} ₽ добавлен")
    context.user_data.clear()
    return ConversationHandler.END

async def contribution_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_committee(update.effective_user.id):
        await update.message.reply_text("❌ Нет прав")
        return ConversationHandler.END
    await update.message.reply_text("Сумма взноса?")
    return AMOUNT

async def amount_contribution(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        context.user_data["amount"] = float(update.message.text)
    except:
        await update.message.reply_text("Введи число")
        return AMOUNT
    await update.message.reply_text("От кого?")
    return WHO

async def who_contribution(update: Update, context: ContextTypes.DEFAULT_TYPE):
    data = load_data()
    contrib = {"date": datetime.now().strftime("%d.%m.%Y"), "amount": context.user_data["amount"], "who": update.message.text}
    data["contributions"].append(contrib)
    save_data(data)
    await update.message.reply_text(f"✅ Взнос {contrib['amount']:,} ₽ от {contrib['who']} принят")
    context.user_data.clear()
    return ConversationHandler.END

async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.clear()
    await update.message.reply_text("❌ Отменено")
    return ConversationHandler.END

def main():
    if not BOT_TOKEN:
        raise RuntimeError("Нет BOT_TOKEN!")
    
    app = Application.builder().token(BOT_TOKEN).build()
    
    expense_conv = ConversationHandler(
        entry_points=[CommandHandler("expense", expense_start)],
        states={
            AMOUNT: [MessageHandler(filters.TEXT & ~filters.COMMAND, amount_expense)],
            CATEGORY: [MessageHandler(filters.TEXT & ~filters.COMMAND, category_expense)],
            DESCRIPTION: [MessageHandler(filters.TEXT & ~filters.COMMAND, description_expense)],
            WHO: [MessageHandler(filters.TEXT & ~filters.COMMAND, who_expense)],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
    )
    
    contribution_conv = ConversationHandler(
        entry_points=[CommandHandler("contribution", contribution_start)],
        states={
            AMOUNT: [MessageHandler(filters.TEXT & ~filters.COMMAND, amount_contribution)],
            WHO: [MessageHandler(filters.TEXT & ~filters.COMMAND, who_contribution)],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
    )
    
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("report", report))
    app.add_handler(CommandHandler("history", history))
    app.add_handler(CommandHandler("delete_expense", delete_expense))
    app.add_handler(expense_conv)
    app.add_handler(contribution_conv)
    
    logger.info("Бот запущен!")
    app.run_polling()

if __name__ == "__main__":
    main()
