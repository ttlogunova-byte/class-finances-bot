"""
Бот учёта финансов класса 1-К
Простая версия — данные хранятся в файле, без Google API.
"""

import os
import json
os.makedirs("/data", exist_ok=True)
import logging
from datetime import datetime
from telegram import Update, ReplyKeyboardMarkup, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    Application, CommandHandler, MessageHandler, filters,
    ConversationHandler, ContextTypes, CallbackQueryHandler
)

# ============ НАСТРОЙКИ ============
BOT_TOKEN = os.environ.get("BOT_TOKEN")
BASE_FUND = 41000
COMMITTEE_IDS = []  # Заполним позже Telegram ID комитета
DATA_FILE = "/data/data.json"

# Категории расходов
CATEGORIES = ["📚 Материалы", "🍕 Питание", "🎉 Мероприятие", "🏠 Организационные", "🚗 Доставка"]

# Состояния диалога
AMOUNT, CATEGORY, DESCRIPTION, WHO = range(4)

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


# ============ РАБОТА С ДАННЫМИ ============
def load_data():
    if os.path.exists(DATA_FILE):
        with open(DATA_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {"expenses": [], "contributions": []}


def save_data(data):
    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def is_committee(user_id):
    if not COMMITTEE_IDS:
        return True
    return user_id in COMMITTEE_IDS


# ============ КОМАНДЫ ============
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    await update.message.reply_text(
        f"👋 Привет! Я бот учёта финансов класса 1-К.\n"
        f"Твой ID: {user_id}\n\n"
        f"📊 /report — посмотреть фонд, расходы и остаток\n"
        f"📋 /history — последние расходы\n\n"
        f"Комитету:\n"
        f"💰 /expense — добавить расход\n"
        f"💸 /contribution — добавить взнос"
    )


async def report(update: Update, context: ContextTypes.DEFAULT_TYPE):
    data = load_data()
    
    total_expenses = sum(e["amount"] for e in data["expenses"])
    total_contributions = sum(c["amount"] for c in data["contributions"])
    
    current_fund = BASE_FUND + total_contributions - total_expenses
    
    # Расходы по категориям
    expenses_by_category = {}
    for expense in data["expenses"]:
        cat = expense.get("category", "Неизвестно")
        expenses_by_category[cat] = expenses_by_category.get(cat, 0) + expense["amount"]
    
    message = f"📊 **ФИНАНСОВЫЙ ОТЧЁТ**\n\n"
    message += f"💰 Начальный фонд: {BASE_FUND:,} ₽\n"
    message += f"➕ Добавлено взносов: {total_contributions:,} ₽\n"
    message += f"➖ Потрачено: {total_expenses:,} ₽\n"
    message += f"📈 **Остаток: {current_fund:,} ₽**\n\n"
    
    if expenses_by_category:
        message += "**По категориям:**\n"
        for cat, amount in sorted(expenses_by_category.items()):
            message += f"{cat}: {amount:,} ₽\n"
    
    await update.message.reply_text(message)


async def history(update: Update, context: ContextTypes.DEFAULT_TYPE):
    data = load_data()
    
    if not data["expenses"]:
        await update.message.reply_text("📋 Расходов пока нет.")
        return
    
    message = "📋 **ПОСЛЕДНИЕ РАСХОДЫ:**\n\n"
    for expense in data["expenses"][-10:]:
        message += (
            f"{expense['date']} | {expense['category']} | "
            f"{expense['amount']:,} ₽ | {expense['description']} "
            f"(добавил: {expense['who']})\n"
        )
    
    await update.message.reply_text(message)


# ============ ДОБАВЛЕНИЕ РАСХОДОВ ============
async def expense_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_committee(update.effective_user.id):
        await update.message.reply_text("❌ У тебя нет прав для добавления расходов.")
        return ConversationHandler.END
    
    await update.message.reply_text("💰 Сколько денег потрачено? (число)")
    return AMOUNT


async def amount_received(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        amount = float(update.message.text)
        context.user_data["amount"] = amount
    except ValueError:
        await update.message.reply_text("❌ Введи число")
        return AMOUNT
    
    # Кнопки категорий
    buttons = [[cat] for cat in CATEGORIES]
    await update.message.reply_text(
        "📂 Выбери категорию:",
        reply_markup=ReplyKeyboardMarkup(buttons, one_time_keyboard=True)
    )
    return CATEGORY


async def category_received(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["category"] = update.message.text
    await update.message.reply_text("📝 Описание (что купили)?")
    return DESCRIPTION


async def description_received(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["description"] = update.message.text
    await update.message.reply_text("👤 От кого (имя)?")
    return WHO


async def who_received(update: Update, context: ContextTypes.DEFAULT_TYPE):
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
    
    await update.message.reply_text(
        f"✅ Расход {expense['amount']:,} ₽ добавлен!\n"
        f"Категория: {expense['category']}"
    )
    
    context.user_data.clear()
    return ConversationHandler.END


# ============ ДОБАВЛЕНИЕ ВЗНОСОВ ============
async def contribution_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_committee(update.effective_user.id):
        await update.message.reply_text("❌ У тебя нет прав для добавления взносов.")
        return ConversationHandler.END
    
    await update.message.reply_text("💸 Сумма взноса? (число)")
    return AMOUNT


async def contribution_amount_received(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        amount = float(update.message.text)
        context.user_data["amount"] = amount
    except ValueError:
        await update.message.reply_text("❌ Введи число")
        return AMOUNT
    
    await update.message.reply_text("👤 От кого (имя)?")
    return WHO


async def contribution_who_received(update: Update, context: ContextTypes.DEFAULT_TYPE):
    data = load_data()
    
    contribution = {
        "date": datetime.now().strftime("%d.%m.%Y"),
        "amount": context.user_data["amount"],
        "who": update.message.text
    }
    
    data["contributions"].append(contribution)
    save_data(data)
    
    await update.message.reply_text(
        f"✅ Взнос {contribution['amount']:,} ₽ от {contribution['who']} принят!"
    )
    
    context.user_data.clear()
    return ConversationHandler.END


async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.clear()
    await update.message.reply_text("❌ Отменено.")
    return ConversationHandler.END


# ============ ЗАПУСК ============
def main():
    if not BOT_TOKEN:
        raise RuntimeError("Не задан BOT_TOKEN в переменных окружения!")

    app = Application.builder().token(BOT_TOKEN).build()

    # Диалог добавления расходов
    expense_conv = ConversationHandler(
        entry_points=[CommandHandler("expense", expense_start)],
        states={
            AMOUNT: [MessageHandler(filters.TEXT & ~filters.COMMAND, amount_received)],
            CATEGORY: [MessageHandler(filters.TEXT & ~filters.COMMAND, category_received)],
            DESCRIPTION: [MessageHandler(filters.TEXT & ~filters.COMMAND, description_received)],
            WHO: [MessageHandler(filters.TEXT & ~filters.COMMAND, who_received)],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
    )

    # Диалог добавления взносов
    contribution_conv = ConversationHandler(
        entry_points=[CommandHandler("contribution", contribution_start)],
        states={
            AMOUNT: [MessageHandler(filters.TEXT & ~filters.COMMAND, contribution_amount_received)],
            WHO: [MessageHandler(filters.TEXT & ~filters.COMMAND, contribution_who_received)],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
    )

    # Регистрируем всё
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("report", report))
    app.add_handler(CommandHandler("history", history))
    app.add_handler(expense_conv)
    app.add_handler(contribution_conv)

    logger.info("Бот запущен!")
    app.run_polling()


if __name__ == "__main__":
    main()
