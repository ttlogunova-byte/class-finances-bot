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
# Токен берётся из переменной окружения Railway (BOT_TOKEN)
BOT_TOKEN = os.environ.get("BOT_TOKEN")

# Начальный фонд (уже собранные деньги)
BASE_FUND = 41000

# ID тех, кто может ДОБАВЛЯТЬ (комитет). Заполним позже — пока пусто = все.
# Сюда впишем Telegram ID 4 человек комитета.
COMMITTEE_IDS = []  # например: [123456789, 987654321]

# Файл с данными (Amvera требует /data)
import os
os.makedirs("/data", exist_ok=True)
DATA_FILE = "/data/data.json"

# Категории расходов
CATEGORIES = ["📚 Материалы", "🍕 Питание", "🎉 Мероприятие", "🏠 Организационные", "🚗 Доставка"]

# ============ ЛОГИ ============
logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO
)
logger = logging.getLogger(__name__)

# Состояния диалога
CHOOSING_TYPE, AMOUNT, CATEGORY, DESCRIPTION, WHO = range(5)


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
    # Если список пуст — разрешаем всем (на время настройки)
    if not COMMITTEE_IDS:
        return True
    return user_id in COMMITTEE_IDS


# ============ КОМАНДЫ ============
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    text = (
        "👋 Привет! Я бот учёта финансов класса 1-К.\n\n"
        "📊 /отчет — посмотреть фонд, расходы и остаток\n"
        "📋 /история — последние расходы\n"
    )
    if is_committee(user_id):
        text += (
            "\n<b>Для комитета:</b>\n"
            "➕ /расход — добавить расход\n"
            "💸 /взнос — добавить взнос\n"
        )
    text += f"\n🆔 Твой ID: <code>{user_id}</code>"
    await update.message.reply_text(text, parse_mode="HTML")


async def report(update: Update, context: ContextTypes.DEFAULT_TYPE):
    data = load_data()
    total_spent = sum(e["amount"] for e in data["expenses"])
    total_contrib = sum(c["amount"] for c in data["contributions"])
    fund = BASE_FUND + total_contrib
    remaining = fund - total_spent

    # Разбивка по категориям
    by_cat = {}
    for e in data["expenses"]:
        by_cat[e["category"]] = by_cat.get(e["category"], 0) + e["amount"]

    text = (
        "💰 <b>ФИНАНСЫ КЛАССА 1-К</b>\n\n"
        f"💵 Фонд собрано: <b>{fund:,.0f} ₽</b>\n".replace(",", " ") +
        f"📉 Потрачено: <b>{total_spent:,.0f} ₽</b>\n".replace(",", " ") +
        f"✅ Остаток: <b>{remaining:,.0f} ₽</b>\n".replace(",", " ")
    )
    if by_cat:
        text += "\n<b>По категориям:</b>\n"
        for cat, amount in by_cat.items():
            text += f"  {cat}: {amount:,.0f} ₽\n".replace(",", " ")

    text += f"\n📝 Всего записей: {len(data['expenses'])}"
    await update.message.reply_text(text, parse_mode="HTML")


async def history(update: Update, context: ContextTypes.DEFAULT_TYPE):
    data = load_data()
    expenses = data["expenses"][-10:][::-1]  # последние 10, новые сверху
    if not expenses:
        await update.message.reply_text("📭 Пока нет расходов.")
        return

    text = "📋 <b>Последние расходы:</b>\n\n"
    for e in expenses:
        text += (
            f"📅 {e['date']} — <b>{e['amount']:,.0f} ₽</b>\n".replace(",", " ") +
            f"{e['category']}\n"
            f"{e['description']}\n"
            f"👤 {e['who']}\n\n"
        )
    await update.message.reply_text(text, parse_mode="HTML")


# ============ ДОБАВЛЕНИЕ РАСХОДА ============
async def expense_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_committee(update.effective_user.id):
        await update.message.reply_text("⛔ Добавлять расходы может только комитет.")
        return ConversationHandler.END
    context.user_data["type"] = "expense"
    await update.message.reply_text("💵 Введи сумму расхода (только число):")
    return AMOUNT


async def contribution_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_committee(update.effective_user.id):
        await update.message.reply_text("⛔ Добавлять взносы может только комитет.")
        return ConversationHandler.END
    context.user_data["type"] = "contribution"
    await update.message.reply_text("💸 Введи сумму взноса (только число):")
    return AMOUNT


async def amount_received(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text.replace(",", ".").replace(" ", "")
    try:
        amount = float(text)
        if amount <= 0:
            raise ValueError
    except ValueError:
        await update.message.reply_text("⚠️ Введи корректное число, например 1500")
        return AMOUNT

    context.user_data["amount"] = amount

    if context.user_data["type"] == "expense":
        # Показываем кнопки категорий
        keyboard = [[c] for c in CATEGORIES]
        await update.message.reply_text(
            "📂 Выбери категорию:",
            reply_markup=ReplyKeyboardMarkup(keyboard, one_time_keyboard=True, resize_keyboard=True)
        )
        return CATEGORY
    else:
        # Для взноса сразу спрашиваем от кого
        await update.message.reply_text("👤 От кого взнос? (ФИ родителя)")
        return WHO


async def category_received(update: Update, context: ContextTypes.DEFAULT_TYPE):
    category = update.message.text
    if category not in CATEGORIES:
        await update.message.reply_text("⚠️ Выбери категорию кнопкой.")
        return CATEGORY
    context.user_data["category"] = category
    await update.message.reply_text("✏️ Что купили? (короткое описание)")
    return DESCRIPTION


async def description_received(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["description"] = update.message.text
    await update.message.reply_text("👤 Кто добавляет? (твоё имя)")
    return WHO


async def who_received(update: Update, context: ContextTypes.DEFAULT_TYPE):
    who = update.message.text
    data = load_data()
    now = datetime.now().strftime("%d.%m.%Y")

    if context.user_data["type"] == "expense":
        data["expenses"].append({
            "date": now,
            "amount": context.user_data["amount"],
            "category": context.user_data["category"],
            "description": context.user_data["description"],
            "who": who,
        })
        save_data(data)
        await update.message.reply_text(
            f"✅ Расход добавлен: {context.user_data['amount']:,.0f} ₽".replace(",", " "),
        )
    else:
        data["contributions"].append({
            "date": now,
            "amount": context.user_data["amount"],
            "who": who,
        })
        save_data(data)
        await update.message.reply_text(
            f"✅ Взнос добавлен: {context.user_data['amount']:,.0f} ₽ от {who}".replace(",", " "),
        )

    context.user_data.clear()
    # Показываем свежий отчёт
    await report(update, context)
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

    # Диалог добавления
    conv = ConversationHandler(
        entry_points=[
            CommandHandler("расход", expense_start),
            CommandHandler("expense", expense_start),
            CommandHandler("взнос", contribution_start),
            CommandHandler("contribution", contribution_start),
        ],
        states={
            AMOUNT: [MessageHandler(filters.TEXT & ~filters.COMMAND, amount_received)],
            CATEGORY: [MessageHandler(filters.TEXT & ~filters.COMMAND, category_received)],
            DESCRIPTION: [MessageHandler(filters.TEXT & ~filters.COMMAND, description_received)],
            WHO: [MessageHandler(filters.TEXT & ~filters.COMMAND, who_received)],
        },
        fallbacks=[CommandHandler("отмена", cancel), CommandHandler("cancel", cancel)],
    )

    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("отчет", report))
    app.add_handler(CommandHandler("report", report))
    app.add_handler(CommandHandler("история", history))
    app.add_handler(CommandHandler("history", history))
    app.add_handler(conv)

    logger.info("Бот запущен!")
    app.run_polling()


if __name__ == "__main__":
    main()
