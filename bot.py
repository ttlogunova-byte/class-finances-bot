import os
import json
import logging
from datetime import datetime
from telegram import Update, ReplyKeyboardMarkup
from telegram.ext import (
    Application, CommandHandler, MessageHandler, filters,
    ConversationHandler, ContextTypes
)
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

BOT_TOKEN = os.environ.get("BOT_TOKEN")
if not BOT_TOKEN:
    raise RuntimeError("BOT_TOKEN не задан!")

BASE_FUND = 41000
DATA_FILE = "data.json"
CATEGORIES = ["📚 Материалы", "🍕 Питание", "🎉 Мероприятие", "🏠 Организационные", "🚗 Доставка"]
COMMITTEE_IDS = [447774674, 6013055364]

# States for conversations
AMOUNT, CATEGORY, DESCRIPTION, WHO = range(4)
CONTRIB_AMOUNT, CONTRIB_WHO = range(4, 6)

BIRTHDAYS = {
    "январь": [("Гилёва Валерия", "11.01"), ("Хрусталев Матвей", "21.01"), ("Власов Тимофей", "23.01")],
    "февраль": [("Булычев Елисей", "13.02")],
    "март": [("Антипов Мирон", "07.03"), ("Волохова Анна", "07.03"), ("Богданов Илья", "11.03"), ("Братковский Гордей", "19.03"), ("Иваненко Алина", "19.03"), ("Ворошилов Вячеслав", "28.03")],
    "апрель": [("Писаренко Николь", "03.04")],
    "май": [("Хан Нелли", "28.05")],
    "июнь": [],
    "июль": [("Лихачева София", "05.07"), ("Сокач Артём", "11.07"), ("Покровский Лев", "25.07")],
    "август": [("Полуэктов Марк", "11.08"), ("Журавлёва Элина", "17.08"), ("Гаврилина Дарина", "24.08"), ("Селянин Владимир", "29.08")],
    "сентябрь": [],
    "октябрь": [("Побережная Есения", "10.10"), ("Власюк Елизавета", "16.10"), ("Володина Мария", "20.10"), ("Боровикова Инна Игоревна (учитель)", "27.10")],
    "ноябрь": [("Чаптыков Александр", "06.11"), ("Надтачеев Богдан", "19.11"), ("Зорина Екатерина", "23.11"), ("Петров Павел", "26.11"), ("Пиндур Вера", "27.11")],
    "декабрь": [("Иванов Артём", "21.12"), ("Брюханова Милана", "25.12"), ("Мухачёв Михаил", "27.12")],
}

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

def load_data():
    if os.path.exists(DATA_FILE):
        try:
            with open(DATA_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except:
            pass
    return {"expenses": [], "contributions": []}

def save_data(data):
    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def is_committee(user_id):
    return user_id in COMMITTEE_IDS

# COMMANDS
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    keyboard = [["📊 Отчёт", "📈 История"], ["📥 Экспорт"]]
    if is_committee(update.effective_user.id):
        keyboard.insert(0, ["➕ Расход", "➕ Взнос"])

    reply_markup = ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
    await update.message.reply_text("🏫 Финансы класса 1-К", reply_markup=reply_markup)

async def report(update: Update, context: ContextTypes.DEFAULT_TYPE):
    data = load_data()
    expenses = sum(e["amount"] for e in data["expenses"])
    contributions = sum(c["amount"] for c in data.get("contributions", []))
    current = BASE_FUND + contributions - expenses

    text = f"""📊 ФИНАНСОВЫЙ ОТЧЁТ

💰 Базовый фонд: {BASE_FUND} ₽
➕ Взносы: {contributions} ₽
➖ Расходы: {expenses} ₽
═══════════════════
✅ Остаток: {current} ₽"""

    await update.message.reply_text(text)

async def history(update: Update, context: ContextTypes.DEFAULT_TYPE):
    data = load_data()
    if not data["expenses"]:
        await update.message.reply_text("📭 Нет расходов")
        return

    text = "📈 ИСТОРИЯ РАСХОДОВ\n\n"
    for exp in data["expenses"][-10:]:
        text += f"📅 {exp['date']} - {exp['amount']}₽ ({exp['category']})\n"

    await update.message.reply_text(text)

async def expense_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_committee(update.effective_user.id):
        await update.message.reply_text("❌ Только казначей!")
        return ConversationHandler.END

    await update.message.reply_text("💵 Сумма (₽):")
    return AMOUNT

async def amount_received(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        context.user_data['amount'] = float(update.message.text)
        keyboard = [[cat] for cat in CATEGORIES]
        reply_markup = ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
        await update.message.reply_text("📂 Категория:", reply_markup=reply_markup)
        return CATEGORY
    except:
        await update.message.reply_text("❌ Неверная сумма")
        return AMOUNT

async def category_received(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data['category'] = update.message.text
    await update.message.reply_text("📝 Описание:")
    return DESCRIPTION

async def description_received(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data['description'] = update.message.text
    await update.message.reply_text("👤 От кого:")
    return WHO

async def who_received(update: Update, context: ContextTypes.DEFAULT_TYPE):
    data = load_data()
    data["expenses"].append({
        "date": datetime.now().strftime('%d.%m.%Y'),
        "amount": context.user_data['amount'],
        "category": context.user_data['category'],
        "description": context.user_data['description'],
        "who": update.message.text
    })
    save_data(data)

    keyboard = [["📊 Отчёт", "📈 История"], ["➕ Расход", "➕ Взнос"], ["📥 Экспорт"]]
    reply_markup = ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
    await update.message.reply_text("✅ Расход добавлен!", reply_markup=reply_markup)
    return ConversationHandler.END

async def contribution_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_committee(update.effective_user.id):
        await update.message.reply_text("❌ Только казначей!")
        return ConversationHandler.END

    await update.message.reply_text("💵 Сумма (₽):")
    return CONTRIB_AMOUNT

async def contrib_amount_received(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        context.user_data['contrib_amount'] = float(update.message.text)
        await update.message.reply_text("👤 От кого:")
        return CONTRIB_WHO
    except:
        await update.message.reply_text("❌ Неверная сумма")
        return CONTRIB_AMOUNT

async def contrib_who_received(update: Update, context: ContextTypes.DEFAULT_TYPE):
    data = load_data()
    data["contributions"].append({
        "date": datetime.now().strftime('%d.%m.%Y'),
        "amount": context.user_data['contrib_amount'],
        "who": update.message.text
    })
    save_data(data)

    keyboard = [["📊 Отчёт", "📈 История"], ["➕ Расход", "➕ Взнос"], ["📥 Экспорт"]]
    reply_markup = ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
    await update.message.reply_text("✅ Взнос добавлен!", reply_markup=reply_markup)
    return ConversationHandler.END

async def export_excel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    data = load_data()
    wb = Workbook()
    ws = wb.active
    ws.title = "Финансы"

    ws['A1'] = "ОТЧЁТ О ФИНАНСАХ класса 1-К"
    ws['A1'].font = Font(size=12, bold=True)

    expenses = sum(e["amount"] for e in data["expenses"])
    contributions = sum(c["amount"] for c in data.get("contributions", []))
    current = BASE_FUND + contributions - expenses

    ws['A3'] = "Базовый фонд:"
    ws['B3'] = BASE_FUND
    ws['A4'] = "Взносы:"
    ws['B4'] = contributions
    ws['A5'] = "Расходы:"
    ws['B5'] = expenses
    ws['A6'] = "Остаток:"
    ws['B6'] = current

    ws['A8'] = "Расходы по датам"
    ws['A8'].font = Font(bold=True)

    row = 9
    for exp in data["expenses"]:
        ws.cell(row=row, column=1).value = exp['date']
        ws.cell(row=row, column=2).value = exp['amount']
        ws.cell(row=row, column=3).value = exp['category']
        row += 1

    filename = "finansy.xlsx"
    wb.save(filename)

    with open(filename, 'rb') as f:
        await update.message.reply_document(f, filename=filename)
    os.remove(filename)

async def january(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = "🎂 Январь:\n"
    for name, date in BIRTHDAYS["январь"]:
        text += f"• {name} - {date}\n"
    await update.message.reply_text(text)

async def february(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = "🎂 Февраль:\n"
    for name, date in BIRTHDAYS["февраль"]:
        text += f"• {name} - {date}\n"
    await update.message.reply_text(text if len(text) > 10 else "Нет дней рождений")

async def march(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = "🎂 Март:\n"
    for name, date in BIRTHDAYS["март"]:
        text += f"• {name} - {date}\n"
    await update.message.reply_text(text)

async def april(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = "🎂 Апрель:\n"
    for name, date in BIRTHDAYS["апрель"]:
        text += f"• {name} - {date}\n"
    await update.message.reply_text(text if len(text) > 10 else "Нет дней рождений")

async def may(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = "🎂 Май:\n"
    for name, date in BIRTHDAYS["май"]:
        text += f"• {name} - {date}\n"
    await update.message.reply_text(text if len(text) > 10 else "Нет дней рождений")

async def june(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Нет дней рождений")

async def july(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = "🎂 Июль:\n"
    for name, date in BIRTHDAYS["июль"]:
        text += f"• {name} - {date}\n"
    await update.message.reply_text(text)

async def august(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = "🎂 Август:\n"
    for name, date in BIRTHDAYS["август"]:
        text += f"• {name} - {date}\n"
    await update.message.reply_text(text)

async def september(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Нет дней рождений")

async def october(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = "🎂 Октябрь:\n"
    for name, date in BIRTHDAYS["октябрь"]:
        text += f"• {name} - {date}\n"
    await update.message.reply_text(text)

async def november(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = "🎂 Ноябрь:\n"
    for name, date in BIRTHDAYS["ноябрь"]:
        text += f"• {name} - {date}\n"
    await update.message.reply_text(text)

async def december(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = "🎂 Декабрь:\n"
    for name, date in BIRTHDAYS["декабрь"]:
        text += f"• {name} - {date}\n"
    await update.message.reply_text(text)

async def handle_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text.lower()
    if text == "📊 отчёт":
        await report(update, context)
    elif text == "📈 история":
        await history(update, context)
    elif text == "➕ расход":
        await expense_start(update, context)
    elif text == "➕ взнос":
        await contribution_start(update, context)
    elif text == "📥 экспорт":
        await export_excel(update, context)

async def main():
    app = Application.builder().token(BOT_TOKEN).build()

    # Conversations
    expense_conv = ConversationHandler(
        entry_points=[CommandHandler("expense", expense_start)],
        states={
            AMOUNT: [MessageHandler(filters.TEXT & ~filters.COMMAND, amount_received)],
            CATEGORY: [MessageHandler(filters.TEXT & ~filters.COMMAND, category_received)],
            DESCRIPTION: [MessageHandler(filters.TEXT & ~filters.COMMAND, description_received)],
            WHO: [MessageHandler(filters.TEXT & ~filters.COMMAND, who_received)],
        },
        fallbacks=[CommandHandler("start", start)],
    )

    contrib_conv = ConversationHandler(
        entry_points=[CommandHandler("contribution", contribution_start)],
        states={
            CONTRIB_AMOUNT: [MessageHandler(filters.TEXT & ~filters.COMMAND, contrib_amount_received)],
            CONTRIB_WHO: [MessageHandler(filters.TEXT & ~filters.COMMAND, contrib_who_received)],
        },
        fallbacks=[CommandHandler("start", start)],
    )

    # Add handlers
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("report", report))
    app.add_handler(CommandHandler("history", history))
    app.add_handler(CommandHandler("export", export_excel))
    app.add_handler(CommandHandler("january", january))
    app.add_handler(CommandHandler("february", february))
    app.add_handler(CommandHandler("march", march))
    app.add_handler(CommandHandler("april", april))
    app.add_handler(CommandHandler("may", may))
    app.add_handler(CommandHandler("june", june))
    app.add_handler(CommandHandler("july", july))
    app.add_handler(CommandHandler("august", august))
    app.add_handler(CommandHandler("september", september))
    app.add_handler(CommandHandler("october", october))
    app.add_handler(CommandHandler("november", november))
    app.add_handler(CommandHandler("december", december))
    app.add_handler(expense_conv)
    app.add_handler(contrib_conv)
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text))

    logger.info("✅ БОТ ЗАПУЩЕН!")
    await app.initialize()
    await app.start()
    await app.updater.start_polling()
    await app.idle()

if __name__ == "__main__":
    import asyncio
    try:
        asyncio.run(main())
    except Exception as e:
        logger.error(f"Ошибка: {e}")
