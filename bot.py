import os
import json
import logging
from datetime import datetime
from telegram import Update, ReplyKeyboardMarkup
from telegram.ext import (
    Application, CommandHandler, MessageHandler, filters,
    ConversationHandler, ContextTypes, PicklePersistence
)
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

BOT_TOKEN = os.environ.get("BOT_TOKEN")
BASE_FUND = 41000
DATA_FILE = "data.json"
PERSISTENCE_FILE = "conv_state"
CATEGORIES = ["📚 Материалы", "🍕 Питание", "🎉 Мероприятие", "🏠 Организационные", "🚗 Доставка"]
COMMITTEE_IDS = [447774674, 6013055364]

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

def create_excel_report(data):
    wb = Workbook()
    ws = wb.active
    ws.title = "Финансы"

    ws['A1'] = "ОТЧЁТ О ФИНАНСАХ класса 1-К"
    ws['A1'].font = Font(size=14, bold=True)
    ws.merge_cells('A1:E1')

    ws['A2'] = f"Создано: {datetime.now().strftime('%d.%m.%Y %H:%M')}"

    expenses = sum(e["amount"] for e in data["expenses"])
    contributions = sum(c["amount"] for c in data.get("contributions", []))
    current = BASE_FUND + contributions - expenses

    ws['A4'] = "ФИНАНСОВАЯ СВОДКА"
    ws['A4'].font = Font(bold=True)
    ws['A5'] = "Базовый фонд:"
    ws['B5'] = BASE_FUND
    ws['A6'] = "Взносы:"
    ws['B6'] = contributions
    ws['A7'] = "Расходы:"
    ws['B7'] = expenses
    ws['A8'] = "Остаток:"
    ws['B8'] = current
    ws['B8'].font = Font(bold=True, color="008000")

    ws['A10'] = "РАСХОДЫ"
    ws['A10'].font = Font(bold=True)

    headers = ["Дата", "Сумма (₽)", "Категория", "Описание", "От кого"]
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=11, column=col)
        cell.value = header
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")

    row = 12
    for exp in data["expenses"]:
        ws.cell(row=row, column=1).value = exp['date']
        ws.cell(row=row, column=2).value = exp['amount']
        ws.cell(row=row, column=3).value = exp['category']
        ws.cell(row=row, column=4).value = exp['description']
        ws.cell(row=row, column=5).value = exp['who']
        row += 1

    ws.column_dimensions['A'].width = 12
    ws.column_dimensions['B'].width = 12
    ws.column_dimensions['C'].width = 20
    ws.column_dimensions['D'].width = 30
    ws.column_dimensions['E'].width = 15

    return wb

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    keyboard = [
        ["📊 Отчёт", "📈 История"],
        ["📥 Экспорт"]
    ]

    if is_committee(update.effective_user.id):
        keyboard.insert(0, ["➕ Расход", "➕ Взнос"])

    reply_markup = ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
    await update.message.reply_text(
        "🏫 Финансы класса 1-К\n\nВыбери действие:",
        reply_markup=reply_markup
    )

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
✅ Остаток: {current} ₽

На одного ребёнка: {current/30:.2f} ₽
"""

    if data["expenses"]:
        text += "\n📋 Последние расходы:\n"
        for exp in data["expenses"][-5:]:
            text += f"• {exp['date']} - {exp['amount']}₽ ({exp['category']})\n"

    await update.message.reply_text(text)

async def history(update: Update, context: ContextTypes.DEFAULT_TYPE):
    data = load_data()

    if not data["expenses"]:
        await update.message.reply_text("📭 Нет записей о расходах")
        return

    text = "📈 ИСТОРИЯ РАСХОДОВ\n\n"
    for exp in data["expenses"]:
        text += f"📅 {exp['date']}\n"
        text += f"   💵 {exp['amount']} ₽\n"
        text += f"   📂 {exp['category']}\n"
        text += f"   📝 {exp['description']}\n\n"

    await update.message.reply_text(text)

async def expense_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_committee(update.effective_user.id):
        await update.message.reply_text("❌ Только казначей может добавлять расходы")
        return ConversationHandler.END

    await update.message.reply_text("💵 Сумма расхода (₽):")
    return AMOUNT

async def amount_received(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        context.user_data['amount'] = float(update.message.text)
    except:
        await update.message.reply_text("❌ Введи сумму правильно")
        return AMOUNT

    keyboard = [[cat] for cat in CATEGORIES]
    reply_markup = ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
    await update.message.reply_text("📂 Выбери категорию:", reply_markup=reply_markup)
    return CATEGORY

async def category_received(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data['category'] = update.message.text
    await update.message.reply_text("📝 Описание:")
    return DESCRIPTION

async def description_received(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data['description'] = update.message.text
    await update.message.reply_text("👤 От кого (имя):")
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
    await update.message.reply_text(f"✅ Расход {context.user_data['amount']} ₽ добавлен!")

    keyboard = [["📊 Отчёт", "📈 История"], ["➕ Расход", "➕ Взнос"], ["📥 Экспорт"]]
    reply_markup = ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
    await update.message.reply_text("Выбери действие:", reply_markup=reply_markup)

    return ConversationHandler.END

async def contribution_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_committee(update.effective_user.id):
        await update.message.reply_text("❌ Только казначей может добавлять взносы")
        return ConversationHandler.END

    await update.message.reply_text("💵 Сумма взноса (₽):")
    return CONTRIB_AMOUNT

async def contrib_amount_received(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        context.user_data['contrib_amount'] = float(update.message.text)
    except:
        await update.message.reply_text("❌ Введи сумму правильно")
        return CONTRIB_AMOUNT

    await update.message.reply_text("👤 От кого (имя):")
    return CONTRIB_WHO

async def contrib_who_received(update: Update, context: ContextTypes.DEFAULT_TYPE):
    data = load_data()

    data["contributions"].append({
        "date": datetime.now().strftime('%d.%m.%Y'),
        "amount": context.user_data['contrib_amount'],
        "who": update.message.text
    })

    save_data(data)
    await update.message.reply_text(f"✅ Взнос {context.user_data['contrib_amount']} ₽ добавлен!")

    keyboard = [["📊 Отчёт", "📈 История"], ["➕ Расход", "➕ Взнос"], ["📥 Экспорт"]]
    reply_markup = ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
    await update.message.reply_text("Выбери действие:", reply_markup=reply_markup)

    return ConversationHandler.END

async def export_excel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    data = load_data()
    wb = create_excel_report(data)
    filename = "finansy_klassa.xlsx"
    wb.save(filename)

    with open(filename, 'rb') as f:
        await update.message.reply_document(f, filename=filename)

    os.remove(filename)

async def show_birthdays(update: Update, month_key: str, month_name: str):
    birthdays = BIRTHDAYS[month_key]

    if not birthdays:
        await update.message.reply_text(f"🎂 В {month_name} дней рождений нет")
        return

    text = f"🎂 ДНЕЙ РОЖДЕНИЙ В {month_name.upper()}\n\n"

    for name, date in birthdays:
        text += f"• {name} - {date}\n"

    await update.message.reply_text(text)

async def january(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await show_birthdays(update, "январь", "Январь")

async def february(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await show_birthdays(update, "февраль", "Февраль")

async def march(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await show_birthdays(update, "март", "Март")

async def april(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await show_birthdays(update, "апрель", "Апрель")

async def may(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await show_birthdays(update, "май", "Май")

async def june(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await show_birthdays(update, "июнь", "Июнь")

async def july(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await show_birthdays(update, "июль", "Июль")

async def august(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await show_birthdays(update, "август", "Август")

async def september(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await show_birthdays(update, "сентябрь", "Сентябрь")

async def october(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await show_birthdays(update, "октябрь", "Октябрь")

async def november(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await show_birthdays(update, "ноябрь", "Ноябрь")

async def december(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await show_birthdays(update, "декабрь", "Декабрь")

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
    if not BOT_TOKEN:
        raise RuntimeError("BOT_TOKEN не задан!")

    persistence = PicklePersistence(filepath=PERSISTENCE_FILE)
    app = Application.builder().token(BOT_TOKEN).persistence(persistence).build()

    # ConversationHandlers
    expense_handler = ConversationHandler(
        entry_points=[CommandHandler("expense", expense_start)],
        states={
            AMOUNT: [MessageHandler(filters.TEXT & ~filters.COMMAND, amount_received)],
            CATEGORY: [MessageHandler(filters.TEXT & ~filters.COMMAND, category_received)],
            DESCRIPTION: [MessageHandler(filters.TEXT & ~filters.COMMAND, description_received)],
            WHO: [MessageHandler(filters.TEXT & ~filters.COMMAND, who_received)],
        },
        fallbacks=[CommandHandler("start", start)],
        name="expense_handler",
        persistent=True,
    )

    contribution_handler = ConversationHandler(
        entry_points=[CommandHandler("contribution", contribution_start)],
        states={
            CONTRIB_AMOUNT: [MessageHandler(filters.TEXT & ~filters.COMMAND, contrib_amount_received)],
            CONTRIB_WHO: [MessageHandler(filters.TEXT & ~filters.COMMAND, contrib_who_received)],
        },
        fallbacks=[CommandHandler("start", start)],
        name="contribution_handler",
        persistent=True,
    )

    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("report", report))
    app.add_handler(CommandHandler("history", history))
    app.add_handler(CommandHandler("export", export_excel))

    # Месяцы команды (только английский!)
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

    app.add_handler(expense_handler)
    app.add_handler(contribution_handler)
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text))

    await app.initialize()
    await app.start()
    await app.updater.start_polling()
    await app.idle()

if __name__ == "__main__":
    import asyncio
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        logger.info("Бот остановлен")
    except Exception as e:
        logger.error(f"Ошибка: {e}")
