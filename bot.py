import os
import json
import logging
from datetime import datetime
from telegram import Update, ReplyKeyboardMarkup
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ConversationHandler, ContextTypes

BOT_TOKEN = os.environ.get("BOT_TOKEN")
if not BOT_TOKEN:
    raise RuntimeError("BOT_TOKEN не задан!")

DATA_FILE = "data.json"
COMMITTEE_IDS = [447774674, 6013055364]
BASE_FUND = 41000

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

CATEGORIES = ["📚 Материалы", "🍕 Питание", "🎉 Мероприятие", "🏠 Организационные", "🚗 Доставка"]

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
    text = f"📊 ФИНАНСОВЫЙ ОТЧЁТ\n\n💰 Базовый фонд: {BASE_FUND} ₽\n➕ Взносы: {contributions} ₽\n➖ Расходы: {expenses} ₽\n═══════════════════\n✅ Остаток: {current} ₽"
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

async def export_excel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    from openpyxl import Workbook
    from openpyxl.styles import Font
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

async def show_birthdays(update: Update, context: ContextTypes.DEFAULT_TYPE, month):
    birthdays = BIRTHDAYS.get(month, [])
    if not birthdays:
        await update.message.reply_text("Нет дней рождений")
        return
    text = f"🎂 {month.capitalize()}:\n"
    for name, date in birthdays:
        text += f"• {name} - {date}\n"
    await update.message.reply_text(text)

async def handle_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text
    if text == "📊 Отчёт":
        await report(update, context)
    elif text == "📈 История":
        await history(update, context)
    elif text == "➕ Расход":
        await expense_start(update, context)
    elif text == "➕ Взнос":
        await contribution_start(update, context)
    elif text == "📥 Экспорт":
        await export_excel(update, context)
    elif text.lower() in BIRTHDAYS:
        await show_birthdays(update, context, text.lower())

async def main():
    app = Application.builder().token(BOT_TOKEN).build()
    
    expense_conv = ConversationHandler(
        entry_points=[CommandHandler("expense", expense_start)],
        states={
            AMOUNT: [MessageHandler(filters.TEXT & ~filters.COMMAND, amount_received)],
            CATEGORY:
