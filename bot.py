"""
Бот учёта финансов и организации класса 1-К
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
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

BOT_TOKEN = os.environ.get("BOT_TOKEN")
BASE_FUND = 41000
DATA_FILE = "/data/data.json"
EXCEL_FILE = "/tmp/expenses.xlsx"
CATEGORIES = ["📚 Материалы", "🍕 Питание", "🎉 Мероприятие", "🏠 Организационные", "🚗 Доставка"]

AMOUNT, CATEGORY, DESCRIPTION, WHO = range(4)
NEWS_TEXT, FUNDRAISER_NAME, FUNDRAISER_GOAL, FUNDRAISER_DESC = range(4, 8)
EVENT_NAME, EVENT_DATE, EVENT_DESC = range(8, 11)

# ДНИ РОЖДЕНИЯ
BIRTHDAYS = {
    "Январь": [("Гилёва Валерия", "11.01"), ("Хрусталев Матвей", "21.01"), ("Власов Тимофей", "23.01")],
    "Февраль": [("Булычев Елисей", "13.02")],
    "Март": [("Антипов Мирон", "07.03"), ("Волохова Анна", "07.03"), ("Богданов Илья", "11.03"), ("Братковский Гордей", "19.03"), ("Иваненко Алина", "19.03"), ("Ворошилов Вячеслав", "28.03")],
    "Апрель": [("Писаренко Николь", "03.04")],
    "Май": [("Хан Нелли", "28.05")],
    "Июнь": [],
    "Июль": [("Лихачева София", "05.07"), ("Сокач Артём", "11.07"), ("Покровский Лев", "25.07")],
    "Август": [("Полуэктов Марк", "11.08"), ("Журавлёва Элина", "17.08"), ("Гаврилина Дарина", "24.08"), ("Селянин Владимир", "29.08")],
    "Сентябрь": [],
    "Октябрь": [("Побережная Есения", "10.10"), ("Власюк Елизавета", "16.10"), ("Володина Мария", "20.10"), ("Боровикова Инна Игоревна (учитель)", "27.10")],
    "Ноябрь": [("Чаптыков Александр", "06.11"), ("Надтачеев Богдан", "19.11"), ("Зорина Екатерина", "23.11"), ("Петров Павел", "26.11"), ("Пиндур Вера", "27.11")],
    "Декабрь": [("Иванов Артём", "21.12"), ("Брюханова Милана", "25.12"), ("Мухачёв Михаил", "27.12")],
}

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

def load_data():
    if os.path.exists(DATA_FILE):
        with open(DATA_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {"expenses": [], "news": [], "fundraisers": [], "events": []}

def save_data(data):
    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def is_committee(user_id):
    return True

def create_excel_report(data):
    """Создаёт Excel файл с расходами"""
    wb = Workbook()
    ws = wb.active
    ws.title = "Расходы"
    
    # Заголовок
    ws['A1'] = "ОТЧЁТ О РАСХОДАХ класса 1-К"
    ws['A1'].font = Font(size=14, bold=True)
    ws.merge_cells('A1:E1')
    
    # Дата создания
    ws['A2'] = f"Создано: {datetime.now().strftime('%d.%m.%Y %H:%M')}"
    
    # Финансовая сводка
    expenses = sum(e["amount"] for e in data["expenses"])
    current = BASE_FUND - expenses
    
    ws['A4'] = "ФИНАНСОВАЯ СВОДКА"
    ws['A4'].font = Font(bold=True)
    ws['A5'] = "Начальный фонд:"
    ws['B5'] = BASE_FUND
    ws['A6'] = "Потрачено:"
    ws['B6'] = expenses
    ws['A7'] = "Остаток:"
    ws['B7'] = current
    ws['B7'].font = Font(bold=True, color="008000")
    
    # Таблица расходов
    ws['A9'] = "ДЕТАЛЬНЫЙ СПИСОК РАСХОДОВ"
    ws['A9'].font = Font(bold=True)
    
    headers = ["Дата", "Сумма (₽)", "Категория", "Описание", "От кого"]
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=10, column=col)
        cell.value = header
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
        cell.alignment = Alignment(horizontal="center")
    
    # Данные
    row = 11
    for exp in data["expenses"]:
        ws.cell(row=row, column=1).value = exp['date']
        ws.cell(row=row, column=2).value = exp['amount']
        ws.cell(row=row, column=3).value = exp['category']
        ws.cell(row=row, column=4).value = exp['description']
        ws.cell(row=row, column=5).value = exp['who']
        row += 1
    
    # Ширина колонок
    ws.column_dimensions['A'].width = 12
    ws.column_dimensions['B'].width = 12
    ws.column_dimensions['C'].width = 20
    ws.column_dimensions['D'].width = 30
    ws.column_dimensions['E'].width = 20
    
    # Итого по категориям
    row += 2
    ws.cell(row=row, column=1).value = "ИТОГО ПО КАТЕГОРИЯМ"
    ws.cell(row=row, column=1).font = Font(bold=True)
    row += 1
    
    expenses_by_cat = {}
    for e in data["expenses"]:
        cat = e.get("category", "?")
        expenses_by_cat[cat] = expenses_by_cat.get(cat, 0) + e["amount"]
    
    for cat, amount in sorted(expenses_by_cat.items()):
        ws.cell(row=row, column=1).value = cat
        ws.cell(row=row, column=2).value = amount
        row += 1
    
    wb.save(EXCEL_FILE)
    return EXCEL_FILE

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    msg = f"👋 Привет! Бот класса 1-К.\nID: {user_id}\n\n"
    msg += "💰 ФИНАНСЫ:\n/отчет — финансы\n/история — расходы\n/экспорт — скачать Excel\n\n"
    msg += "🎂 ДНИ РОЖДЕНИЯ:\n/др — все ДР\n/октябрь и т.д. — по месяцам\n\n"
    msg += "📰 НОВОСТИ:\n/новости — все новости\n\n"
    msg += "🎯 СБОРЫ:\n/сборы — открытые сборы\n\n"
    msg += "📅 СОБЫТИЯ:\n/события — расписание\n\n"
    msg += "🛠 КОМИТЕТ:\n/расход — расход\n/новость — добавить новость\n/сбор — создать сбор\n/событие — добавить событие\n/удалить — удалить расход"
    await update.message.reply_text(msg)

async def export_expenses(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Экспортирует расходы в Excel"""
    await update.message.reply_text("⏳ Создаю Excel файл...")
    
    data = load_data()
    try:
        file_path = create_excel_report(data)
        with open(file_path, 'rb') as f:
            await update.message.reply_document(f, filename=f"Расходы_1К_{datetime.now().strftime('%d.%m.%Y')}.xlsx")
        logger.info("✅ Файл отправлен")
    except Exception as e:
        await update.message.reply_text(f"❌ Ошибка: {str(e)}")

async def report(update: Update, context: ContextTypes.DEFAULT_TYPE):
    data = load_data()
    expenses = sum(e["amount"] for e in data["expenses"])
    current = BASE_FUND - expenses
    msg = f"📊 ФИНАНСЫ\n\n💰 Фонд: {BASE_FUND:,} ₽\n➖ Расходы: {expenses:,} ₽\n📈 **Остаток: {current:,} ₽**"
    expenses_by_cat = {}
    for e in data["expenses"]:
        cat = e.get("category", "?")
        expenses_by_cat[cat] = expenses_by_cat.get(cat, 0) + e["amount"]
    if expenses_by_cat:
        msg += "\n\n**По категориям:**\n"
        for cat, amount in sorted(expenses_by_cat.items()):
            msg += f"{cat}: {amount:,} ₽\n"
    await update.message.reply_text(msg)

async def history(update: Update, context: ContextTypes.DEFAULT_TYPE):
    data = load_data()
    if not data["expenses"]:
        await update.message.reply_text("Расходов нет")
        return
    msg = "📋 Последние расходы:\n\n"
    for i, e in enumerate(data["expenses"][-10:], 1):
        msg += f"{i}. {e['date']} | {e['category']} | {e['amount']:,} ₽\n   {e['description']} ({e['who']})\n"
    await update.message.reply_text(msg)

async def show_news(update: Update, context: ContextTypes.DEFAULT_TYPE):
    data = load_data()
    if not data["news"]:
        await update.message.reply_text("📰 Новостей нет")
        return
    msg = "📰 **НОВОСТИ КЛАССА**\n\n"
    for n in data["news"][-10:]:
        msg += f"📌 {n['date']} | {n['author']}\n{n['text']}\n\n"
    await update.message.reply_text(msg)

async def add_news_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_committee(update.effective_user.id):
        await update.message.reply_text("❌ Нет прав")
        return ConversationHandler.END
    await update.message.reply_text("📝 Напиши новость:")
    return NEWS_TEXT

async def news_text_received(update: Update, context: ContextTypes.DEFAULT_TYPE):
    data = load_data()
    news = {
        "date": datetime.now().strftime("%d.%m.%Y %H:%M"),
        "text": update.message.text,
        "author": update.effective_user.first_name or "Комитет"
    }
    data["news"].append(news)
    save_data(data)
    await update.message.reply_text("✅ Новость добавлена!")
    return ConversationHandler.END

async def show_fundraisers(update: Update, context: ContextTypes.DEFAULT_TYPE):
    data = load_data()
    if not data["fundraisers"]:
        await update.message.reply_text("🎯 Активных сборов нет")
        return
    msg = "🎯 **ОТКРЫТЫЕ СБОРЫ**\n\n"
    for f in data["fundraisers"]:
        collected = sum(p["amount"] for p in f["payments"])
        progress = int((collected / f["goal"]) * 100) if f["goal"] > 0 else 0
        msg += f"💰 {f['name']}\n"
        msg += f"Цель: {f['goal']:,} ₽ | Собрано: {collected:,} ₽ ({progress}%)\n"
        msg += f"📝 {f['description']}\n"
        if f["payments"]:
            msg += "Участники: "
            msg += ", ".join([f"{p['name']} ({p['amount']:,}₽)" for p in f["payments"]])
        msg += "\n\n"
    await update.message.reply_text(msg)

async def create_fundraiser_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_committee(update.effective_user.id):
        await update.message.reply_text("❌ Нет прав")
        return ConversationHandler.END
    await update.message.reply_text("Название сбора?")
    return FUNDRAISER_NAME

async def fundraiser_name_received(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["fundraiser_name"] = update.message.text
    await update.message.reply_text("Целевая сумма (₽)?")
    return FUNDRAISER_GOAL

async def fundraiser_goal_received(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        context.user_data["fundraiser_goal"] = float(update.message.text)
    except:
        await update.message.reply_text("Введи число")
        return FUNDRAISER_GOAL
    await update.message.reply_text("Описание?")
    return FUNDRAISER_DESC

async def fundraiser_desc_received(update: Update, context: ContextTypes.DEFAULT_TYPE):
    data = load_data()
    fundraiser = {
        "id": len(data["fundraisers"]),
        "name": context.user_data["fundraiser_name"],
        "goal": context.user_data["fundraiser_goal"],
        "description": update.message.text,
        "created": datetime.now().strftime("%d.%m.%Y"),
        "payments": []
    }
    data["fundraisers"].append(fundraiser)
    save_data(data)
    await update.message.reply_text(f"✅ Сбор '{fundraiser['name']}' создан!")
    context.user_data.clear()
    return ConversationHandler.END

async def show_events(update: Update, context: ContextTypes.DEFAULT_TYPE):
    data = load_data()
    if not data["events"]:
        await update.message.reply_text("📅 События не запланированы")
        return
    msg = "📅 **СОБЫТИЯ И МЕРОПРИЯТИЯ**\n\n"
    for e in sorted(data["events"], key=lambda x: x["date"]):
        msg += f"📌 {e['date']} — {e['name']}\n{e['description']}\n\n"
    await update.message.reply_text(msg)

async def create_event_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_committee(update.effective_user.id):
        await update.message.reply_text("❌ Нет прав")
        return ConversationHandler.END
    await update.message.reply_text("Название события?")
    return EVENT_NAME

async def event_name_received(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["event_name"] = update.message.text
    await update.message.reply_text("Дата? (дд.мм.гггг)")
    return EVENT_DATE

async def event_date_received(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["event_date"] = update.message.text
    await update.message.reply_text("Описание?")
    return EVENT_DESC

async def event_desc_received(update: Update, context: ContextTypes.DEFAULT_TYPE):
    data = load_data()
    event = {
        "name": context.user_data["event_name"],
        "date": context.user_data["event_date"],
        "description": update.message.text
    }
    data["events"].append(event)
    save_data(data)
    await update.message.reply_text(f"✅ Событие '{event['name']}' добавлено!")
    context.user_data.clear()
    return ConversationHandler.END

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
    await update.message.reply_text(f"✅ Удалён: {deleted['amount']:,} ₽ ({deleted['description']})")

async def expense_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_committee(update.effective_user.id):
        await update.message.reply_text("❌ Нет прав")
        return ConversationHandler.END
    await update.message.reply_text("💰 Сумма?")
    return AMOUNT

async def amount_received(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        context.user_data["amount"] = float(update.message.text)
    except:
        await update.message.reply_text("❌ Число")
        return AMOUNT
    buttons = [[c] for c in CATEGORIES]
    await update.message.reply_text("Категория?", reply_markup=ReplyKeyboardMarkup(buttons, one_time_keyboard=True))
    return CATEGORY

async def category_received(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["category"] = update.message.text
    await update.message.reply_text("Описание?")
    return DESCRIPTION

async def description_received(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["description"] = update.message.text
    await update.message.reply_text("От кого?")
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
    await update.message.reply_text(f"✅ Расход добавлен!")
    context.user_data.clear()
    return ConversationHandler.END

async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.clear()
    await update.message.reply_text("❌ Отменено")
    return ConversationHandler.END

async def birthdays_all(update: Update, context: ContextTypes.DEFAULT_TYPE):
    msg = "🎂 **ДНЕЙ РОЖДЕНИЯ**\n\n"
    for m, bdays in BIRTHDAYS.items():
        if bdays:
            msg += f"**{m}:** "
            msg += ", ".join([f"{name} ({date})" for name, date in bdays])
            msg += "\n"
    await update.message.reply_text(msg)

async def october(update: Update, context: ContextTypes.DEFAULT_TYPE):
    msg = "🎂 **ОКТЯБРЬ:**\n"
    for name, date in BIRTHDAYS["Октябрь"]:
        msg += f"{date} — {name}\n"
    await update.message.reply_text(msg)

def main():
    if not BOT_TOKEN:
        raise RuntimeError("Нет BOT_TOKEN!")
    
    app = Application.builder().token(BOT_TOKEN).build()
    
    expense_conv = ConversationHandler(
        entry_points=[CommandHandler("расход", expense_start)],
        states={
            AMOUNT: [MessageHandler(filters.TEXT & ~filters.COMMAND, amount_received)],
            CATEGORY: [MessageHandler(filters.TEXT & ~filters.COMMAND, category_received)],
            DESCRIPTION: [MessageHandler(filters.TEXT & ~filters.COMMAND, description_received)],
            WHO: [MessageHandler(filters.TEXT & ~filters.COMMAND, who_received)],
        },
        fallbacks=[CommandHandler("cancel", cancel), CommandHandler("отмена", cancel)],
    )
    
    news_conv = ConversationHandler(
        entry_points=[CommandHandler("новость", add_news_start)],
        states={NEWS_TEXT: [MessageHandler(filters.TEXT & ~filters.COMMAND, news_text_received)]},
        fallbacks=[CommandHandler("cancel", cancel)],
    )
    
    fundraiser_conv = ConversationHandler(
        entry_points=[CommandHandler("сбор", create_fundraiser_start)],
        states={
            FUNDRAISER_NAME: [MessageHandler(filters.TEXT & ~filters.COMMAND, fundraiser_name_received)],
            FUNDRAISER_GOAL: [MessageHandler(filters.TEXT & ~filters.COMMAND, fundraiser_goal_received)],
            FUNDRAISER_DESC: [MessageHandler(filters.TEXT & ~filters.COMMAND, fundraiser_desc_received)],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
    )
    
    event_conv = ConversationHandler(
        entry_points=[CommandHandler("событие", create_event_start)],
        states={
            EVENT_NAME: [MessageHandler(filters.TEXT & ~filters.COMMAND, event_name_received)],
            EVENT_DATE: [MessageHandler(filters.TEXT & ~filters.COMMAND, event_date_received)],
            EVENT_DESC: [MessageHandler(filters.TEXT & ~filters.COMMAND, event_desc_received)],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
    )
    
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("отчет", report))
    app.add_handler(CommandHandler("история", history))
    app.add_handler(CommandHandler("экспорт", export_expenses))
    app.add_handler(CommandHandler("новости", show_news))
    app.add_handler(CommandHandler("сборы", show_fundraisers))
    app.add_handler(CommandHandler("события", show_events))
    app.add_handler(CommandHandler("др", birthdays_all))
    app.add_handler(CommandHandler("октябрь", october))
    app.add_handler(CommandHandler("удалить", delete_expense))
    app.add_handler(news_conv)
    app.add_handler(fundraiser_conv)
    app.add_handler(event_conv)
    app.add_handler(expense_conv)
    
    logger.info("✅ Бот класса 1-К запущен!")
    app.run_polling()

if __name__ == "__main__":
    main()
