import os
import json
os.makedirs("/data", exist_ok=True)

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
DATA_FILE = "/data/data.json"
PERSISTENCE_FILE = "/data/conv_state"
CATEGORIES = ["📚 Материалы", "🍕 Питание", "🎉 Мероприятие", "🏠 Организационные", "🚗 Доставка"]
COMMITTEE_IDS = [447774674, 6013055364]

AMOUNT, CATEGORY, DESCRIPTION, WHO = range(4)
NEWS_TEXT, FUNDRAISER_NAME, FUNDRAISER_GOAL, FUNDRAISER_DESC = range(4, 8)
EVENT_NAME, EVENT_DATE, EVENT_DESC = range(8, 11)
CONTRIB_AMOUNT, CONTRIB_WHO = range(11, 13)

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
        try:
            with open(DATA_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except:
            pass
    return {"expenses": [], "contributions": [], "news": [], "fundraisers": [], "events": []}

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
    ws.column_dimensions['E'].width = 20
    
    wb.save("/tmp/expenses.xlsx")
    return "/tmp/expenses.xlsx"

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    name = update.effective_user.first_name or "Друг"
    
    if is_committee(user_id):
        msg = f"👋 Привет, {name}! Бот класса 1-К (КОМИТЕТ)\n"
        msg += f"🆔 ID: {user_id}\n\n"
        msg += "💰 ФИНАНСЫ:\n/report — финансы\n/history — расходы\n/export — скачать Excel\n\n"
        msg += "🎂 ДНИ РОЖДЕНИЯ:\n/birthdays — все ДР\n\n"
        msg += "📰 НОВОСТИ:\n/news — новости\n\n"
        msg += "🎯 СБОРЫ:\n/fundraisers — сборы\n\n"
        msg += "📅 СОБЫТИЯ:\n/events — события\n\n"
        msg += "🛠️ КОМИТЕТ:\n/expense — расход\n/contribution — взнос\n/addnews — новость\n/newevent — событие\n/deletexp — удалить расход"
    else:
        msg = f"👋 Привет, {name}! Бот класса 1-К\n"
        msg += f"🆔 ID: {user_id}\n\n"
        msg += "💰 ФИНАНСЫ:\n/report — финансы\n/history — расходы\n/export — скачать Excel\n\n"
        msg += "🎂 ДНИ РОЖДЕНИЯ:\n/birthdays — все ДР\n\n"
        msg += "📰 НОВОСТИ:\n/news — новости\n\n"
        msg += "🎯 СБОРЫ:\n/fundraisers — сборы\n\n"
        msg += "📅 СОБЫТИЯ:\n/events — события"
    
    await update.message.reply_text(msg)

async def report(update: Update, context: ContextTypes.DEFAULT_TYPE):
    data = load_data()
    expenses = sum(e["amount"] for e in data["expenses"])
    contributions = sum(c["amount"] for c in data.get("contributions", []))
    current = BASE_FUND + contributions - expenses
    msg = f"📊 ФИНАНСЫ\n\n"
    msg += f"💰 Базовый фонд: {BASE_FUND:,} ₽\n"
    msg += f"➕ Взносы: {contributions:,} ₽\n"
    msg += f"➖ Расходы: {expenses:,} ₽\n"
    msg += f"📈 **Остаток: {current:,} ₽**"
    await update.message.reply_text(msg)

async def history(update: Update, context: ContextTypes.DEFAULT_TYPE):
    data = load_data()
    if not data["expenses"]:
        await update.message.reply_text("Расходов нет")
        return
    msg = "📋 Расходы:\n\n"
    for i, e in enumerate(data["expenses"][-10:], 1):
        msg += f"{i}. {e['date']} | {e['amount']:,} ₽ | {e['description']}\n"
    await update.message.reply_text(msg)

async def export_expenses(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("⏳ Создаю Excel...")
    data = load_data()
    try:
        file_path = create_excel_report(data)
        with open(file_path, 'rb') as f:
            await update.message.reply_document(
                document=f,
                filename=f"Финансы_1К_{datetime.now().strftime('%d.%m.%Y')}.xlsx",
                caption="📊 Финансовый отчёт"
            )
        os.remove(file_path)
    except Exception as e:
        await update.message.reply_text(f"❌ Ошибка: {str(e)}")

async def show_news(update: Update, context: ContextTypes.DEFAULT_TYPE):
    data = load_data()
    if not data["news"]:
        await update.message.reply_text("Новостей нет")
        return
    msg = "📰 НОВОСТИ:\n\n"
    for n in data["news"][-10:]:
        msg += f"📌 {n['date']}: {n['text']}\n\n"
    await update.message.reply_text(msg)

async def show_birthdays(update: Update, context: ContextTypes.DEFAULT_TYPE):
    msg = "🎂 ДНЕЙ РОЖДЕНИЯ:\n\n"
    for m, bdays in BIRTHDAYS.items():
        if bdays:
            msg += f"{m}: "
            msg += ", ".join([f"{name} ({date})" for name, date in bdays])
            msg += "\n"
    await update.message.reply_text(msg)

async def show_fundraisers(update: Update, context: ContextTypes.DEFAULT_TYPE):
    data = load_data()
    if not data["fundraisers"]:
        await update.message.reply_text("Сборов нет")
        return
    msg = "🎯 СБОРЫ:\n\n"
    for f in data["fundraisers"]:
        msg += f"💰 {f['name']}\nЦель: {f['goal']:,} ₽\n{f['description']}\n\n"
    await update.message.reply_text(msg)

async def show_events(update: Update, context: ContextTypes.DEFAULT_TYPE):
    data = load_data()
    if not data["events"]:
        await update.message.reply_text("События не запланированы")
        return
    msg = "📅 СОБЫТИЯ:\n\n"
    for e in data["events"]:
        msg += f"📌 {e['date']} — {e['name']}\n{e['description']}\n\n"
    await update.message.reply_text(msg)

async def expense_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_committee(update.effective_user.id):
        await update.message.reply_text("❌ Нет прав")
        return ConversationHandler.END
    await update.message.reply_text("💰 Сумма расхода?")
    return AMOUNT

async def amount_received(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        context.user_data["amount"] = float(update.message.text)
        buttons = [[c] for c in CATEGORIES]
        await update.message.reply_text("Категория?", reply_markup=ReplyKeyboardMarkup(buttons, one_time_keyboard=True))
        return CATEGORY
    except:
        await update.message.reply_text("Введи число")
        return AMOUNT

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
    data["expenses"].append({
        "date": datetime.now().strftime("%d.%m.%Y"),
        "amount": context.user_data["amount"],
        "category": context.user_data["category"],
        "description": context.user_data["description"],
        "who": update.message.text
    })
    save_data(data)
    await update.message.reply_text("✅ Расход добавлен!")
    return ConversationHandler.END

async def contribution_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_committee(update.effective_user.id):
        await update.message.reply_text("❌ Нет прав")
        return ConversationHandler.END
    await update.message.reply_text("💳 Сумма взноса?")
    return CONTRIB_AMOUNT

async def contrib_amount_received(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        context.user_data["contrib_amount"] = float(update.message.text)
        await update.message.reply_text("👤 От кого?")
        return CONTRIB_WHO
    except:
        await update.message.reply_text("Введи число")
        return CONTRIB_AMOUNT

async def contrib_who_received(update: Update, context: ContextTypes.DEFAULT_TYPE):
    data = load_data()
    data["contributions"].append({
        "date": datetime.now().strftime("%d.%m.%Y"),
        "amount": context.user_data["contrib_amount"],
        "who": update.message.text
    })
    save_data(data)
    await update.message.reply_text("✅ Взнос добавлен!")
    return ConversationHandler.END

async def add_news_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_committee(update.effective_user.id):
        await update.message.reply_text("❌ Нет прав")
        return ConversationHandler.END
    await update.message.reply_text("Напиши новость:")
    return NEWS_TEXT

async def news_received(update: Update, context: ContextTypes.DEFAULT_TYPE):
    data = load_data()
    data["news"].append({
        "date": datetime.now().strftime("%d.%m.%Y"),
        "text": update.message.text,
        "author": update.effective_user.first_name or "Комитет"
    })
    save_data(data)
    await update.message.reply_text("✅ Новость добавлена!")
    return ConversationHandler.END

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
    data["events"].append({
        "name": context.user_data["event_name"],
        "date": context.user_data["event_date"],
        "description": update.message.text
    })
    save_data(data)
    await update.message.reply_text("✅ Событие добавлено!")
    return ConversationHandler.END

async def delete_expense(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_committee(update.effective_user.id):
        await update.message.reply_text("❌ Нет прав")
        return
    data = load_data()
    if not data["expenses"]:
        await update.message.reply_text("Расходов нет")
        return
    e = data["expenses"].pop()
    save_data(data)
    await update.message.reply_text(f"✅ Удалён: {e['amount']:,} ₽ ({e['description']})")

async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("❌ Отменено")
    return ConversationHandler.END

def main():
    if not BOT_TOKEN:
        raise RuntimeError("Нет BOT_TOKEN!")
    
    persistence = PicklePersistence(filepath=PERSISTENCE_FILE)
    app = Application.builder().token(BOT_TOKEN).persistence(persistence).build()
    
    expense_conv = ConversationHandler(
        entry_points=[CommandHandler("expense", expense_start)],
        states={
            AMOUNT: [MessageHandler(filters.TEXT & ~filters.COMMAND, amount_received)],
            CATEGORY: [MessageHandler(filters.TEXT & ~filters.COMMAND, category_received)],
            DESCRIPTION: [MessageHandler(filters.TEXT & ~filters.COMMAND, description_received)],
            WHO: [MessageHandler(filters.TEXT & ~filters.COMMAND, who_received)],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
        name="expense",
        persistent=True
    )
    
    contrib_conv = ConversationHandler(
        entry_points=[CommandHandler("contribution", contribution_start)],
        states={
            CONTRIB_AMOUNT: [MessageHandler(filters.TEXT & ~filters.COMMAND, contrib_amount_received)],
            CONTRIB_WHO: [MessageHandler(filters.TEXT & ~filters.COMMAND, contrib_who_received)],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
        name="contribution",
        persistent=True
    )
    
    news_conv = ConversationHandler(
        entry_points=[CommandHandler("addnews", add_news_start)],
        states={NEWS_TEXT: [MessageHandler(filters.TEXT & ~filters.COMMAND, news_received)]},
        fallbacks=[CommandHandler("cancel", cancel)],
        name="news",
        persistent=True
    )
    
    event_conv = ConversationHandler(
        entry_points=[CommandHandler("newevent", create_event_start)],
        states={
            EVENT_NAME: [MessageHandler(filters.TEXT & ~filters.COMMAND, event_name_received)],
            EVENT_DATE: [MessageHandler(filters.TEXT & ~filters.COMMAND, event_date_received)],
            EVENT_DESC: [MessageHandler(filters.TEXT & ~filters.COMMAND, event_desc_received)],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
        name="event",
        persistent=True
    )
    
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("report", report))
    app.add_handler(CommandHandler("history", history))
    app.add_handler(CommandHandler("export", export_expenses))
    app.add_handler(CommandHandler("news", show_news))
    app.add_handler(CommandHandler("birthdays", show_birthdays))
    app.add_handler(CommandHandler("fundraisers", show_fundraisers))
    app.add_handler(CommandHandler("events", show_events))
    app.add_handler(CommandHandler("deletexp", delete_expense))
    app.add_handler(expense_conv)
    app.add_handler(contrib_conv)
    app.add_handler(news_conv)
    app.add_handler(event_conv)
    
    logger.info("✅ БОТ ЗАПУЩЕН!")
    app.run_polling()

if __name__ == "__main__":
    main()
