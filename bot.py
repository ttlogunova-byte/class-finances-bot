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
    ConversationHandler, ContextTypes, PicklePersistence
)
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

BOT_TOKEN = os.environ.get("BOT_TOKEN")
BASE_FUND = 41000
DATA_FILE = "/data/data.json"
EXCEL_FILE = "/tmp/expenses.xlsx"
PERSISTENCE_FILE = "/data/conversation_state"
CATEGORIES = ["📚 Материалы", "🍕 Питание", "🎉 Мероприятие", "🏠 Организационные", "🚗 Доставка"]

# ID ЧЛЕНОВ КОМИТЕТА
COMMITTEE_IDS = [447774674, 6013055364]

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
    return user_id in COMMITTEE_IDS

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
    name = update.effective_user.first_name or "Друг"
    
    if is_committee(user_id):
        # СООБЩЕНИЕ ДЛЯ КОМИТЕТА
        msg = f"👋 Привет, {name}! Бот класса 1-К (КОМИТЕТ)\n"
        msg += f"🆔 **ID: {user_id}**\n\n"
        msg += "💰 ФИНАНСЫ:\n/report — финансы\n/history — расходы\n/export — скачать Excel\n\n"
        msg += "🎂 ДНИ РОЖДЕНИЯ:\n/birthdays — все ДР\n/january, /february, /march... /december\n\n"
        msg += "📰 НОВОСТИ:\n/news — все новости\n\n"
        msg += "🎯 СБОРЫ:\n/fundraisers — открытые сборы\n\n"
        msg += "📅 СОБЫТИЯ:\n/events — расписание\n\n"
        msg += "🛠️ ТОЛЬКО ДЛЯ КОМИТЕТА:\n"
        msg += "/expense — добавить расход\n"
        msg += "/addnews — добавить новость\n"
        msg += "/newfund — создать сбор\n"
        msg += "/newevent — добавить событие\n"
        msg += "/deletexp — удалить расход"
    else:
        # СООБЩЕНИЕ ДЛЯ РОДИТЕЛЕЙ
        msg = f"👋 Привет, {name}! Бот класса 1-К\n"
        msg += f"🆔 **ID: {user_id}**\n\n"
        msg += "💰 ФИНАНСЫ:\n/report — финансы\n/history — расходы\n/export — скачать Excel\n\n"
        msg += "🎂 ДНИ РОЖДЕНИЯ:\n/birthdays — все ДР\n/january, /february, /march... /december\n\n"
        msg += "📰 НОВОСТИ:\n/news — все новости класса\n\n"
        msg += "🎯 СБОРЫ:\n/fundraisers — открытые сборы\n\n"
        msg += "📅 СОБЫТИЯ:\n/events — расписание мероприятий"
    
    await update.message.reply_text(msg)

async def export_expenses(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Экспортирует расходы в Excel"""
    await update.message.reply_text("⏳ Создаю Excel файл...")
    
    data = load_data()
    try:
        file_path = create_excel_report(data)
        with open(file_path, 'rb') as f:
            await update.message.reply_document(
                document=f,
                filename=f"Расходы_1К_{datetime.now().strftime('%d.%m.%Y')}.xlsx",
                caption="📊 Отчёт по расходам класса 1-К"
            )
        os.remove(file_path)
        logger.info("✅ Файл отправлен")
    except Exception as e:
        logger.error(f"Ошибка экспорта: {str(e)}")
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
    
    # Показываем последний расход
    last_expense = data["expenses"][-1]
    msg = "⚠️ **ПОСЛЕДНИЙ РАСХОД:**\n\n"
    msg += f"📅 Дата: {last_expense['date']}\n"
    msg += f"💰 Сумма: {last_expense['amount']:,} ₽\n"
    msg += f"📌 Категория: {last_expense['category']}\n"
    msg += f"📝 Описание: {last_expense['description']}\n"
    msg += f"👤 От кого: {last_expense['who']}\n\n"
    msg += "❓ Ты уверен? Напиши **да** чтобы удалить или **нет** чтобы отменить"
    
    context.user_data["pending_delete"] = True
    await update.message.reply_text(msg)

async def confirm_delete(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_committee(update.effective_user.id):
        return
    
    if context.user_data.get("pending_delete"):
        if update.message.text.lower() == "да":
            data = load_data()
            if data["expenses"]:
                deleted = data["expenses"].pop()
                save_data(data)
                await update.message.reply_text(f"✅ УДАЛЕНО: {deleted['amount']:,} ₽ ({deleted['description']})")
            context.user_data["pending_delete"] = False
        elif update.message.text.lower() == "нет":
            await update.message.reply_text("❌ Отменено")
            context.user_data["pending_delete"] = False

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

# ФУНКЦИИ ДЛЯ КАЖДОГО МЕСЯЦА
async def january(update: Update, context: ContextTypes.DEFAULT_TYPE):
    bdays = BIRTHDAYS["Январь"]
    if not bdays:
        await update.message.reply_text("В январе дней рождения нет")
        return
    msg = "🎂 **ЯНВАРЬ:**\n"
    for name, date in bdays:
        msg += f"{date} — {name}\n"
    await update.message.reply_text(msg)

async def february(update: Update, context: ContextTypes.DEFAULT_TYPE):
    bdays = BIRTHDAYS["Февраль"]
    if not bdays:
        await update.message.reply_text("В феврале дней рождения нет")
        return
    msg = "🎂 **ФЕВРАЛЬ:**\n"
    for name, date in bdays:
        msg += f"{date} — {name}\n"
    await update.message.reply_text(msg)

async def march(update: Update, context: ContextTypes.DEFAULT_TYPE):
    bdays = BIRTHDAYS["Март"]
    if not bdays:
        await update.message.reply_text("В марте дней рождения нет")
        return
    msg = "🎂 **МАРТ:**\n"
    for name, date in bdays:
        msg += f"{date} — {name}\n"
    await update.message.reply_text(msg)

async def april(update: Update, context: ContextTypes.DEFAULT_TYPE):
    bdays = BIRTHDAYS["Апрель"]
    if not bdays:
        await update.message.reply_text("В апреле дней рождения нет")
        return
    msg = "🎂 **АПРЕЛЬ:**\n"
    for name, date in bdays:
        msg += f"{date} — {name}\n"
    await update.message.reply_text(msg)

async def may(update: Update, context: ContextTypes.DEFAULT_TYPE):
    bdays = BIRTHDAYS["Май"]
    if not bdays:
        await update.message.reply_text("В мае дней рождения нет")
        return
    msg = "🎂 **МАЙ:**\n"
    for name, date in bdays:
        msg += f"{date} — {name}\n"
    await update.message.reply_text(msg)

async def june(update: Update, context: ContextTypes.DEFAULT_TYPE):
    bdays = BIRTHDAYS["Июнь"]
    if not bdays:
        await update.message.reply_text("В июне дней рождения нет")
        return
    msg = "🎂 **ИЮНЬ:**\n"
    for name, date in bdays:
        msg += f"{date} — {name}\n"
    await update.message.reply_text(msg)

async def july(update: Update, context: ContextTypes.DEFAULT_TYPE):
    bdays = BIRTHDAYS["Июль"]
    if not bdays:
        await update.message.reply_text("В июле дней рождения нет")
        return
    msg = "🎂 **ИЮЛЬ:**\n"
    for name, date in bdays:
        msg += f"{date} — {name}\n"
    await update.message.reply_text(msg)

async def august(update: Update, context: ContextTypes.DEFAULT_TYPE):
    bdays = BIRTHDAYS["Август"]
    if not bdays:
        await update.message.reply_text("В августе дней рождения нет")
        return
    msg = "🎂 **АВГУСТ:**\n"
    for name, date in bdays:
        msg += f"{date} — {name}\n"
    await update.message.reply_text(msg)

async def september(update: Update, context: ContextTypes.DEFAULT_TYPE):
    bdays = BIRTHDAYS["Сентябрь"]
    if not bdays:
        await update.message.reply_text("В сентябре дней рождения нет")
        return
    msg = "🎂 **СЕНТЯБРЬ:**\n"
    for name, date in bdays:
        msg += f"{date} — {name}\n"
    await update.message.reply_text(msg)

async def october(update: Update, context: ContextTypes.DEFAULT_TYPE):
    bdays = BIRTHDAYS["Октябрь"]
    if not bdays:
        await update.message.reply_text("В октябре дней рождения нет")
        return
    msg = "🎂 **ОКТЯБРЬ:**\n"
    for name, date in bdays:
        msg += f"{date} — {name}\n"
    await update.message.reply_text(msg)

async def november(update: Update, context: ContextTypes.DEFAULT_TYPE):
    bdays = BIRTHDAYS["Ноябрь"]
    if not bdays:
        await update.message.reply_text("В ноябре дней рождения нет")
        return
    msg = "🎂 **НОЯБРЬ:**\n"
    for name, date in bdays:
        msg += f"{date} — {name}\n"
    await update.message.reply_text(msg)

async def december(update: Update, context: ContextTypes.DEFAULT_TYPE):
    bdays = BIRTHDAYS["Декабрь"]
    if not bdays:
        await update.message.reply_text("В декабре дней рождения нет")
        return
    msg = "🎂 **ДЕКАБРЬ:**\n"
    for name, date in bdays:
        msg += f"{date} — {name}\n"
    await update.message.reply_text(msg)

def main():
    if not BOT_TOKEN:
        raise RuntimeError("Нет BOT_TOKEN!")
    
    # Добавляем persistence для сохранения состояния
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
        name="expense_conv",
        persistent=True
    )
    
    news_conv = ConversationHandler(
        entry_points=[CommandHandler("addnews", add_news_start)],
        states={NEWS_TEXT: [MessageHandler(filters.TEXT & ~filters.COMMAND, news_text_received)]},
        fallbacks=[CommandHandler("cancel", cancel)],
        name="news_conv",
        persistent=True
    )
    
    fundraiser_conv = ConversationHandler(
        entry_points=[CommandHandler("newfund", create_fundraiser_start)],
        states={
            FUNDRAISER_NAME: [MessageHandler(filters.TEXT & ~filters.COMMAND, fundraiser_name_received)],
            FUNDRAISER_GOAL: [MessageHandler(filters.TEXT & ~filters.COMMAND, fundraiser_goal_received)],
            FUNDRAISER_DESC: [MessageHandler(filters.TEXT & ~filters.COMMAND, fundraiser_desc_received)],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
        name="fundraiser_conv",
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
        name="event_conv",
        persistent=True
    )
    
    # Команды
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("report", report))
    app.add_handler(CommandHandler("history", history))
    app.add_handler(CommandHandler("export", export_expenses))
    app.add_handler(CommandHandler("news", show_news))
    app.add_handler(CommandHandler("fundraisers", show_fundraisers))
    app.add_handler(CommandHandler("events", show_events))
    app.add_handler(CommandHandler("birthdays", birthdays_all))
    
    # ДР по месяцам
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
    
    app.add_handler(CommandHandler("deletexp", delete_expense))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, confirm_delete))
    app.add_handler(news_conv)
    app.add_handler(fundraiser_conv)
    app.add_handler(event_conv)
    app.add_handler(expense_conv)
    
    logger.info("✅ Бот класса 1-К запущен!")
    app.run_polling()

if __name__ == "__main__":
    main()
