import os
import json
import logging
from datetime import datetime

from telegram import Update, ReplyKeyboardMarkup
from telegram.ext import (
    Application,
    CommandHandler,
    MessageHandler,
    ConversationHandler,
    ContextTypes,
    filters,
)

BOT_TOKEN = os.environ.get("BOT_TOKEN")
if not BOT_TOKEN:
    raise RuntimeError("BOT_TOKEN не задан!")

BASE_FUND = 41000
DATA_FILE = "data.json"
COMMITTEE_IDS = [447774674, 6013055364]

CATEGORIES = [
    "Материалы",
    "Питание",
    "Мероприятие",
    "Организационные",
    "Доставка",
]

BIRTHDAYS = {
    "январь": [
        ("Гилёва Валерия", "11.01"),
        ("Хрусталев Матвей", "21.01"),
        ("Власов Тимофей", "23.01"),
    ],
    "февраль": [
        ("Булычев Елисей", "13.02"),
    ],
    "март": [
        ("Антипов Мирон", "07.03"),
        ("Волохова Анна", "07.03"),
        ("Богданов Илья", "11.03"),
        ("Братковский Гордей", "19.03"),
        ("Иваненко Алина", "19.03"),
        ("Ворошилов Вячеслав", "28.03"),
    ],
    "апрель": [
        ("Писаренко Николь", "03.04"),
    ],
    "май": [
        ("Хан Нелли", "28.05"),
    ],
    "июнь": [],
    "июль": [
        ("Лихачева София", "05.07"),
        ("Сокач Артём", "11.07"),
        ("Покровский Лев", "25.07"),
    ],
    "август": [
        ("Полуэктов Марк", "11.08"),
        ("Журавлёва Элина", "17.08"),
        ("Гаврилина Дарина", "24.08"),
        ("Селянин Владимир", "29.08"),
    ],
    "сентябрь": [],
    "октябрь": [
        ("Побережная Есения", "10.10"),
        ("Власюк Елизавета", "16.10"),
        ("Володина Мария", "20.10"),
        ("Боровикова Инна Игоревна (учитель)", "27.10"),
    ],
    "ноябрь": [
        ("Чаптыков Александр", "06.11"),
        ("Надтачеев Богдан", "19.11"),
        ("Зорина Екатерина", "23.11"),
        ("Петров Павел", "26.11"),
        ("Пиндур Вера", "27.11"),
    ],
    "декабрь": [
        ("Иванов Артём", "21.12"),
        ("Брюханова Милана", "25.12"),
        ("Мухачёв Михаил", "27.12"),
    ],
}

AMOUNT, CATEGORY, DESCRIPTION, WHO = range(4)
CONTRIB_AMOUNT, CONTRIB_WHO = range(4, 6)

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO,
)
logging.getLogger("httpx").setLevel(logging.WARNING)
logger = logging.getLogger(__name__)


def load_data():
    if os.path.exists(DATA_FILE):
        try:
            with open(DATA_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            data.setdefault("expenses", [])
            data.setdefault("contributions", [])
            return data
        except Exception as e:
            logger.error("Не удалось прочитать %s: %s", DATA_FILE, e)
    return {"expenses": [], "contributions": []}


def save_data(data):
    try:
        with open(DATA_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception as e:
        logger.error("Не удалось сохранить %s: %s", DATA_FILE, e)


def is_committee(user_id):
    return user_id in COMMITTEE_IDS


def main_keyboard(user_id):
    rows = []
    if is_committee(user_id):
        rows.append(["Расход", "Взнос"])
    rows.append(["Отчёт", "История"])
    rows.append(["Дни рождения", "Экспорт"])
    return ReplyKeyboardMarkup(rows, resize_keyboard=True)


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    text = (
        "Финансы класса 1-К\n\n"
        "Отчёт — остаток и итоги\n"
        "История — последние расходы\n"
        "Экспорт — файл Excel\n\n"
        "Дни рождения: напишите месяц, например: март"
    )
    await update.message.reply_text(text, reply_markup=main_keyboard(user_id))


async def report(update: Update, context: ContextTypes.DEFAULT_TYPE):
    data = load_data()
    spent = sum(e["amount"] for e in data["expenses"])
    got = sum(c["amount"] for c in data["contributions"])
    left = BASE_FUND + got - spent
    text = (
        "ФИНАНСОВЫЙ ОТЧЁТ\n\n"
        f"Базовый фонд: {BASE_FUND:.0f} руб.\n"
        f"Взносы: {got:.0f} руб.\n"
        f"Расходы: {spent:.0f} руб.\n"
        "--------------------\n"
        f"Остаток: {left:.0f} руб."
    )
    await update.message.reply_text(text)


async def history(update: Update, context: ContextTypes.DEFAULT_TYPE):
    data = load_data()
    if not data["expenses"]:
        await update.message.reply_text("Расходов пока нет.")
        return
    lines = ["ПОСЛЕДНИЕ РАСХОДЫ", ""]
    for e in data["expenses"][-15:]:
        lines.append(f"{e['date']} — {e['amount']:.0f} руб. — {e['category']}")
        if e.get("description"):
            lines.append(f"    {e['description']}")
    await update.message.reply_text("\n".join(lines))


async def birthdays_all(update: Update, context: ContextTypes.DEFAULT_TYPE):
    lines = ["ДНИ РОЖДЕНИЯ КЛАССА 1-К", ""]
    for month, people in BIRTHDAYS.items():
        if not people:
            continue
        lines.append(month.upper())
        for name, date in people:
            lines.append(f"  {date} — {name}")
        lines.append("")
    lines.append("Напишите месяц, чтобы посмотреть отдельно.")
    await update.message.reply_text("\n".join(lines))


async def birthdays_month(update: Update, context: ContextTypes.DEFAULT_TYPE, month):
    people = BIRTHDAYS.get(month, [])
    if not people:
        await update.message.reply_text(f"{month.capitalize()}: дней рождения нет.")
        return
    lines = [f"{month.upper()}", ""]
    for name, date in people:
        lines.append(f"{date} — {name}")
    await update.message.reply_text("\n".join(lines))


async def export_excel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    from openpyxl import Workbook
    from openpyxl.styles import Font

    data = load_data()
    spent = sum(e["amount"] for e in data["expenses"])
    got = sum(c["amount"] for c in data["contributions"])
    left = BASE_FUND + got - spent

    wb = Workbook()
    ws = wb.active
    ws.title = "Финансы"

    ws["A1"] = "ОТЧЁТ О ФИНАНСАХ КЛАССА 1-К"
    ws["A1"].font = Font(size=14, bold=True)
    ws["A2"] = f"Сформирован: {datetime.now().strftime('%d.%m.%Y')}"

    ws["A4"] = "Базовый фонд"
    ws["B4"] = BASE_FUND
    ws["A5"] = "Взносы"
    ws["B5"] = got
    ws["A6"] = "Расходы"
    ws["B6"] = spent
    ws["A7"] = "Остаток"
    ws["B7"] = left
    ws["A7"].font = Font(bold=True)
    ws["B7"].font = Font(bold=True)

    row = 9
    ws.cell(row=row, column=1).value = "РАСХОДЫ"
    ws.cell(row=row, column=1).font = Font(bold=True)
    row += 1
    for header, col in (("Дата", 1), ("Сумма", 2), ("Категория", 3), ("Описание", 4), ("Кто", 5)):
        ws.cell(row=row, column=col).value = header
        ws.cell(row=row, column=col).font = Font(bold=True)
    row += 1
    for e in data["expenses"]:
        ws.cell(row=row, column=1).value = e["date"]
        ws.cell(row=row, column=2).value = e["amount"]
        ws.cell(row=row, column=3).value = e["category"]
        ws.cell(row=row, column=4).value = e.get("description", "")
        ws.cell(row=row, column=5).value = e.get("who", "")
        row += 1

    row += 1
    ws.cell(row=row, column=1).value = "ВЗНОСЫ"
    ws.cell(row=row, column=1).font = Font(bold=True)
    row += 1
    for header, col in (("Дата", 1), ("Сумма", 2), ("От кого", 3)):
        ws.cell(row=row, column=col).value = header
        ws.cell(row=row, column=col).font = Font(bold=True)
    row += 1
    for c in data["contributions"]:
        ws.cell(row=row, column=1).value = c["date"]
        ws.cell(row=row, column=2).value = c["amount"]
        ws.cell(row=row, column=3).value = c.get("who", "")
        row += 1

    ws.column_dimensions["A"].width = 16
    ws.column_dimensions["B"].width = 12
    ws.column_dimensions["C"].width = 20
    ws.column_dimensions["D"].width = 35
    ws.column_dimensions["E"].width = 20

    filename = "finansy_1k.xlsx"
    wb.save(filename)
    with open(filename, "rb") as f:
        await update.message.reply_document(f, filename=filename)
    os.remove(filename)


async def expense_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_committee(update.effective_user.id):
        await update.message.reply_text("Добавлять расходы может только комитет.")
        return ConversationHandler.END
    await update.message.reply_text("Сумма расхода в рублях:")
    return AMOUNT


async def expense_amount(update: Update, context: ContextTypes.DEFAULT_TYPE):
    raw = update.message.text.replace(",", ".").strip()
    try:
        context.user_data["amount"] = float(raw)
    except ValueError:
        await update.message.reply_text("Введите число, например: 1500")
        return AMOUNT
    keyboard = ReplyKeyboardMarkup([[c] for c in CATEGORIES], resize_keyboard=True)
    await update.message.reply_text("Категория:", reply_markup=keyboard)
    return CATEGORY


async def expense_category(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["category"] = update.message.text
    await update.message.reply_text("Описание (на что потратили):")
    return DESCRIPTION


async def expense_description(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["description"] = update.message.text
    await update.message.reply_text("Кто оплатил:")
    return WHO


async def expense_who(update: Update, context: ContextTypes.DEFAULT_TYPE):
    data = load_data()
    data["expenses"].append(
        {
            "date": datetime.now().strftime("%d.%m.%Y"),
            "amount": context.user_data["amount"],
            "category": context.user_data["category"],
            "description": context.user_data["description"],
            "who": update.message.text,
        }
    )
    save_data(data)
    context.user_data.clear()
    await update.message.reply_text(
        "Расход добавлен.", reply_markup=main_keyboard(update.effective_user.id)
    )
    return ConversationHandler.END


async def contrib_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_committee(update.effective_user.id):
        await update.message.reply_text("Добавлять взносы может только комитет.")
        return ConversationHandler.END
    await update.message.reply_text("Сумма взноса в рублях:")
    return CONTRIB_AMOUNT


async def contrib_amount(update: Update, context: ContextTypes.DEFAULT_TYPE):
    raw = update.message.text.replace(",", ".").strip()
    try:
        context.user_data["contrib_amount"] = float(raw)
    except ValueError:
        await update.message.reply_text("Введите число, например: 500")
        return CONTRIB_AMOUNT
    await update.message.reply_text("От кого взнос:")
    return CONTRIB_WHO


async def contrib_who(update: Update, context: ContextTypes.DEFAULT_TYPE):
    data = load_data()
    data["contributions"].append(
        {
            "date": datetime.now().strftime("%d.%m.%Y"),
            "amount": context.user_data["contrib_amount"],
            "who": update.message.text,
        }
    )
    save_data(data)
    context.user_data.clear()
    await update.message.reply_text(
        "Взнос добавлен.", reply_markup=main_keyboard(update.effective_user.id)
    )
    return ConversationHandler.END


async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.clear()
    await update.message.reply_text(
        "Отменено.", reply_markup=main_keyboard(update.effective_user.id)
    )
    return ConversationHandler.END


async def handle_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = (update.message.text or "").strip().lower()

    if text in BIRTHDAYS:
        await birthdays_month(update, context, text)
        return
    if text in ("отчёт", "отчет"):
        await report(update, context)
        return
    if text == "история":
        await history(update, context)
        return
    if text == "экспорт":
        await export_excel(update, context)
        return
    if text in ("дни рождения", "дни рождений"):
        await birthdays_all(update, context)
        return
    if text == "расход":
        await expense_start(update, context)
        return
    if text == "взнос":
        await contrib_start(update, context)
        return

    await update.message.reply_text(
        "Не понял. Нажмите кнопку или напишите месяц, например: март",
        reply_markup=main_keyboard(update.effective_user.id),
    )


async def on_error(update, context):
    logger.error("Ошибка при обработке обновления", exc_info=context.error)


def main():
    app = Application.builder().token(BOT_TOKEN).build()

    expense_conv = ConversationHandler(
        entry_points=[
            CommandHandler("expense", expense_start),
            MessageHandler(filters.Regex(r"^Расход$"), expense_start),
        ],
        states={
            AMOUNT: [MessageHandler(filters.TEXT & ~filters.COMMAND, expense_amount)],
            CATEGORY: [MessageHandler(filters.TEXT & ~filters.COMMAND, expense_category)],
            DESCRIPTION: [MessageHandler(filters.TEXT & ~filters.COMMAND, expense_description)],
            WHO: [MessageHandler(filters.TEXT & ~filters.COMMAND, expense_who)],
        },
        fallbacks=[CommandHandler("cancel", cancel), CommandHandler("start", cancel)],
    )

    contrib_conv = ConversationHandler(
        entry_points=[
            CommandHandler("contribution", contrib_start),
            MessageHandler(filters.Regex(r"^Взнос$"), contrib_start),
        ],
        states={
            CONTRIB_AMOUNT: [MessageHandler(filters.TEXT & ~filters.COMMAND, contrib_amount)],
            CONTRIB_WHO: [MessageHandler(filters.TEXT & ~filters.COMMAND, contrib_who)],
        },
        fallbacks=[CommandHandler("cancel", cancel), CommandHandler("start", cancel)],
    )

    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("help", start))
    app.add_handler(CommandHandler("report", report))
    app.add_handler(CommandHandler("history", history))
    app.add_handler(CommandHandler("export", export_excel))
    app.add_handler(CommandHandler("birthdays", birthdays_all))
    app.add_handler(expense_conv)
    app.add_handler(contrib_conv)
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text))
    app.add_error_handler(on_error)

    logger.info("Бот запущен")
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
