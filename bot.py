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

# Родительский комитет класса 1-К
COMMITTEE_IDS = [447774674, 6013055364]

CATEGORIES = [
    "📚 Материалы",
    "🍕 Питание",
    "🎉 Мероприятие",
    "🏠 Организационные",
    "🚗 Доставка",
    "🎁 Подарки",
    "📦 Прочее",
]

MONTHS = [
    "январь",
    "февраль",
    "март",
    "апрель",
    "май",
    "июнь",
    "июль",
    "август",
    "сентябрь",
    "октябрь",
    "ноябрь",
    "декабрь",
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
    format="%(asctime)s - %(levelname)s - %(message)s",
    level=logging.INFO,
)
logging.getLogger("httpx").setLevel(logging.WARNING)
logger = logging.getLogger(__name__)


# ---------- данные ----------

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


def money(value):
    return f"{value:,.0f}".replace(",", " ")


def totals(data):
    spent = sum(e["amount"] for e in data["expenses"])
    got = sum(c["amount"] for c in data["contributions"])
    return spent, got, BASE_FUND + got - spent


def normalize(text):
    """Убирает эмодзи и лишние пробелы, приводит к нижнему регистру."""
    cleaned = "".join(ch for ch in text if ch.isalpha() or ch.isspace() or ch == "-")
    return " ".join(cleaned.split()).lower()


# ---------- клавиатуры ----------

def main_keyboard(user_id):
    rows = []
    if is_committee(user_id):
        rows.append(["➕ Расход", "➕ Взнос"])
    rows.append(["📊 Отчёт", "📈 История"])
    rows.append(["🎂 Дни рождения", "📥 Экспорт"])
    if is_committee(user_id):
        rows.append(["↩️ Отменить последний расход"])
    return ReplyKeyboardMarkup(rows, resize_keyboard=True)


def months_keyboard():
    rows = []
    for i in range(0, 12, 3):
        rows.append([m.capitalize() for m in MONTHS[i : i + 3]])
    rows.append(["⬅️ Назад"])
    return ReplyKeyboardMarkup(rows, resize_keyboard=True)


# ---------- команды ----------

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    if is_committee(user.id):
        role = "Вы в родительском комитете — можете вносить расходы и взносы."
    else:
        role = "Режим просмотра: отчёты, история, дни рождения и выгрузка в Excel."

    text = (
        "🏫 ФИНАНСЫ КЛАССА 1-К\n"
        "МБОУ «СШ №23», Красноярск\n\n"
        f"{role}\n\n"
        "📊 Отчёт — остаток и траты по категориям\n"
        "📈 История — последние расходы\n"
        "🎂 Дни рождения — по месяцам\n"
        "📥 Экспорт — файл Excel\n\n"
        "Чтобы посмотреть именинников, можно просто написать месяц, например: март"
    )
    await update.message.reply_text(text, reply_markup=main_keyboard(user.id))


async def report(update: Update, context: ContextTypes.DEFAULT_TYPE):
    data = load_data()
    spent, got, left = totals(data)

    lines = [
        "📊 ФИНАНСОВЫЙ ОТЧЁТ",
        "",
        f"💰 Базовый фонд: {money(BASE_FUND)} ₽",
        f"➕ Взносы: {money(got)} ₽",
        f"➖ Расходы: {money(spent)} ₽",
        "━━━━━━━━━━━━━━━━",
        f"✅ Остаток: {money(left)} ₽",
    ]

    if data["expenses"]:
        by_cat = {}
        for e in data["expenses"]:
            by_cat[e["category"]] = by_cat.get(e["category"], 0) + e["amount"]
        lines.append("")
        lines.append("📂 РАСХОДЫ ПО КАТЕГОРИЯМ")
        lines.append("")
        for cat, total in sorted(by_cat.items(), key=lambda x: -x[1]):
            share = total / spent * 100 if spent else 0
            lines.append(f"{cat}")
            lines.append(f"   {money(total)} ₽ · {share:.0f}%")
        lines.append("")
        lines.append(f"Всего операций: {len(data['expenses'])}")

    await update.message.reply_text("\n".join(lines))


async def history(update: Update, context: ContextTypes.DEFAULT_TYPE):
    data = load_data()
    if not data["expenses"]:
        await update.message.reply_text("📭 Расходов пока нет.")
        return

    lines = ["📈 ПОСЛЕДНИЕ РАСХОДЫ", ""]
    for e in data["expenses"][-15:][::-1]:
        lines.append(f"📅 {e['date']} · {money(e['amount'])} ₽")
        lines.append(f"   {e['category']}")
        if e.get("description"):
            lines.append(f"   {e['description']}")
        if e.get("who"):
            lines.append(f"   оплатил(а): {e['who']}")
        lines.append("")

    if data["contributions"]:
        lines.append("💵 ПОСЛЕДНИЕ ВЗНОСЫ")
        lines.append("")
        for c in data["contributions"][-10:][::-1]:
            lines.append(f"📅 {c['date']} · {money(c['amount'])} ₽ · {c.get('who', '')}")

    await update.message.reply_text("\n".join(lines))


async def birthdays_menu(update: Update, context: ContextTypes.DEFAULT_TYPE):
    total = sum(len(v) for v in BIRTHDAYS.values())
    text = (
        "🎂 ДНИ РОЖДЕНИЯ КЛАССА 1-К\n\n"
        f"Всего в списке: {total}\n\n"
        "Выберите месяц на клавиатуре или напишите его названием."
    )
    await update.message.reply_text(text, reply_markup=months_keyboard())


async def birthdays_month(update: Update, context: ContextTypes.DEFAULT_TYPE, month):
    people = BIRTHDAYS.get(month, [])
    user_id = update.effective_user.id

    if not people:
        await update.message.reply_text(
            f"🎂 {month.capitalize()}\n\nВ этом месяце именинников нет.",
            reply_markup=months_keyboard(),
        )
        return

    lines = [f"🎂 {month.upper()}", ""]
    for name, date in people:
        lines.append(f"🎈 {date} — {name}")
    lines.append("")
    lines.append(f"Всего: {len(people)}")

    await update.message.reply_text("\n".join(lines), reply_markup=months_keyboard())


async def birthdays_all(update: Update, context: ContextTypes.DEFAULT_TYPE):
    lines = ["🎂 ВСЕ ДНИ РОЖДЕНИЯ КЛАССА 1-К", ""]
    for month in MONTHS:
        people = BIRTHDAYS[month]
        if not people:
            continue
        lines.append(f"── {month.upper()} ──")
        for name, date in people:
            lines.append(f"🎈 {date} — {name}")
        lines.append("")
    await update.message.reply_text("\n".join(lines))


async def export_excel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment

    data = load_data()
    spent, got, left = totals(data)

    header_fill = PatternFill("solid", fgColor="D9E2F3")
    bold = Font(bold=True)
    title_font = Font(size=14, bold=True)

    wb = Workbook()

    # Лист 1 — сводка
    ws = wb.active
    ws.title = "Сводка"
    ws["A1"] = "ОТЧЁТ О ФИНАНСАХ КЛАССА 1-К"
    ws["A1"].font = title_font
    ws["A2"] = "МБОУ «СШ №23», Красноярск"
    ws["A3"] = f"Сформирован: {datetime.now().strftime('%d.%m.%Y')}"

    ws["A5"] = "Базовый фонд"
    ws["B5"] = BASE_FUND
    ws["A6"] = "Взносы"
    ws["B6"] = got
    ws["A7"] = "Расходы"
    ws["B7"] = spent
    ws["A8"] = "Остаток"
    ws["B8"] = left
    ws["A8"].font = bold
    ws["B8"].font = bold

    if data["expenses"]:
        by_cat = {}
        for e in data["expenses"]:
            by_cat[e["category"]] = by_cat.get(e["category"], 0) + e["amount"]
        ws["A10"] = "РАСХОДЫ ПО КАТЕГОРИЯМ"
        ws["A10"].font = bold
        ws["A11"] = "Категория"
        ws["B11"] = "Сумма"
        ws["C11"] = "Доля"
        for col in ("A11", "B11", "C11"):
            ws[col].font = bold
            ws[col].fill = header_fill
        row = 12
        for cat, total in sorted(by_cat.items(), key=lambda x: -x[1]):
            ws.cell(row=row, column=1).value = cat
            ws.cell(row=row, column=2).value = total
            ws.cell(row=row, column=3).value = total / spent if spent else 0
            ws.cell(row=row, column=3).number_format = "0%"
            row += 1

    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 10

    # Лист 2 — расходы
    ws2 = wb.create_sheet("Расходы")
    headers = ["Дата", "Сумма", "Категория", "Описание", "Кто оплатил"]
    for i, h in enumerate(headers, start=1):
        c = ws2.cell(row=1, column=i)
        c.value = h
        c.font = bold
        c.fill = header_fill
    row = 2
    for e in data["expenses"]:
        ws2.cell(row=row, column=1).value = e["date"]
        ws2.cell(row=row, column=2).value = e["amount"]
        ws2.cell(row=row, column=3).value = e["category"]
        ws2.cell(row=row, column=4).value = e.get("description", "")
        ws2.cell(row=row, column=5).value = e.get("who", "")
        row += 1
    for col, w in zip("ABCDE", (14, 12, 22, 40, 22)):
        ws2.column_dimensions[col].width = w

    # Лист 3 — взносы
    ws3 = wb.create_sheet("Взносы")
    for i, h in enumerate(["Дата", "Сумма", "От кого"], start=1):
        c = ws3.cell(row=1, column=i)
        c.value = h
        c.font = bold
        c.fill = header_fill
    row = 2
    for c_ in data["contributions"]:
        ws3.cell(row=row, column=1).value = c_["date"]
        ws3.cell(row=row, column=2).value = c_["amount"]
        ws3.cell(row=row, column=3).value = c_.get("who", "")
        row += 1
    for col, w in zip("ABC", (14, 12, 28)):
        ws3.column_dimensions[col].width = w

    # Лист 4 — дни рождения
    ws4 = wb.create_sheet("Дни рождения")
    for i, h in enumerate(["Месяц", "Дата", "Имя"], start=1):
        c = ws4.cell(row=1, column=i)
        c.value = h
        c.font = bold
        c.fill = header_fill
    row = 2
    for month in MONTHS:
        for name, date in BIRTHDAYS[month]:
            ws4.cell(row=row, column=1).value = month.capitalize()
            ws4.cell(row=row, column=2).value = date
            ws4.cell(row=row, column=3).value = name
            row += 1
    for col, w in zip("ABC", (14, 10, 38)):
        ws4.column_dimensions[col].width = w

    filename = f"finansy_1k_{datetime.now().strftime('%d%m%Y')}.xlsx"
    wb.save(filename)
    with open(filename, "rb") as f:
        await update.message.reply_document(f, filename=filename)
    os.remove(filename)


async def undo_last(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_committee(update.effective_user.id):
        await update.message.reply_text("🔒 Это может сделать только комитет.")
        return
    data = load_data()
    if not data["expenses"]:
        await update.message.reply_text("📭 Нечего отменять — расходов нет.")
        return
    removed = data["expenses"].pop()
    save_data(data)
    await update.message.reply_text(
        "↩️ Расход удалён:\n\n"
        f"📅 {removed['date']} · {money(removed['amount'])} ₽\n"
        f"{removed['category']}\n"
        f"{removed.get('description', '')}",
        reply_markup=main_keyboard(update.effective_user.id),
    )


# ---------- диалог: расход ----------

async def expense_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_committee(update.effective_user.id):
        await update.message.reply_text(
            "🔒 Добавлять расходы может только родительский комитет."
        )
        return ConversationHandler.END
    await update.message.reply_text("💵 Введите сумму расхода в рублях:")
    return AMOUNT


async def expense_amount(update: Update, context: ContextTypes.DEFAULT_TYPE):
    raw = update.message.text.replace(",", ".").replace(" ", "").strip()
    try:
        value = float(raw)
        if value <= 0:
            raise ValueError
        context.user_data["amount"] = value
    except ValueError:
        await update.message.reply_text("❌ Нужно число больше нуля. Например: 1500")
        return AMOUNT
    keyboard = ReplyKeyboardMarkup([[c] for c in CATEGORIES], resize_keyboard=True)
    await update.message.reply_text("📂 Выберите категорию:", reply_markup=keyboard)
    return CATEGORY


async def expense_category(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["category"] = update.message.text
    await update.message.reply_text("📝 На что потратили:")
    return DESCRIPTION


async def expense_description(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["description"] = update.message.text
    await update.message.reply_text("👤 Кто оплатил:")
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
    spent, got, left = totals(data)
    amount = context.user_data["amount"]
    category = context.user_data["category"]
    context.user_data.clear()

    await update.message.reply_text(
        "✅ Расход добавлен\n\n"
        f"💵 {money(amount)} ₽\n"
        f"{category}\n\n"
        f"Остаток: {money(left)} ₽",
        reply_markup=main_keyboard(update.effective_user.id),
    )
    return ConversationHandler.END


# ---------- диалог: взнос ----------

async def contrib_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_committee(update.effective_user.id):
        await update.message.reply_text(
            "🔒 Добавлять взносы может только родительский комитет."
        )
        return ConversationHandler.END
    await update.message.reply_text("💵 Введите сумму взноса в рублях:")
    return CONTRIB_AMOUNT


async def contrib_amount(update: Update, context: ContextTypes.DEFAULT_TYPE):
    raw = update.message.text.replace(",", ".").replace(" ", "").strip()
    try:
        value = float(raw)
        if value <= 0:
            raise ValueError
        context.user_data["contrib_amount"] = value
    except ValueError:
        await update.message.reply_text("❌ Нужно число больше нуля. Например: 500")
        return CONTRIB_AMOUNT
    await update.message.reply_text("👤 От кого взнос:")
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
    spent, got, left = totals(data)
    amount = context.user_data["contrib_amount"]
    context.user_data.clear()

    await update.message.reply_text(
        "✅ Взнос добавлен\n\n"
        f"💵 {money(amount)} ₽\n"
        f"от {update.message.text}\n\n"
        f"Остаток: {money(left)} ₽",
        reply_markup=main_keyboard(update.effective_user.id),
    )
    return ConversationHandler.END


async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.clear()
    await update.message.reply_text(
        "Отменено.", reply_markup=main_keyboard(update.effective_user.id)
    )
    return ConversationHandler.END


# ---------- свободный текст ----------

async def handle_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    raw = (update.message.text or "").strip()
    text = normalize(raw)
    user_id = update.effective_user.id

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
    if text in ("дни рождения", "дни рождений", "днирождения"):
        await birthdays_menu(update, context)
        return
    if text in ("все дни рождения", "весь список"):
        await birthdays_all(update, context)
        return
    if text == "расход":
        await expense_start(update, context)
        return
    if text == "взнос":
        await contrib_start(update, context)
        return
    if text in ("отменить последний расход", "отменить"):
        await undo_last(update, context)
        return
    if text in ("назад", "меню"):
        await update.message.reply_text(
            "Главное меню", reply_markup=main_keyboard(user_id)
        )
        return

    await update.message.reply_text(
        "Не понял команду. Нажмите кнопку внизу или напишите месяц, например: март",
        reply_markup=main_keyboard(user_id),
    )


async def on_error(update, context):
    logger.error("Ошибка при обработке обновления", exc_info=context.error)


# ---------- запуск ----------

def main():
    app = Application.builder().token(BOT_TOKEN).build()

    expense_conv = ConversationHandler(
        entry_points=[
            CommandHandler("expense", expense_start),
            MessageHandler(filters.Regex(r"^➕ Расход$"), expense_start),
        ],
        states={
            AMOUNT: [MessageHandler(filters.TEXT & ~filters.COMMAND, expense_amount)],
            CATEGORY: [MessageHandler(filters.TEXT & ~filters.COMMAND, expense_category)],
            DESCRIPTION: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, expense_description)
            ],
            WHO: [MessageHandler(filters.TEXT & ~filters.COMMAND, expense_who)],
        },
        fallbacks=[CommandHandler("cancel", cancel), CommandHandler("start", cancel)],
    )

    contrib_conv = ConversationHandler(
        entry_points=[
            CommandHandler("contribution", contrib_start),
            MessageHandler(filters.Regex(r"^➕ Взнос$"), contrib_start),
        ],
        states={
            CONTRIB_AMOUNT: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, contrib_amount)
            ],
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
    app.add_handler(CommandHandler("undo", undo_last))
    app.add_handler(expense_conv)
    app.add_handler(contrib_conv)
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text))
    app.add_error_handler(on_error)

    logger.info("Бот запущен")
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
