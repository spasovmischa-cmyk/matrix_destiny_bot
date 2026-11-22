import logging
import os
import re
from datetime import datetime
from pathlib import Path

from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    Application, CallbackQueryHandler, CommandHandler, ContextTypes,
    ConversationHandler, MessageHandler, filters,
)

# Set BOT_TOKEN in Railway Variables. Never put a real token in GitHub/code.
BOT_TOKEN = os.environ.get("BOT_TOKEN")
if not BOT_TOKEN:
    raise RuntimeError("BOT_TOKEN is not configured in Railway Variables")

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

(NAME, DOB, METHOD, EXTRA, TIME, PLACE, FAMILY, CONFIRM) = range(8)
METHODS = {
    "m_pyth": "Пифагорейская",
    "m_kabb": "Каббалистическая",
    "m_vedic": "Ведическая",
}


def keyboard(rows):
    return InlineKeyboardMarkup([
        [InlineKeyboardButton(label, callback_data=data) for label, data in row]
        for row in rows
    ])


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    context.user_data.clear()
    first = update.effective_user.first_name or "друг"
    await update.message.reply_text(
        f"🌟 Добро пожаловать, {first}! Я помогу подготовить персональный нумерологический разбор.\n\n"
        "Это символическая система для саморефлексии, а не научный прогноз.\n\n"
        "Для начала введите полное ФИО: имя, отчество и фамилию (три слова).",
        reply_markup=keyboard([[('ℹ️ Помощь', 'help')]])
    )
    return NAME


async def help_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    await update.message.reply_text(
        "Как пройти разбор:\n1) Введите ФИО из трёх слов.\n2) Дату рождения в формате ДД.ММ.ГГГГ.\n"
        "3) Выберите нумерологическую традицию.\n4) При желании добавьте время/место рождения и данные близких.\n"
        "5) Проверьте сведения и подтвердите создание DOCX.\n\n"
        "Команды: /start — начать заново; /cancel — отменить."
    )


async def help_button(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    q = update.callback_query
    await q.answer()
    await q.message.reply_text(
        "Введите полное ФИО (имя, отчество и фамилию — три слова). Дополнительные сведения можно пропустить."
    )
    return NAME


async def name_input(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    raw = " ".join(update.message.text.strip().split())
    parts = raw.split()
    if len(parts) < 3 or not all(re.fullmatch(r"[A-Za-zА-Яа-яЁё’'-]+", p) for p in parts):
        await update.message.reply_text(
            "Пожалуйста, укажите ФИО тремя словами: имя, отчество, фамилия.\nНапример: Иван Иванович Иванов."
        )
        return NAME
    context.user_data['full_name'] = raw
    context.user_data['first_name'] = parts[0]
    await update.message.reply_text(f"Спасибо, {parts[0]}! Введите дату рождения в формате ДД.ММ.ГГГГ.")
    return DOB


async def dob_input(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    try:
        dob = datetime.strptime(update.message.text.strip(), "%d.%m.%Y")
        if dob.date() > datetime.now().date():
            raise ValueError
    except ValueError:
        await update.message.reply_text("Не удалось распознать дату. Введите её как ДД.ММ.ГГГГ, например 15.03.1990.")
        return DOB
    context.user_data['dob'] = dob
    context.user_data['dob_str'] = dob.strftime("%d.%m.%Y")
    await update.message.reply_text("Выберите методику:", reply_markup=keyboard([
        [('Пифагорейская', 'm_pyth')], [('Каббалистическая', 'm_kabb')], [('Ведическая', 'm_vedic')]
    ]))
    return METHOD


async def method_choice(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    q = update.callback_query
    await q.answer()
    context.user_data['method'] = METHODS[q.data]
    await q.edit_message_text(
        f"Методика: {METHODS[q.data]}. Хотите добавить дополнительные данные?",
        reply_markup=keyboard([[('Да, добавить', 'extra_yes'), ('Пропустить', 'extra_no')]])
    )
    return EXTRA


async def extra_choice(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    q = update.callback_query
    await q.answer()
    if q.data == 'extra_no':
        context.user_data['birth_time'] = 'не указано'
        context.user_data['birth_place'] = 'не указано'
        context.user_data['family'] = ''
        await show_review(q, context)
        return CONFIRM
    await q.edit_message_text(
        "Знаете время рождения?",
        reply_markup=keyboard([[('Введу время', 'time_yes'), ('Не знаю / пропустить', 'time_no')]])
    )
    return TIME


async def time_choice(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    q = update.callback_query
    await q.answer()
    if q.data == 'time_no':
        context.user_data['birth_time'] = 'не указано'
        await q.edit_message_text(
            "Знаете место рождения?",
            reply_markup=keyboard([[('Введу город и страну', 'place_yes'), ('Пропустить', 'place_no')]])
        )
        return PLACE
    await q.edit_message_text("Введите время рождения в формате ЧЧ:ММ (например, 14:30).")
    return TIME


async def time_input(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    text = update.message.text.strip()
    try:
        parsed = datetime.strptime(text, "%H:%M")
        context.user_data['birth_time'] = parsed.strftime("%H:%M")
    except ValueError:
        await update.message.reply_text("Неверный формат. Введите время как ЧЧ:ММ, например 08:45, или /cancel.")
        return TIME
    await update.message.reply_text(
        "Знаете место рождения?",
        reply_markup=keyboard([[('Введу город и страну', 'place_yes'), ('Пропустить', 'place_no')]])
    )
    return PLACE


async def place_choice(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    q = update.callback_query
    await q.answer()
    if q.data == 'place_no':
        context.user_data['birth_place'] = 'не указано'
        await ask_family(q)
        return FAMILY
    await q.edit_message_text("Введите город и страну рождения (например: Казань, Россия).")
    return PLACE


async def place_input(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    place = update.message.text.strip()
    if len(place) < 2 or len(place) > 100:
        await update.message.reply_text("Укажите город и страну коротко (до 100 символов) или нажмите /cancel.")
        return PLACE
    context.user_data['birth_place'] = place
    await update.message.reply_text(
        "Добавить сведения о партнёре, детях или родителях?",
        reply_markup=keyboard([[('Добавить сведения', 'family_yes'), ('Пропустить', 'family_no')]])
    )
    return FAMILY


async def ask_family(q):
    await q.edit_message_text(
        "Добавить сведения о партнёре, детях или родителях?",
        reply_markup=keyboard([[('Добавить сведения', 'family_yes'), ('Пропустить', 'family_no')]])
    )


async def family_choice(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    q = update.callback_query
    await q.answer()
    if q.data == 'family_no':
        context.user_data['family'] = ''
        await show_review(q, context)
        return CONFIRM
    await q.edit_message_text(
        "Напишите только те сведения, которые хотите включить: например, «партнёр — Анна, 12.04.1991; сын — …». "
        "Не отправляйте чувствительные данные, которые не нужны для разбора."
    )
    return FAMILY


async def family_input(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    context.user_data['family'] = update.message.text.strip()[:1500]
    await show_review(update, context)
    return CONFIRM


async def show_review(update_or_query, context: ContextTypes.DEFAULT_TYPE) -> None:
    d = context.user_data
    msg = (
        "Проверьте данные перед созданием документа:\n\n"
        f"ФИО: {d.get('full_name')}\nДата рождения: {d.get('dob_str')}\n"
        f"Методика: {d.get('method')}\nВремя рождения: {d.get('birth_time', 'не указано')}\n"
        f"Место рождения: {d.get('birth_place', 'не указано')}\n"
        f"Сведения о близких: {'добавлены' if d.get('family') else 'не добавлены'}\n\nВсё верно?"
    )
    markup = keyboard([[('✅ Да, создать документ', 'confirm_yes')], [('↩️ Начать заново', 'confirm_restart')]])
    if hasattr(update_or_query, 'callback_query') and update_or_query.callback_query:
        await update_or_query.callback_query.message.reply_text(msg, reply_markup=markup)
    elif hasattr(update_or_query, 'message') and update_or_query.message:
        await update_or_query.message.reply_text(msg, reply_markup=markup)
    else:
        await update_or_query.message.reply_text(msg, reply_markup=markup)


async def confirm(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    q = update.callback_query
    await q.answer()
    if q.data == 'confirm_restart':
        context.user_data.clear()
        await q.edit_message_text("Хорошо, начнём заново. Введите полное ФИО (три слова).")
        return NAME
    await q.edit_message_text("Готовлю документ. Обычно это занимает несколько секунд…")
    try:
        path = create_doc(context.user_data)
        with open(path, 'rb') as f:
            await context.bot.send_document(
                chat_id=q.from_user.id,
                document=f,
                filename=Path(path).name,
                caption=f"Готово, {context.user_data['first_name']}! Ваш документ прикреплён."
            )
        os.remove(path)
        await context.bot.send_message(q.from_user.id, "Чтобы подготовить новый разбор, отправьте /start.")
    except Exception:
        logger.exception("Document generation/delivery failed")
        await context.bot.send_message(q.from_user.id, "Не удалось подготовить или отправить файл. Попробуйте /start позже.")
    context.user_data.clear()
    return ConversationHandler.END


NUMBERS = {
    1: ('Инициатива и самостоятельность', 'Сильная сторона — способность начинать и брать ответственность. Полезно сочетать решительность с умением слушать и распределять задачи.'),
    2: ('Сотрудничество и восприимчивость', 'Сильная сторона — тонкое внимание к отношениям и контексту. Важно сохранять собственные границы и не подменять согласие избеганием конфликта.'),
    3: ('Выражение и творчество', 'Сильная сторона — коммуникация, воображение и способность оживлять идеи. Результат поддерживают фокус, регулярная практика и завершение начатого.'),
    4: ('Структура и последовательность', 'Сильная сторона — надёжность и умение выстраивать процессы. Гибкость помогает не превращать устойчивость в сопротивление переменам.'),
    5: ('Изменения и адаптивность', 'Сильная сторона — любознательность и способность быстро ориентироваться. Полезно заранее оценивать риски и оставлять достаточно времени для доведения дел до конца.'),
    6: ('Забота и ответственность', 'Сильная сторона — стремление поддерживать людей и создавать устойчивую среду. Здоровые границы позволяют помогать без истощения и чрезмерного контроля.'),
    7: ('Исследование и рефлексия', 'Сильная сторона — наблюдательность, анализ и интерес к смыслу. Важно переводить размышления в проверяемые действия и поддерживать контакт с людьми.'),
    8: ('Организация и управление ресурсами', 'Сильная сторона — стратегическое мышление и ориентация на результат. Устойчивый успех требует этичности, реалистичной оценки риска и баланса с личной жизнью.'),
    9: ('Широкий взгляд и завершение', 'Сильная сторона — способность видеть общую картину и учитывать интересы других. Полезно сочетать идеалы с конкретными границами и практическими шагами.'),
}


def reduce_num(n: int) -> int:
    while n > 9:
        n = sum(int(x) for x in str(n))
    return n


def name_number(name: str) -> int:
    alphabet = 'АБВГДЕЁЖЗИЙКЛМНОПРСТУФХЦЧШЩЪЫЬЭЮЯ'
    values = {ch: i % 9 + 1 for i, ch in enumerate(alphabet)}
    return reduce_num(sum(values.get(ch, 0) for ch in name.upper()))


def add_section(doc, title, paragraphs):
    doc.add_heading(title, level=1)
    for text in paragraphs:
        doc.add_paragraph(text)


def create_doc(d: dict) -> str:
    full = d['full_name']
    first = d['first_name']
    dob = d['dob']
    life = reduce_num(sum(int(c) for c in dob.strftime('%d%m%Y')))
    expr = name_number(full)
    trait, description = NUMBERS[life]
    name_trait, name_description = NUMBERS[expr]
    doc = Document()
    normal = doc.styles['Normal']
    normal.font.name = 'Arial'
    normal.font.size = Pt(10.5)
    normal.paragraph_format.space_after = Pt(7)
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run('ПЕРСОНАЛЬНЫЙ НУМЕРОЛОГИЧЕСКИЙ РАЗБОР')
    run.bold = True
    run.font.size = Pt(18)
    run.font.color.rgb = RGBColor(40, 70, 110)
    doc.add_paragraph(f"ФИО: {full}\nДата рождения: {d['dob_str']}\nМетодика: {d['method']}\nДата подготовки: {datetime.now():%d.%m.%Y}")
    doc.add_paragraph(
        "Важно: нумерология — символическая практика саморефлексии, а не научный метод диагностики или предсказания. "
        "Интерпретации ниже следует воспринимать как темы для размышления, а не как установленные факты о личности или будущем."
    )
    add_section(doc, '1. Главный числовой код', [
        f"Для даты {d['dob_str']} в используемом здесь расчёте сумма цифр сводится к числу {life}. В выбранной символической трактовке оно связано с темой «{trait.lower()}». Это не приговор и не измерение способностей: полезнее рассматривать число как вопрос к себе — где эта тема уже проявляется, а где хотелось бы развить её осознаннее.",
        description,
        f"В практическом применении {first}, можно выбрать одну конкретную область на ближайший месяц и проверить, помогает ли вам эта тема: например, фиксировать решения, замечать повторяющиеся ситуации и обсуждать выводы с людьми, которым доверяете. Не принимайте серьёзных решений о здоровье, деньгах или отношениях только на основании нумерологической интерпретации."
    ])
    add_section(doc, '2. Числовая сумма имени', [
        f"В принятой таблице кириллических значений полное имя даёт число {expr}, символически связанное с темой «{name_trait.lower()}». Разные школы используют разные алфавитные таблицы, поэтому это число имеет смысл только в рамках выбранного правила расчёта.",
        name_description,
        "Для самонаблюдения полезно сопоставить эту трактовку с реальным опытом: какие качества подтверждаются вашими поступками, а какие не подходят? Несовпадение не является ошибкой или недостатком — оно помогает отделить собственные ценности от красивого, но чужого описания."
    ])
    add_section(doc, '3. Работа, призвание и деньги', [
        f"Символическая комбинация жизненного числа {life} и числа имени {expr} предлагает рассмотреть, как вы сочетаете тему {trait.lower()} с темой {name_trait.lower()}. В профессиональной жизни это можно перевести в конкретные вопросы: какие задачи дают вам энергию, в каком формате вы работаете лучше и какие навыки стоит развить?",
        "Нумерология не может определить гарантированный доход, идеальную профессию или удачные инвестиции. Для выбора направления опирайтесь на опыт, обучение, спрос на рынке и личные обстоятельства. Составьте список из трёх задач, которые вам удаются, и трёх условий, без которых работа становится тяжёлой; используйте его при оценке вакансий или проектов.",
        "В денежных вопросах полезнее всего практические инструменты: бюджет, резерв, сравнение условий, проверка рисков и консультация квалифицированного специалиста при необходимости. Символическое чтение может стать поводом пересмотреть привычки, но не должно заменять финансовое планирование."
    ])
    add_section(doc, '4. Отношения и семейная жизнь', [
        f"Тема {trait.lower()} может служить отправной точкой для размышлений о том, как вы строите близость: через инициативу, совместность, творчество, надёжность, свободу, заботу, глубину, общие цели или широкое понимание. Это лишь возможные вопросы, не вывод о вашем характере.",
        "Здоровые отношения не определяются числовой совместимостью. На них влияют уважение, безопасность, способность договариваться, общие ценности и поведение партнёров. Вместо процентов совместимости полезнее обсудить ожидания, границы, распределение обязанностей и способы разрешать разногласия.",
        f"Дополнительные сведения о близких: {d.get('family') or 'не предоставлены'}. Данные родственников не дают оснований утверждать, что известны их мысли, характер или будущее; их можно использовать только как контекст, который предоставил пользователь."
    ])
    add_section(doc, '5. Дополнительные сведения о рождении', [
        f"Время рождения: {d.get('birth_time', 'не указано')}. Место рождения: {d.get('birth_place', 'не указано')}. Время и место могут использоваться в некоторых астрологических системах, но сами по себе не повышают точность нумерологического расчёта даты и имени.",
        "В этом документе они зафиксированы как контекст. Асцендент и астрологические положения не рассчитываются: для них нужна отдельная астрологическая методика и точные часовой пояс, координаты и время."
    ])
    add_section(doc, '6. Практика саморефлексии', [
        "В течение недели записывайте ситуации, где выбранная тема проявилась конструктивно, и ситуации, где она мешала. Отметьте конкретные обстоятельства и собственные действия, не объясняя события «судьбой».",
        "Выберите один небольшой шаг: например, заранее подготовить разговор, завершить одну задачу, попросить о поддержке или выделить время на отдых. Через неделю оцените результат по фактам. Если символические практики — цвета, камни или мантры — вам нравятся, используйте их как личный ритуал или напоминание; не приписывайте им гарантированного лечебного или финансового эффекта."
    ])
    doc.add_heading('Итог', level=1)
    doc.add_paragraph(
        f"{first}, в рамках выбранной схемы центральными темами этого разбора стали «{trait.lower()}» и «{name_trait.lower()}». Воспринимайте их как приглашение к наблюдению и выбору, а не как окончательное описание личности. Ваши решения, отношения, способности и будущее нельзя надёжно вывести из одной даты или имени."
    )
    safe_name = re.sub(r'[^A-Za-zА-Яа-яЁё_-]+', '_', first)
    path = f"/tmp/matrica_{safe_name}_{datetime.now():%Y%m%d_%H%M%S}.docx"
    doc.save(path)
    if not os.path.exists(path) or os.path.getsize(path) < 1000:
        raise RuntimeError('DOCX was not created correctly')
    return path


async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    context.user_data.clear()
    await update.message.reply_text("Процесс отменён. Чтобы начать заново, отправьте /start.")
    return ConversationHandler.END


def main() -> None:
    app = Application.builder().token(BOT_TOKEN).build()
    conv = ConversationHandler(
        entry_points=[CommandHandler('start', start)],
        states={
            NAME: [CallbackQueryHandler(help_button, pattern='^help$'), MessageHandler(filters.TEXT & ~filters.COMMAND, name_input)],
            DOB: [MessageHandler(filters.TEXT & ~filters.COMMAND, dob_input)],
            METHOD: [CallbackQueryHandler(method_choice, pattern='^m_(pyth|kabb|vedic)$')],
            EXTRA: [CallbackQueryHandler(extra_choice, pattern='^extra_(yes|no)$')],
            TIME: [CallbackQueryHandler(time_choice, pattern='^time_(yes|no)$'), MessageHandler(filters.TEXT & ~filters.COMMAND, time_input)],
            PLACE: [CallbackQueryHandler(place_choice, pattern='^place_(yes|no)$'), MessageHandler(filters.TEXT & ~filters.COMMAND, place_input)],
            FAMILY: [CallbackQueryHandler(family_choice, pattern='^family_(yes|no)$'), MessageHandler(filters.TEXT & ~filters.COMMAND, family_input)],
            CONFIRM: [CallbackQueryHandler(confirm, pattern='^confirm_(yes|restart)$')],
        },
        fallbacks=[CommandHandler('cancel', cancel)],
        allow_reentry=True,
    )
    app.add_handler(conv)
    app.add_handler(CommandHandler('help', help_cmd))
    logger.info('Bot polling started')
    app.run_polling()


if __name__ == '__main__':
    main()

