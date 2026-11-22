"""Telegram bot: guided numerology report flow.

Configure BOT_TOKEN as a Railway Variable. Do not hardcode bot credentials.
Numerology content is framed as symbolic reflection, not factual prediction.
"""
import logging
import os
import re
from datetime import datetime
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, RGBColor
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup, ReplyKeyboardMarkup
from telegram.ext import (
    Application, CallbackQueryHandler, CommandHandler, ContextTypes,
    ConversationHandler, MessageHandler, filters,
)

TOKEN = os.getenv("BOT_TOKEN")
if not TOKEN:
    raise RuntimeError("BOT_TOKEN is missing. Add it in Railway > Variables.")

logging.basicConfig(level=logging.INFO)
log = logging.getLogger("arhet_bot")
(NAME, DOB, METHOD, EXTRA, TIME, PLACE, PARTNER_YN, PARTNER_INFO,
 CHILDREN_YN, CHILDREN_INFO, PARENTS_YN, PARENTS_INFO, REVIEW) = range(13)
METHODS = {"method_py": "Пифагорейская", "method_kab": "Каббалистическая", "method_ve": "Ведическая"}
PANEL = ReplyKeyboardMarkup([["🔄 Старт", "✖️ Отмена"]], resize_keyboard=True)
PANEL_PATTERN = r"^(🔄 Старт|✖️ Отмена)$"


def kb(rows):
    return InlineKeyboardMarkup([[InlineKeyboardButton(text, callback_data=data) for text, data in row] for row in rows])


def first_name(text):
    return text.strip().split()[0] if text and text.strip() else "друг"


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    context.user_data.clear()
    first = update.effective_user.first_name or "друг"
    greeting = (
        f"🌟 Добро пожаловать, {first}! 🌟\n\n"
        "Спасибо, что доверили мне эту часть своего пути. Я помогу вам бережно и подробно исследовать символический язык чисел — "
        "ваши сильные стороны, возможные точки роста, отношения и жизненные ориентиры.\n\n"
        "Это не приговор и не точное предсказание, а повод лучше прислушаться к себе.\n\n"
        "Для начала введите полное ФИО: имя, отчество и фамилию (три слова)."
    )
    await update.message.reply_text(greeting + "\n\nДля помощи отправьте /help. Кнопки «Старт» и «Отмена» всегда доступны внизу.", reply_markup=PANEL)
    return NAME


async def help_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    await update.message.reply_text(
        "Как проходит разбор:\n• ФИО и дата рождения;\n• выбор одной из трёх символических систем;\n"
        "• по желанию — время/место рождения и отдельные вопросы о партнёре, детях и родителях;\n"
        "• проверка сведений и отправка Word-документа.\n\n"
        "Данные близких вводите только с их согласия. Команды: /start — начать заново, /cancel — отменить."
    )


async def help_button(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    q = update.callback_query
    await q.answer()
    await q.message.reply_text("Введите полное ФИО (имя, отчество, фамилия — три слова). Дополнительные вопросы можно пропускать.")
    return NAME


async def name_input(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    name = " ".join(update.message.text.strip().split())
    parts = name.split()
    if len(parts) < 3 or not all(re.fullmatch(r"[A-Za-zА-Яа-яЁё’'-]+", p) for p in parts):
        await update.message.reply_text("Пожалуйста, введите ФИО тремя словами: имя, отчество, фамилия. Например: Иван Иванович Иванов.")
        return NAME
    context.user_data.update(full_name=name, first_name=parts[0])
    await update.message.reply_text(f"Спасибо, {parts[0]}! Теперь укажите дату рождения в формате ДД.ММ.ГГГГ, например 15.03.1990.")
    return DOB


async def dob_input(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    try:
        dob = datetime.strptime(update.message.text.strip(), "%d.%m.%Y")
        if dob.date() > datetime.now().date():
            raise ValueError
    except ValueError:
        await update.message.reply_text("Не удалось распознать дату. Введите её как ДД.ММ.ГГГГ, например 15.03.1990.")
        return DOB
    context.user_data.update(dob=dob, dob_str=dob.strftime("%d.%m.%Y"))
    await update.message.reply_text(
        "Выберите методику. Коротко: Пифагорейская — базовый разбор чисел даты и имени; "
        "каббалистическая — символическое чтение чисел в традиции гематрии; "
        "ведическая — числа в планетарной символике. Методики различаются традициями и правилами.",
        reply_markup=kb([[('🔺 Пифагорейская', 'method_py')], [('✡️ Каббалистическая', 'method_kab')], [('🕉️ Ведическая', 'method_ve')]])
    )
    return METHOD


async def method_choice(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    q = update.callback_query
    await q.answer("Методика выбрана")
    context.user_data['method'] = METHODS[q.data]
    await q.edit_message_text(
        f"✨ Вы выбрали методику «{METHODS[q.data]}».\n\nХотите добавить время и место рождения?",
        reply_markup=kb([[('Да, добавить', 'extra_yes'), ('Пропустить', 'extra_no')]])
    )
    return EXTRA


async def extra_choice(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    q = update.callback_query
    await q.answer()
    if q.data == 'extra_yes':
        await q.edit_message_text("Знаете время рождения?", reply_markup=kb([[('Да', 'time_yes'), ('Нет / пропустить', 'time_no')]]))
        return TIME
    context.user_data.update(birth_time='не указано', birth_place='не указано')
    await ask_partner(q, context)
    return PARTNER_YN


async def time_choice(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    q = update.callback_query
    await q.answer()
    if q.data == 'time_yes':
        await q.edit_message_text("Введите время рождения в формате ЧЧ:ММ, например 14:30.")
        return TIME
    context.user_data['birth_time'] = 'не указано'
    await ask_place(q, context)
    return PLACE


async def time_input(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    try:
        context.user_data['birth_time'] = datetime.strptime(update.message.text.strip(), "%H:%M").strftime("%H:%M")
    except ValueError:
        await update.message.reply_text("Формат времени: ЧЧ:ММ, например 08:45. Попробуйте ещё раз или нажмите /cancel.")
        return TIME
    await update.message.reply_text("Хотите указать место рождения?", reply_markup=kb([[('Да', 'place_yes'), ('Нет / пропустить', 'place_no')]]))
    return PLACE


async def ask_place(q, context):
    await q.edit_message_text("Хотите указать место рождения?", reply_markup=kb([[('Да', 'place_yes'), ('Нет / пропустить', 'place_no')]]))


async def place_choice(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    q = update.callback_query
    await q.answer()
    if q.data == 'place_yes':
        await q.edit_message_text("Введите город и страну рождения, например: Казань, Россия.")
        return PLACE
    context.user_data['birth_place'] = 'не указано'
    await ask_partner(q, context)
    return PARTNER_YN


async def place_input(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    place = " ".join(update.message.text.strip().split())
    if len(place) < 2 or len(place) > 100:
        await update.message.reply_text("Укажите город и страну (до 100 символов), либо /cancel.")
        return PLACE
    context.user_data['birth_place'] = place
    await ask_partner(update, context)
    return PARTNER_YN


async def ask_partner(target, context):
    msg = "Хотите включить партнёра или супруга в дополнительный анализ совместимости?"
    markup = kb([[('Да', 'partner_yes'), ('Нет / пропустить', 'partner_no')]])
    if getattr(target, 'callback_query', None):
        await target.callback_query.message.reply_text(msg, reply_markup=markup)
    else:
        await target.message.reply_text(msg, reply_markup=markup)


async def partner_choice(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    q = update.callback_query
    await q.answer()
    if q.data == 'partner_yes':
        await q.edit_message_text("Введите имя партнёра и, если хотите, дату рождения (ДД.ММ.ГГГГ). Например: Анна, 12.04.1991. Не добавляйте сведения без согласия человека.")
        return PARTNER_INFO
    context.user_data['partner'] = ''
    await ask_children(q, context)
    return CHILDREN_YN


async def partner_input(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    context.user_data['partner'] = update.message.text.strip()[:300]
    await ask_children(update, context)
    return CHILDREN_YN


async def ask_children(target, context):
    msg = "Хотите включить сведения о детях?"
    markup = kb([[('Да', 'children_yes'), ('Нет / пропустить', 'children_no')]])
    if getattr(target, 'callback_query', None): await target.callback_query.message.reply_text(msg, reply_markup=markup)
    else: await target.message.reply_text(msg, reply_markup=markup)


async def children_choice(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    q = update.callback_query
    await q.answer()
    if q.data == 'children_yes':
        await q.edit_message_text("Укажите только желаемые сведения о детях (например, число детей; имена/даты рождения — только при необходимости и с согласия).")
        return CHILDREN_INFO
    context.user_data['children'] = ''
    await ask_parents(q, context)
    return PARENTS_YN


async def children_input(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    context.user_data['children'] = update.message.text.strip()[:500]
    await ask_parents(update, context)
    return PARENTS_YN


async def ask_parents(target, context):
    msg = "Хотите включить сведения о родителях и семейных отношениях?"
    markup = kb([[('Да', 'parents_yes'), ('Нет / пропустить', 'parents_no')]])
    if getattr(target, 'callback_query', None): await target.callback_query.message.reply_text(msg, reply_markup=markup)
    else: await target.message.reply_text(msg, reply_markup=markup)


async def parents_choice(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    q = update.callback_query
    await q.answer()
    if q.data == 'parents_yes':
        await q.edit_message_text("Напишите кратко, какой аспект отношений с родителями хотите включить. Имена и даты рождения не обязательны.")
        return PARENTS_INFO
    context.user_data['parents'] = ''
    await show_review(update, context)
    return REVIEW


async def parents_input(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    context.user_data['parents'] = update.message.text.strip()[:800]
    await show_review(update, context)
    return REVIEW


async def show_review(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    d = context.user_data
    details = []
    if d.get('partner'): details.append('партнёр: да')
    if d.get('children'): details.append('дети: да')
    if d.get('parents'): details.append('родители: да')
    text = (
        "Пожалуйста, проверьте сведения перед созданием документа:\n\n"
        f"ФИО: {d['full_name']}\nДата рождения: {d['dob_str']}\nМетодика: {d['method']}\n"
        f"Время рождения: {d.get('birth_time', 'не указано')}\nМесто рождения: {d.get('birth_place', 'не указано')}\n"
        f"Дополнительные разделы: {', '.join(details) if details else 'не добавлены'}\n\nВсё верно?"
    )
    markup = kb([[('✅ Да, создать разбор', 'confirm_yes')], [('↩️ Начать заново', 'confirm_restart')]])
    if update.callback_query:
        await update.callback_query.message.reply_text(text, reply_markup=markup)
    else:
        await update.message.reply_text(text, reply_markup=markup)


async def confirm(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    q = update.callback_query
    await q.answer()
    if q.data == 'confirm_restart':
        context.user_data.clear()
        await q.edit_message_text("Начнём заново. Введите полное ФИО (имя, отчество, фамилия).")
        return NAME
    await q.edit_message_text("✨ Спасибо за подтверждение! Подготавливаю ваш подробный документ…")
    path = None
    try:
        path = build_report(context.user_data)
        with open(path, 'rb') as f:
            await context.bot.send_document(
                chat_id=q.from_user.id, document=f, filename=Path(path).name,
                caption=f"Ваш разбор готов, {context.user_data['first_name']}! Спасибо, что выбрали нас — мы ценим ваше доверие и подготовили для вас подробный материал."
            )
        await context.bot.send_message(
            q.from_user.id,
            f"💛 {context.user_data['first_name']}, спасибо, что выбрали нас. Мы ценим ваше доверие и надеемся, что разбор станет полезным поводом для размышлений.\n\nЧтобы начать заново, отправьте /start."
        )
    except Exception:
        log.exception("Report generation or document delivery failed")
        await context.bot.send_message(q.from_user.id, "Не получилось сформировать или отправить документ. Попробуйте /start ещё раз. Если ошибка повторится, администратору нужно проверить Railway Logs.")
    finally:
        if path and os.path.exists(path):
            os.remove(path)
        context.user_data.clear()
    return ConversationHandler.END


async def panel_action(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Routes persistent reply-keyboard buttons in any dialogue state."""
    pressed = update.message.text.strip()
    if pressed == "🔄 Старт":
        return await start(update, context)
    if pressed == "✖️ Отмена":
        return await cancel(update, context)
    return ConversationHandler.END


async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    context.user_data.clear()
    await update.message.reply_text("Разбор отменён. Нажмите «🔄 Старт», чтобы начать заново.", reply_markup=PANEL)
    return ConversationHandler.END


NUM = {
    1: ("Инициатива и самостоятельность", "Эта тема приглашает замечать, где вам естественно брать инициативу, а где полезнее делегировать или выслушать других. Сильная сторона — способность начать движение и обозначить направление. Возможная трудность — ощущение, что всё нужно делать самому, или нетерпение к чужому темпу. В работе попробуйте различать ответственность и контроль: лидерство может быть не только личным рывком, но и созданием условий, в которых другие тоже действуют. В отношениях прямота помогает, если рядом с ней есть интерес к потребностям другого человека. Практический вопрос: в какой сфере сейчас достаточно сделать первый небольшой шаг, не требуя от себя сразу идеального результата?"),
    2: ("Сотрудничество и чуткость", "Эта тема связана с вниманием к нюансам общения, умением замечать разные точки зрения и выстраивать партнёрство. Её ресурс — способность создавать диалог и видеть контекст. Трудность может возникать, когда желание сохранить мир заставляет откладывать собственные решения. Полезная практика — сформулировать свою позицию до обсуждения и проверить, что компромисс не превращается в отказ от важных для вас границ. В работе сотрудничество особенно продуктивно, если роли и ожидания проговорены. В отношениях ясность часто добрее догадок: говорите о потребностях конкретно и спокойно."),
    3: ("Самовыражение и творчество", "Эта тема приглашает исследовать способы выражать идеи, чувства и наблюдения. Ресурс — воображение, общение и способность находить выразительную форму. Возможная сложность — начинать много интересного и терять внимание на этапе завершения. Выберите один проект и определите маленький, измеримый результат на неделю. В профессиональной сфере творческая подача особенно ценна, когда соединена с регулярностью и пониманием аудитории. В отношениях важно не только разряжать напряжение юмором, но и оставаться в разговоре, когда он становится серьёзным. Спросите себя: что я давно хочу сказать, создать или показать, но всё откладываю?"),
    4: ("Опора и последовательность", "Эта тема связана со структурой, надёжностью и постепенным созданием результата. Ресурс — умение выстраивать порядок и выдерживать длинную дистанцию. Риск — слишком крепко держаться за привычный план, даже когда обстоятельства изменились. Помогают регулярный пересмотр целей и пространство для эксперимента. В работе полезны понятные критерии качества и реалистичный график; не обязательно делать всё идеально с первого раза. В отношениях стабильность проявляется также в способности обсуждать перемены, а не только сохранять привычное. Подумайте, какая опора в вашей жизни сейчас действительно поддерживает вас, а какая стала ограничением."),
    5: ("Гибкость и перемены", "Эта тема предлагает исследовать любознательность, свободу выбора и способность адаптироваться. Ресурс — быстро замечать новые варианты и учиться на опыте. Возможная трудность — импульсивно менять курс или путать свободу с отсутствием обязательств. Сохраняйте пространство для нового, но перед крупным решением запишите последствия, расходы и критерии пересмотра. В профессиональной сфере разнообразие задач может поддерживать интерес, если есть одна главная цель. В близости важно согласовывать потребность в личном пространстве с надёжностью договорённостей. Что нового вы можете попробовать безопасно, не разрушая то, что уже работает?"),
    6: ("Забота и ответственность", "Эта тема связана с вниманием к близким, стремлением создавать комфорт и принимать обязательства. Ресурс — надёжность и участие. Трудность — чрезмерно брать на себя чужие задачи, а затем испытывать усталость или обиду. Полезно проверять: меня просили о помощи, есть ли у меня ресурс и что я ожидаю взамен? В работе забота о качестве не должна означать бесконечную доступность. В отношениях поддержку укрепляют прямые просьбы, взаимность и уважение к границам. Включайте в заботу о других заботу о себе: отдых — не награда за истощение, а часть устойчивого ритма."),
    7: ("Анализ и внутренний поиск", "Эта тема связана с наблюдением, изучением и стремлением понять причины происходящего. Ресурс — внимательность к деталям и способность глубоко погружаться в вопрос. Возможная трудность — изоляция, чрезмерное обдумывание или ожидание полной уверенности до первого действия. Помогает сочетать размышление с проверкой на практике: сформулируйте гипотезу и найдите небольшой способ проверить её. В работе выделяйте время для концентрации, но не забывайте синхронизироваться с командой. В отношениях делитесь выводами и чувствами простым языком, не ожидая, что другие догадаются о вашем внутреннем мире."),
    8: ("Цели и управление ресурсами", "Эта тема предлагает исследовать амбиции, организацию и отношение к влиянию. Ресурс — умение ставить цели, распределять ресурсы и отвечать за результат. Возможная трудность — измерять собственную ценность только достижениями или игнорировать цену постоянной нагрузки. Успех становится устойчивее, когда включает этику, здоровье и качество отношений. Перед крупным финансовым или карьерным шагом проверяйте факты, риски и запасной план. В команде авторитет укрепляется не давлением, а ясными ожиданиями и справедливостью. Определите, какой результат для вас важен не только внешне, но и по тому, как вы хотите жить."),
    9: ("Смысл, завершение и широкий взгляд", "Эта тема связана с умением видеть общую картину, извлекать уроки и завершать этапы. Ресурс — сочувствие и способность учитывать интересы разных людей. Возможная трудность — идеализировать окружающих, чрезмерно отдавать силы или откладывать собственные начинания ради чужих задач. Полезно определять границы помощи и конкретный критерий завершения. В работе смысл поддерживает мотивацию, но практические договорённости помогают доводить замыслы до результата. В отношениях сострадание не требует терпеть небезопасное или неуважительное поведение. Какой завершённый этап освободит место для следующего шага?"),
}


def digit_sum(value: str) -> int:
    n = sum(int(ch) for ch in value if ch.isdigit())
    while n > 9:
        n = sum(int(ch) for ch in str(n))
    return n or 1


def name_num(full_name: str) -> int:
    alphabet = 'АБВГДЕЁЖЗИЙКЛМНОПРСТУФХЦЧШЩЪЫЬЭЮЯ'
    values = {c: i % 9 + 1 for i, c in enumerate(alphabet)}
    return digit_sum(str(sum(values.get(ch, 0) for ch in full_name.upper())))


def add_section(doc, heading, paragraphs):
    doc.add_heading(heading, level=1)
    for p in paragraphs:
        doc.add_paragraph(p)


def build_report(d: dict) -> str:
    full = d['full_name']
    first = d['first_name']
    life = digit_sum(d['dob'].strftime('%d%m%Y'))
    expr = name_num(full)
    life_title, life_text = NUM[life]
    name_title, name_text = NUM[expr]
    doc = Document()
    normal = doc.styles['Normal']
    normal.font.name = 'Arial'
    normal.font.size = Pt(10)
    normal.paragraph_format.space_after = Pt(5)
    normal.paragraph_format.line_spacing = 1.08

    # Five intentionally structured sections/pages; pagination may vary slightly by Word version.
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run('ВАША МАТРИЦА ЧИСЕЛ')
    r.bold = True; r.font.size = Pt(22); r.font.color.rgb = RGBColor(40, 70, 110)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run('Персональный символический разбор').italic = True
    doc.add_paragraph(f"ФИО: {full}\nДата рождения: {d['dob_str']}\nВыбранная традиция: {d['method']}\nДата подготовки: {datetime.now():%d.%m.%Y}")
    doc.add_paragraph("Важно: нумерология — символическая практика для саморефлексии, а не научная диагностика или достоверное предсказание. Расчёты зависят от выбранной таблицы и метода.")
    add_section(doc, '1. Как читать этот разбор', [
        f"{first}, этот материал подготовлен по данным, которые вы указали, и выбранной вами традиции — {d['method'].lower()}. Полное имя приведено выше как исходное данное; далее обращение идёт только по имени.",
        "Числа здесь используются как символические темы для размышления. Они не могут достоверно определить характер, детские события, здоровье, профессию, отношения или будущее. Если какое-то описание вам не подходит, не нужно подгонять под него свой опыт: полезнее отметить расхождение и опираться на реальные наблюдения.",
        "Расчёт жизненного числа в этом документе — сумма цифр даты рождения с приведением к однозначному числу. Число имени получено по циклическому соответствию русских букв значениям 1–9. В разных школах правила отличаются, поэтому результаты могут не совпадать с другими калькуляторами."
    ])
    doc.add_page_break()

    add_section(doc, f'2. Основные числовые темы: {life} и {expr}', [
        f"Жизненный код {life} в выбранной символической трактовке связан с темой «{life_title.lower()}». {life_text}",
        f"Число имени {expr} связано с темой «{name_title.lower()}». {name_text}",
        f"Вместе темы «{life_title.lower()}» и «{name_title.lower()}» предлагают не готовый портрет, а полезное сопоставление. Где вы чаще проявляете одну тему? Где другая помогает уравновесить первую? Например, стремление к последовательности может поддержать творческий замысел, а гибкость — смягчить чрезмерную привязанность к плану. Ответы стоит искать в конкретных ситуациях, а не считать неизбежными свойствами.",
        f"Практика на неделю для {first}: запишите три ситуации, где эти темы были заметны. Для каждой отметьте факт, своё решение и результат. Затем выберите один навык, который хотите попробовать укрепить. Такой подход превращает символическое чтение в личное наблюдение, не выдавая его за объективный диагноз."
    ])
    doc.add_page_break()

    add_section(doc, '3. Деятельность, цели и ресурсы', [
        f"Связка чисел {life} и {expr} может стать поводом подумать о том, как вы принимаете решения, взаимодействуете с задачами и поддерживаете мотивацию. Вопросы к себе: какие занятия дают ощущение смысла? В каких условиях вы работаете устойчивее — самостоятельно, в команде, по гибкому графику или по ясному плану? Какой навык хочется развить независимо от нумерологической трактовки?",
        "Нельзя по дате рождения надёжно определить профессию, уровень дохода или благоприятный момент для инвестиций. Для выбора работы полезнее соединить интересы и способности с образованием, опытом, спросом и бытовыми условиями. Составьте перечень задач, которые у вас получаются, и условий, которые важны для качества жизни; сравнивайте с ними конкретные предложения.",
        "В финансовых вопросах используйте проверяемые инструменты: бюджет, резерв на непредвиденные расходы, анализ условий и рисков. Не принимайте финансовых решений только по числам, цветам, камням или датам. Символическая практика может напоминать о цели, но не заменяет расчёт и профессиональную консультацию.",
        "Мини-план: сформулируйте одну цель на ближайшие 30 дней; определите первый шаг, который можно выполнить за час; назначьте дату проверки результата. Если цель не продвигается, уточните препятствие и уменьшите шаг, вместо того чтобы объяснять задержку «неподходящим числом»."
    ])
    doc.add_page_break()

    relationship_paras = [
        f"Тема «{life_title.lower()}» может помочь задать вопросы о способе строить близость, но сама по себе не устанавливает совместимость. Отношения зависят от поведения людей, уважения, безопасности, общих ценностей, обстоятельств и готовности договариваться.",
        "Вместо процентов совместимости попробуйте обсудить четыре практические области: как каждый просит о поддержке; как договариваетесь о времени и личном пространстве; как распределяете бытовые обязанности; как восстанавливаете контакт после разногласия. Если есть страх, принуждение или насилие, обращайтесь за помощью к надёжным людям и профильным службам — нумерологический разбор не заменяет поддержку.",
        f"Сведения о партнёре: {d.get('partner') or 'не добавлены'}. Сведения о детях: {d.get('children') or 'не добавлены'}. Сведения о родителях/семейной теме: {d.get('parents') or 'не добавлены'}. Эти сведения включены только как предоставленный вами контекст; они не позволяют достоверно утверждать, что мы знаем внутренние качества или будущее ваших близких.",
        "Если вы добавили сведения о семье, используйте их как основу для вопросов, а не ярлыков: что помогает вам чувствовать себя услышанным? Какие договорённости стоит проговорить? Где полезно обозначить границы? Не пересылайте документ с личными данными близким без согласования."
    ]
    add_section(doc, '4. Отношения и семейный контекст', relationship_paras)
    doc.add_page_break()

    rec = {
        1: ('инициатива', 'выбрать один приоритет и сделать первый небольшой шаг'),
        2: ('сотрудничество', 'сформулировать личную просьбу и спокойно обсудить её'),
        3: ('самовыражение', 'уделить 20 минут одному творческому или коммуникативному делу'),
        4: ('последовательность', 'разбить важную задачу на три выполнимых этапа'),
        5: ('гибкость', 'попробовать одно безопасное изменение и оценить результат'),
        6: ('забота', 'сделать один полезный шаг для себя наряду с помощью близким'),
        7: ('рефлексия', 'записать вопрос и проверить одну гипотезу действием'),
        8: ('цели и ресурсы', 'оценить цель по срокам, ресурсам, рискам и личной цене'),
        9: ('завершение', 'закрыть один небольшой незавершённый этап и освободить внимание'),
    }
    theme, action = rec[life]
    add_section(doc, '5. Практические ориентиры и завершение', [
        f"Время рождения: {d.get('birth_time', 'не указано')}. Место рождения: {d.get('birth_place', 'не указано')}. Эти сведения сохранены как контекст. Асцендент и положение небесных тел здесь не рассчитываются: для этого нужна отдельная астрологическая методика, точное время и часовой пояс. Для нумерологического сложения они не требуются.",
        f"Практический ориентир по теме «{theme}»: {action}. Относитесь к этому как к эксперименту, а не обязательному ритуалу. Через неделю спросите себя, что действительно помогло, что оказалось нейтральным и что хочется изменить.",
        "Цвета, минералы, ароматы и мантры некоторые традиции связывают с числами. Если они вам нравятся, используйте их как личные символы или напоминания; их действие не гарантирует здоровье, защиту, любовь или финансовый успех. Не заменяйте ими медицинскую помощь, обучение, открытый разговор или профессиональный совет.",
        f"{first}, спасибо, что выбрали нас и доверили нам подготовку этого разбора. Мы ценим ваше доверие и надеемся, что материал получился для вас подробным, бережным и полезным. Возвращайтесь к тем мыслям, которые откликаются именно вашему опыту, и оставляйте то, что вам не подходит. Ваш путь определяется вашими решениями и обстоятельствами, а не одной цифрой."
    ])

    safe = re.sub(r'[^A-Za-zА-Яа-яЁё_-]+', '_', first)
    path = f"/tmp/matrix_{safe}_{datetime.now():%Y%m%d_%H%M%S}.docx"
    doc.save(path)
    if not os.path.exists(path) or os.path.getsize(path) < 1000:
        raise RuntimeError('DOCX file was not created correctly')
    return path


def main():
    app = Application.builder().token(TOKEN).build()
    panel_handler = MessageHandler(filters.Regex(PANEL_PATTERN), panel_action)
    conv = ConversationHandler(
        entry_points=[CommandHandler('start', start), panel_handler],
        states={
            NAME: [panel_handler, CallbackQueryHandler(help_button, pattern='^help$'), MessageHandler(filters.TEXT & ~filters.COMMAND, name_input)],
            DOB: [panel_handler, MessageHandler(filters.TEXT & ~filters.COMMAND, dob_input)],
            METHOD: [panel_handler, CallbackQueryHandler(method_choice, pattern='^method_(py|kab|ve)$')],
            EXTRA: [panel_handler, CallbackQueryHandler(extra_choice, pattern='^extra_(yes|no)$')],
            TIME: [panel_handler, CallbackQueryHandler(time_choice, pattern='^time_(yes|no)$'), MessageHandler(filters.TEXT & ~filters.COMMAND, time_input)],
            PLACE: [panel_handler, CallbackQueryHandler(place_choice, pattern='^place_(yes|no)$'), MessageHandler(filters.TEXT & ~filters.COMMAND, place_input)],
            PARTNER_YN: [panel_handler, CallbackQueryHandler(partner_choice, pattern='^partner_(yes|no)$')],
            PARTNER_INFO: [panel_handler, MessageHandler(filters.TEXT & ~filters.COMMAND, partner_input)],
            CHILDREN_YN: [panel_handler, CallbackQueryHandler(children_choice, pattern='^children_(yes|no)$')],
            CHILDREN_INFO: [panel_handler, MessageHandler(filters.TEXT & ~filters.COMMAND, children_input)],
            PARENTS_YN: [panel_handler, CallbackQueryHandler(parents_choice, pattern='^parents_(yes|no)$')],
            PARENTS_INFO: [panel_handler, MessageHandler(filters.TEXT & ~filters.COMMAND, parents_input)],
            REVIEW: [panel_handler, CallbackQueryHandler(confirm, pattern='^confirm_(yes|restart)$')],
        },
        fallbacks=[CommandHandler('cancel', cancel)],
        allow_reentry=True,
    )
    app.add_handler(conv)
    app.add_handler(CommandHandler('help', help_cmd))
    # Handles the persistent keyboard also when no conversation is active.
    app.add_handler(MessageHandler(filters.Regex(PANEL_PATTERN), panel_action))
    log.info('Bot started with persistent Start/Cancel keyboard')
    app.run_polling()


if __name__ == '__main__':
    main()
