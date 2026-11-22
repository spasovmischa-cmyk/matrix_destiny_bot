import logging
import os
from datetime import datetime
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes, ConversationHandler
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

# Логирование
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Состояния диалога
WAITING_NAME, WAITING_DATE, WAITING_METHODOLOGY = range(3)

# Токен бота
BOT_TOKEN = "9025753024:AAEdoVJNl30GUrIbFGK33MTJK6v3UNdCZL8"

# Данные пользователя (временно в памяти)
user_data = {}

# ==================== ПРИВЕТСТВИЕ И ПОМОЩЬ ====================

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Команда /start с приветствием"""
    user = update.effective_user
    
    welcome_text = f"""
🌟 *Добро пожаловать, {user.first_name}!* 🌟

Я — ваш персональный гид по матрице судьбы.

Я создам для вас полный нумерологический разбор жизни, который раскроет:
✨ Ваше истинное призвание и предназначение
✨ Скрытые таланты и способности
✨ Путь к гармонии в отношениях
✨ Финансовый потенциал
✨ Здоровье и энергию
✨ Духовный путь

📋 *Как это работает:*
1. Введите ваше полное имя
2. Укажите дату рождения (ДД.MM.ГГГГ)
3. Выберите методику анализа
4. Получите подробный разбор в виде красивого документа

🎯 Готовы начать? Нажмите кнопку ниже или введите /help для справки.
    """
    
    keyboard = [
        [InlineKeyboardButton("🚀 Начать разбор", callback_data="start_analysis")],
        [InlineKeyboardButton("ℹ️ Справка", callback_data="help")]
    ]
    reply_markup = InlineKeyboardMarkup(keyboard)
    
    await update.message.reply_text(
        welcome_text,
        parse_mode="Markdown",
        reply_markup=reply_markup
    )
    
    return WAITING_NAME


async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Команда /help с информацией"""
    help_text = """
📚 *СПРАВКА ПО БОТУ*

*Что такое матрица судьбы?*
Это полный нумерологический анализ вашей жизни, основанный на древних системах:
• Пифагорейская нумерология (западная классика)
• Каббалистическая нумерология (еврейская мистика)
• Ведическая нумерология (индийская система)

*Что вы получите?*
✓ Полный разбор в документе Word
✓ 11 глав анализа (от рождения до будущего)
✓ Рекомендации по цветам, камням, мантрам
✓ Анализ карьеры и финансов
✓ Советы по здоровью и отношениям

*Как начать?*
Нажмите /start или напишите любое сообщение

*Частые вопросы:*
❓ Сколько времени занимает разбор?
Несколько минут. Вы сразу получите готовый документ.

❓ Точна ли нумерология?
Это система символов, которая помогает понять глубокие закономерности жизни.

❓ Нужна дата рождения по старому стилю?
Нет, по современному календарю (ДД.MM.ГГГГ).

*Команды:*
/start - Начать новый разбор
/help - Эта справка
/cancel - Отмена текущего процесса

✨ Готовы открыть секреты своей судьбы?
    """
    
    await update.message.reply_text(help_text, parse_mode="Markdown")


async def button_callback(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Обработка кнопок меню"""
    query = update.callback_query
    await query.answer()
    
    if query.data == "start_analysis":
        await query.edit_message_text(
            text="📝 *Введите ваше полное имя:*\n\n(Например: Иван Иванович Иванов)",
            parse_mode="Markdown"
        )
        return WAITING_NAME
    
    elif query.data == "help":
        help_text = """
📚 *СПРАВКА ПО БОТУ*

*Что такое матрица судьбы?*
Это полный нумерологический анализ вашей жизни.

*Как начать?*
Нажмите "Начать разбор" и следуйте инструкциям.

*Команды:*
/start - Начать новый разбор
/help - Справка
/cancel - Отмена
        """
        await query.edit_message_text(text=help_text, parse_mode="Markdown")


async def receive_name(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Получение имени пользователя"""
    name = update.message.text.strip()
    
    if len(name) < 2:
        await update.message.reply_text("❌ Пожалуйста, введите полное имя (минимум 2 символа)")
        return WAITING_NAME
    
    user_id = update.effective_user.id
    user_data[user_id] = {"name": name}
    
    await update.message.reply_text(
        f"✅ Спасибо, {name}!\n\n📅 Теперь введите вашу дату рождения:\n\n(Формат: ДД.MM.ГГГГ\nНапример: 15.03.1990)"
    )
    
    return WAITING_DATE


async def receive_date(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Получение даты рождения"""
    date_str = update.message.text.strip()
    
    try:
        birth_date = datetime.strptime(date_str, "%d.%m.%Y")
        
        user_id = update.effective_user.id
        user_data[user_id]["birth_date"] = birth_date
        user_data[user_id]["birth_date_str"] = date_str
        
        # Выбор методики
        keyboard = [
            [InlineKeyboardButton("🔺 Пифагорейская", callback_data="method_pythagoras")],
            [InlineKeyboardButton("✡️ Каббалистическая", callback_data="method_kabbalah")],
            [InlineKeyboardButton("🕉️ Ведическая", callback_data="method_vedic")]
        ]
        reply_markup = InlineKeyboardMarkup(keyboard)
        
        await update.message.reply_text(
            f"✅ Дата рождения: {date_str}\n\n🌟 *Выберите методику анализа:*",
            parse_mode="Markdown",
            reply_markup=reply_markup
        )
        
        return WAITING_METHODOLOGY
        
    except ValueError:
        await update.message.reply_text(
            "❌ Неверный формат даты!\n\nПожалуйста, введите в формате: ДД.MM.ГГГГ\n(Например: 15.03.1990)"
        )
        return WAITING_DATE


async def receive_methodology(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Получение выбранной методики"""
    query = update.callback_query
    await query.answer()
    
    user_id = query.from_user.id
    
    methodology_map = {
        "method_pythagoras": "Пифагорейская",
        "method_kabbalah": "Каббалистическая",
        "method_vedic": "Ведическая"
    }
    
    methodology = methodology_map[query.data]
    user_data[user_id]["methodology"] = methodology
    
    # Показываем статус обработки
    await query.edit_message_text(
        text=f"✅ Методика: *{methodology}*\n\n⏳ Создаю вашу матрицу судьбы...",
        parse_mode="Markdown"
    )
    
    # Генерируем разбор
    try:
        file_path = generate_destiny_matrix(user_data[user_id])
        
        # Отправляем файл
        with open(file_path, 'rb') as file:
            await context.bot.send_document(
                chat_id=user_id,
                document=file,
                caption=f"🌟 *Ваша матрица судьбы готова!*\n\nИмя: {user_data[user_id]['name']}\nДата рождения: {user_data[user_id]['birth_date_str']}\nМетодика: {methodology}\n\n✨ Откройте документ и изучите все 11 глав вашей судьбы!",
                parse_mode="Markdown"
            )
        
        # Финальное сообщение
        await context.bot.send_message(
            chat_id=user_id,
            text="✨ *Спасибо за использование бота!*\n\n💫 Пусть числа вашей судьбы ведут вас к гармонии и процветанию!\n\n🔄 Введите /start, чтобы создать ещё один разбор.",
            parse_mode="Markdown"
        )
        
        # Удаляем файл после отправки
        os.remove(file_path)
        
    except Exception as e:
        logger.error(f"Ошибка при генерации разбора: {e}")
        await context.bot.send_message(
            chat_id=user_id,
            text="❌ Произошла ошибка при создании разбора. Пожалуйста, попробуйте позже или введите /start для повтора."
        )
    
    return ConversationHandler.END


async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Отмена процесса"""
    user_id = update.effective_user.id
    if user_id in user_data:
        del user_data[user_id]
    
    await update.message.reply_text(
        "❌ Процесс отменён.\n\n🔄 Введите /start, чтобы начать заново."
    )
    return ConversationHandler.END


# ==================== ГЕНЕРАЦИЯ РАЗБОРА ====================

def generate_destiny_matrix(data: dict) -> str:
    """Генерирует Word документ с матрицей судьбы"""
    
    name = data["name"]
    birth_date_str = data["birth_date_str"]
    birth_date = data["birth_date"]
    methodology = data["methodology"]
    
    # Вычисляем числа судьбы
    day, month, year = birth_date.day, birth_date.month, birth_date.year
    
    # Число Судьбы = сумма дня+месяца+года
    destiny_number = sum_to_single_digit(day + month + year)
    
    # Число имени
    name_number = calculate_name_number(name)
    
    # Создаём документ
    doc = Document()
    
    # Титульный лист
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run('МАТРИЦА СУДЬБЫ')
    run.font.size = Pt(28)
    run.font.bold = True
    run.font.color.rgb = RGBColor(0, 102, 204)
    
    subtitle = doc.add_paragraph()
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = subtitle.add_run('Персональный нумерологический разбор')
    run.font.size = Pt(14)
    run.font.italic = True
    
    # Личные данные
    doc.add_paragraph()
    info_table = doc.add_table(rows=4, cols=2)
    info_table.style = 'Light Grid Accent 1'
    
    info_data = [
        ("Имя:", name),
        ("Дата рождения:", birth_date_str),
        ("Методика анализа:", methodology),
        ("Дата разбора:", datetime.now().strftime("%d.%m.%Y"))
    ]
    
    for i, (key, value) in enumerate(info_data):
        info_table.rows[i].cells[0].text = key
        info_table.rows[i].cells[1].text = value
    
    doc.add_page_break()
    
    # ==================== ОСНОВНЫЕ ЧИСЛА ====================
    
    doc.add_heading('ВАШИ ЧИСЛОВЫЕ КОДЫ', level=1)
    
    doc.add_heading(f'Число Судьбы: {destiny_number}', level=2)
    doc.add_paragraph(
        f'Это ваш главный жизненный путь. Число {destiny_number} определяет ваше основное предназначение и энергию, которая движет вашей жизнью.'
    )
    
    # Описание числа судьбы
    destiny_descriptions = {
        1: "Вы лидер и инициатор. Ваша задача — начинать новое, быть независимым и творить. Вас ждёт путь к власти и влиянию через собственное развитие.",
        2: "Вы дипломат и медиатор. Ваша сила в гармонии, сотрудничестве и интуиции. Вы научитесь видеть скрытые смыслы и находить общий язык с людьми.",
        3: "Вы творец и коммуникатор. Ваша задача — выражать себя, общаться, творить. Успех приходит через искусство, слово и самовыражение.",
        4: "Вы созидатель и строитель. Ваша миссия — создавать прочное основание, быть надёжным и практичным. Богатство приходит через упорный труд.",
        5: "Вы путешественник и свободный дух. Вам даны перемены, гибкость, приключения. Ваш путь — это постоянное движение и адаптация.",
        6: "Вы помощник и ответственный. Ваша задача — служить, заботиться, создавать гармонию. Вы найдёте счастье в семье и служении.",
        7: "Вы мудрец и духовный искатель. Вас зовут к анализу, знаниям и духовности. Ваш путь — это внутренняя работа и просветление.",
        8: "Вы материалист и человек силы. Вам дана власть, амбиция, достаток. Ваша задача — создавать материальное благоустройство и влиять на мир.",
        9: "Вы гуманист и завершитель. Ваша миссия — служить человечеству, завершать циклы, быть мудрецом. Вас ждёт глубокое понимание жизни."
    }
    
    doc.add_paragraph(destiny_descriptions.get(destiny_number, ""))
    
    doc.add_heading(f'Число Имени: {name_number}', level=2)
    doc.add_paragraph(
        f'Это ваша личность и то, как вас воспринимают люди. Число {name_number} описывает ваши основные качества и способности.'
    )
    
    name_descriptions = {
        1: "Вы воспринимаетесь как лидер, независимый и сильный человек.",
        2: "Вы кажетесь чувствительным, дипломатичным и интуитивным.",
        3: "Вы выглядите творческим, весёлым и общительным.",
        4: "Вы производите впечатление надёжного, практичного и серьёзного человека.",
        5: "Вы кажетесь свободолюбивым, активным и приключенческим.",
        6: "Вы выглядите заботливым, ответственным и гармоничным.",
        7: "Вы производите впечатление мудрого, аналитичного и духовного человека.",
        8: "Вы кажетесь амбициозным, сильным и материально успешным.",
        9: "Вы выглядите мудрым, сострадательным и универсальным."
    }
    
    doc.add_paragraph(name_descriptions.get(name_number, ""))
    
    doc.add_page_break()
    
    # ==================== КАРЬЕРА И ФИНАНСЫ ====================
    
    doc.add_heading('КАРЬЕРА И ФИНАНСЫ', level=1)
    
    career_advice = {
        1: "Идеальны профессии: предприниматель, руководитель, тренер. Ваш финансовый потенциал высок, если вы создаёте собственное дело.",
        2: "Идеальны профессии: психолог, медиатор, художник. Финансовая стабильность приходит через партнёрство и сотрудничество.",
        3: "Идеальны профессии: писатель, актёр, маркетолог. Доход зависит от вашей креативности и способности общаться.",
        4: "Идеальны профессии: инженер, архитектор, бухгалтер. Богатство приходит через дисциплину и упорный труд.",
        5: "Идеальны профессии: журналист, продавец, консультант. Финансы волнообразны, но возможности велики.",
        6: "Идеальны профессии: врач, учитель, социальный работник. Финансы стабильны благодаря служению.",
        7: "Идеальны профессии: учёный, философ, программист. Потенциал огромен через знания и инновации.",
        8: "Идеальны профессии: бизнесмен, юрист, финансист. Ваш финансовый потенциал исключительно высок.",
        9: "Идеальны профессии: философ, благотворитель, врач. Финансы приходят в конце цикла, нужна терпение."
    }
    
    doc.add_paragraph(career_advice.get(destiny_number, ""))
    
    doc.add_page_break()
    
    # ==================== ОТНОШЕНИЯ ====================
    
    doc.add_heading('ОТНОШЕНИЯ И СЕМЬЯ', level=1)
    
    relationships_advice = {
        1: "Вам нужен партнёр, который уважает вашу независимость. Лучше всего с числами 3, 5, 9.",
        2: "Вам важна гармония и поддержка. Лучше всего подходят числа 2, 4, 6, 8.",
        3: "Вам нужна весёлая и творческая пара. Хорошо с числами 3, 5, 6.",
        4: "Вам нужна надёжная и стабильная пара. Идеально с числами 2, 4, 8.",
        5: "Вам нужна свободная и приключенческая пара. Хорошо с числами 1, 3, 5, 7.",
        6: "Вам нужна забота и близость. Идеально с числами 2, 4, 6, 9.",
        7: "Вам нужна интеллектуальная и духовная связь. Хорошо с числами 2, 7, 9.",
        8: "Вам нужен партнёр, понимающий вашу амбицию. Идеально с числами 1, 4, 8.",
        9: "Вам нужна мудрая и духовная пара. Хорошо с числами 1, 6, 9."
    }
    
    doc.add_paragraph(relationships_advice.get(destiny_number, ""))
    
    doc.add_page_break()
    
    # ==================== РЕКОМЕНДАЦИИ ====================
    
    doc.add_heading('РЕКОМЕНДАЦИИ И ПРАКТИКИ', level=1)
    
    recommendations = {
        1: {
            "цвета": "Красный, золотой, оранжевый",
            "камни": "Рубин, гранат, яшма",
            "мантра": "Я - сильный лидер",
            "практика": "Медитируйте на число 1, занимайтесь спортом, лидируйте в малых группах"
        },
        2: {
            "цвета": "Оранжевый, кремовый, белый",
            "камни": "Лунный камень, жемчуг",
            "мантра": "Я спокоен и гармоничен",
            "практика": "Йога, медитация, работа с эмоциями, слушание музыки"
        },
        3: {
            "цвета": "Жёлтый, золотой, лиловый",
            "камни": "Топаз, цитрин, аметист",
            "мантра": "Я выражаю свою суть",
            "практика": "Творчество, пение, танцы, общение с друзьями"
        },
        4: {
            "цвета": "Зелёный, синий, коричневый",
            "камни": "Изумруд, нефрит, малахит",
            "мантра": "Я стабилен и надёжен",
            "практика": "Упражнения, структурированные занятия, работа с руками"
        },
        5: {
            "цвета": "Серый, голубой, бирюзовый",
            "камни": "Бирюза, аквамарин",
            "мантра": "Я свободен и адаптивен",
            "практика": "Путешествия, новые знакомства, изучение языков"
        },
        6: {
            "цвета": "Голубой, розовый, белый",
            "камни": "Опал, турмалин, агат",
            "мантра": "Я люблю и забочусь",
            "практика": "Работа с семьёй, волонтёрство, забота о других"
        },
        7: {
            "цвета": "Фиолетовый, индиго, тёмно-синий",
            "камни": "Аметист, лабрадорит",
            "мантра": "Я ищу истину",
            "практика": "Медитация, чтение, природа, уединение"
        },
        8: {
            "цвета": "Чёрный, зелёный, коричневый",
            "камни": "Гематит, обсидиан, оникс",
            "мантра": "Я привлекаю достаток",
            "практика": "Работа, стратегическое планирование, финансовое образование"
        },
        9: {
            "цвета": "Красный, малиновый, пурпурный",
            "камни": "Гранат, пироп",
            "мантра": "Я мудр и сострадателен",
            "практика": "Благотворительность, медитация, помощь другим"
        }
    }
    
    rec = recommendations.get(destiny_number, {})
    
    doc.add_paragraph(f"🎨 Благоприятные цвета: {rec.get('цвета', '')}")
    doc.add_paragraph(f"💎 Защитные камни: {rec.get('камни', '')}")
    doc.add_paragraph(f"🕉️ Ваша мантра: {rec.get('мантра', '')}")
    doc.add_paragraph(f"🧘 Рекомендуемые практики: {rec.get('практика', '')}")
    
    doc.add_page_break()
    
    # ==================== БЛАГОСЛОВЕНИЕ ====================
    
    doc.add_heading('БЛАГОСЛОВЕНИЕ', level=1)
    
    blessing = f"""
✨ Дорогой(ая) {name}! ✨

Благодарим вас за доверие и внимание к матрице вашей судьбы.

Числа, которые вас сопровождают, — это не случайность, а голос вселенной, говорящий с вами на языке символов.

Пусть этот разбор станет компасом на пути к:
🌟 Пониманию своего истинного предназначения
🌟 Раскрытию скрытых способностей
🌟 Гармонии во всех сферах жизни
🌟 Материальному благополучию
🌟 Любви и глубоким отношениям
🌟 Духовному развитию и мудрости

Помните: вы — не жертва своей судьбы, а сотворец. Числа показывают направление, но вы выбираете скорость, с которой идёте.

Иди с открытым сердцем, слушай голос чисел, и твоя жизнь превратится в подлинное произведение искусства.

🙏 С любовью и уважением,
Ваш Нумерологический Гид

✨ Дата разбора: {datetime.now().strftime("%d.%m.%Y в %H:%M")} ✨
    """
    
    doc.add_paragraph(blessing)
    
    # Сохраняем файл
    filename = f"Matrica_Sudby_{name.replace(' ', '_')}_{datetime.now().strftime('%d%m%Y')}.docx"
    filepath = f"/tmp/{filename}"
    doc.save(filepath)
    
    return filepath


def sum_to_single_digit(n: int) -> int:
    """Сводит число к одной цифре"""
    while n >= 10:
        n = sum(int(digit) for digit in str(n))
    return n


def calculate_name_number(name: str) -> int:
    """Вычисляет число имени"""
    # Используем кириллицу (русский алфавит)
    cyrillic_values = {
        'А': 1, 'Б': 2, 'В': 3, 'Г': 4, 'Д': 5, 'Е': 6, 'Ё': 7, 'Ж': 8, 'З': 9,
        'И': 1, 'Й': 2, 'К': 3, 'Л': 4, 'М': 5, 'Н': 6, 'О': 7, 'П': 8, 'Р': 9,
        'С': 1, 'Т': 2, 'У': 3, 'Ф': 4, 'Х': 5, 'Ц': 6, 'Ч': 7, 'Ш': 8, 'Щ': 9,
        'Ъ': 1, 'Ы': 2, 'Ь': 3, 'Э': 4, 'Ю': 5, 'Я': 6
    }
    
    total = 0
    for char in name.upper():
        if char in cyrillic_values:
            total += cyrillic_values[char]
    
    return sum_to_single_digit(total)


# ==================== MAIN ====================

def main() -> None:
    """Запуск бота"""
    
    # Создаём приложение
    app = Application.builder().token(BOT_TOKEN).build()
    
    # Диалог
    conv_handler = ConversationHandler(
        entry_points=[
            CommandHandler("start", start),
            MessageHandler(filters.TEXT & ~filters.COMMAND, receive_name)
        ],
        states={
            WAITING_NAME: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, receive_name),
                CommandHandler("cancel", cancel)
            ],
            WAITING_DATE: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, receive_date),
                CommandHandler("cancel", cancel)
            ],
            WAITING_METHODOLOGY: [
                MessageHandler(filters.ALL, receive_methodology),
                CommandHandler("cancel", cancel)
            ]
        },
        fallbacks=[CommandHandler("cancel", cancel)]
    )
    
    # Добавляем обработчики
    app.add_handler(conv_handler)
    app.add_handler(CommandHandler("help", help_command))
    app.add_handler(MessageHandler(filters.Regex("^(Начать|начать)"), start))
    
    # Запускаем бота
    app.run_polling()


if __name__ == '__main__':
    main()
