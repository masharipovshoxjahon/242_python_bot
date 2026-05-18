"""
Telegram Quiz Bot
=================
O'rnatish:
    pip install pyTelegramBotAPI openpyxl python-dotenv

Ishga tushirish:
    python quiz_bot.py

.env fayl:
    TELEGRAM_BOT_TOKEN=your_token_here
    ADMIN_ID=your_telegram_id_here

Excel fayl formati (birinchi varaq):
    A ustun: Savol
    B ustun: Variant A
    C ustun: Variant B
    D ustun: Variant C  (ixtiyoriy)
    E ustun: Variant D  (ixtiyoriy)
    F ustun: To'g'ri javob harfi: A, B, C yoki D
"""

import os
import io
import random
import telebot
from telebot import types
from openpyxl import load_workbook
from flask import Flask
from threading import Thread

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

# ─── Sozlamalar ────────────────────────────────────────────────────────────────

TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN", "8958925484:AAEyfBnB2PpMK-Vp2DuYPcHlzAbTVE7QPtQ")
ADMIN_ID = os.environ.get("ADMIN_ID", "6847269931")

if not TOKEN:
    print("TELEGRAM_BOT_TOKEN topilmadi! .env faylga TOKEN yozing.")
    exit(1)

bot = telebot.TeleBot(TOKEN, parse_mode=None)
app = Flask(__name__)

@app.route("/")
def home():
    return "Quiz Bot ishlayapti!"

def run_web():
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)

# ─── Ma'lumotlar ───────────────────────────────────────────────────────────────

# Admin yuklagan fayllar: {fayl_nomi: [savollar]}
quiz_files: dict[str, list[dict]] = {}

# Har bir foydalanuvchining quiz sessiyasi
# {chat_id: {"file_name": str, "questions": [...], "current": int, "score": int}}
sessions: dict[int, dict] = {}

# ─── Yordamchi funksiyalar ─────────────────────────────────────────────────────

def is_admin(user_id: int) -> bool:
    if not ADMIN_ID:
        return False
    return str(user_id) == str(ADMIN_ID)


def parse_xlsx(file_bytes: bytes) -> list[dict]:
    """xlsx faylni o'qib, savollar ro'yxatini qaytaradi."""
    wb = load_workbook(filename=io.BytesIO(file_bytes), read_only=True, data_only=True)
    ws = wb.active
    questions = []

    for row in ws.iter_rows(values_only=True):
        if not row or not row[0]:
            continue

        question = str(row[0]).strip()
        if not question:
            continue

        options = []
        for col_idx in range(1, 5):
            if col_idx < len(row) and row[col_idx]:
                val = str(row[col_idx]).strip()
                if val:
                    options.append(val)

        if len(options) < 2:
            continue

        correct_raw = ""
        if len(row) > 5 and row[5]:
            correct_raw = str(row[5]).strip().upper()

        correct_map = {"A": 0, "B": 1, "C": 2, "D": 3}
        correct_index = correct_map.get(correct_raw, 0)

        questions.append({
            "question": question,
            "options": options,
            "correct": correct_index,
        })

    wb.close()
    return questions


def shuffle_question(q: dict) -> dict:
    """
    Savol variantlarini tasodifiy aralashtiradi va yangi to'g'ri javob
    indeksini hisoblab qaytaradi.
    """
    original_options = q["options"]
    original_correct = q["correct"]
    correct_text = original_options[original_correct]

    # Indekslarni aralashtirish
    indices = list(range(len(original_options)))
    random.shuffle(indices)

    shuffled_options = [original_options[i] for i in indices]
    new_correct = shuffled_options.index(correct_text)

    return {
        "question": q["question"],
        "options": shuffled_options,
        "correct": new_correct,
    }


def make_answer_keyboard(options: list[str]) -> types.InlineKeyboardMarkup:
    """Javob variantlari uchun inline tugmalar."""
    letters = ["A", "B", "C", "D"]
    keyboard = types.InlineKeyboardMarkup(row_width=2)
    buttons = [
        types.InlineKeyboardButton(f"{letters[i]}) {opt}", callback_data=f"ans:{i}")
        for i, opt in enumerate(options)
    ]
    keyboard.add(*buttons)
    return keyboard


def make_files_keyboard() -> types.InlineKeyboardMarkup:
    """Mavjud quiz fayllarini tanlash uchun inline tugmalar."""
    keyboard = types.InlineKeyboardMarkup(row_width=1)
    for file_name in quiz_files:
        count = len(quiz_files[file_name])
        label = f"📄 {file_name}  ({count} ta savol)"
        btn = types.InlineKeyboardButton(label, callback_data=f"file:{file_name}")
        keyboard.add(btn)
    return keyboard


def send_question(chat_id: int):
    """Joriy savolni foydalanuvchiga yuboradi."""
    session = sessions.get(chat_id)
    if not session:
        return

    idx = session["current"]
    total = len(session["questions"])
    q = session["questions"][idx]
    letters = ["A", "B", "C", "D"]

    text = (
        f"📝 Savol {idx + 1}/{total}\n\n"
        f"{q['question']}\n\n"
        + "\n".join(f"{letters[i]}) {opt}" for i, opt in enumerate(q["options"]))
    )

    bot.send_message(chat_id, text, reply_markup=make_answer_keyboard(q["options"]))


# ─── Buyruqlar ─────────────────────────────────────────────────────────────────

@bot.message_handler(commands=["start"])
def cmd_start(message: types.Message):
    user_id = message.from_user.id
    chat_id = message.chat.id

    if not ADMIN_ID:
        bot.send_message(
            chat_id,
            f"Sizning Telegram ID: {user_id}\n\n"
            f"Bu raqamni ADMIN_ID sifatida .env fayliga yozing va botni qayta ishga tushiring."
        )
        return

    if is_admin(user_id):
        count = len(quiz_files)
        total_q = sum(len(q) for q in quiz_files.values())
        bot.send_message(
            chat_id,
            f"Salom, Admin! 👋\n\n"
            f"📂 Hozirda: {count} ta fayl, {total_q} ta savol yuklangan.\n\n"
            f"Admin buyruqlari:\n"
            f"📤 xlsx fayl yuboring — yangi fayl yuklash\n"
            f"/files — yuklangan fayllar ro'yxati\n"
            f"/clear — barcha fayllarni o'chirish\n\n"
            f"Foydalanuvchilar uchun:\n"
            f"/quiz — quizni boshlash (fayl tanlash)\n"
            f"/stop — quizni to'xtatish"
        )
    else:
        count = len(quiz_files)
        if count > 0:
            bot.send_message(
                chat_id,
                f"Salom! Quiz Botga xush kelibsiz! 🎉\n\n"
                f"📂 {count} ta quiz mavjud!\n\n"
                f"/quiz — quizni boshlash"
            )
        else:
            bot.send_message(
                chat_id,
                "Salom! Quiz Botga xush kelibsiz! 🎉\n\n"
                "⏳ Admin hali savollarni yuklamagan. Kuting."
            )


@bot.message_handler(commands=["files"])
def cmd_files(message: types.Message):
    if not is_admin(message.from_user.id):
        bot.send_message(message.chat.id, "Bu buyruq faqat admin uchun.")
        return

    if not quiz_files:
        bot.send_message(message.chat.id, "Hech qanday fayl yuklanmagan.")
        return

    lines = [f"📂 Yuklangan fayllar ({len(quiz_files)} ta):\n"]
    for i, (name, qs) in enumerate(quiz_files.items(), 1):
        lines.append(f"{i}. {name}  — {len(qs)} ta savol")

    bot.send_message(message.chat.id, "\n".join(lines))


@bot.message_handler(commands=["clear"])
def cmd_clear(message: types.Message):
    if not is_admin(message.from_user.id):
        bot.send_message(message.chat.id, "Bu buyruq faqat admin uchun.")
        return
    quiz_files.clear()
    sessions.clear()
    bot.send_message(message.chat.id, "Barcha fayllar va sessiyalar o'chirildi.")


@bot.message_handler(commands=["myid"])
def cmd_myid(message: types.Message):
    bot.send_message(message.chat.id, f"Sizning Telegram ID: {message.from_user.id}")


@bot.message_handler(commands=["quiz"])
def cmd_quiz(message: types.Message):
    chat_id = message.chat.id

    if chat_id in sessions:
        bot.send_message(chat_id, "Quiz allaqachon davom etmoqda! To'xtatish uchun /stop yuboring.")
        return

    if not quiz_files:
        bot.send_message(chat_id, "Hali savollar yuklanmagan. Admin yuklaguncha kuting.")
        return

    if len(quiz_files) == 1:
        file_name = next(iter(quiz_files))
        start_quiz(chat_id, file_name)
    else:
        bot.send_message(
            chat_id,
            "📂 Qaysi quiz bo'yicha ishlashni tanlang:",
            reply_markup=make_files_keyboard()
        )


@bot.message_handler(commands=["stop"])
def cmd_stop(message: types.Message):
    chat_id = message.chat.id
    if chat_id in sessions:
        sessions.pop(chat_id)
        bot.send_message(chat_id, "Quiz to'xtatildi. Qayta boshlash uchun /quiz yuboring.")
    else:
        bot.send_message(chat_id, "Faol quiz topilmadi.")


@bot.message_handler(commands=["help"])
def cmd_help(message: types.Message):
    bot.send_message(
        message.chat.id,
        "Buyruqlar:\n"
        "/quiz — quizni boshlash\n"
        "/stop — quizni to'xtatish\n"
        "/myid — Telegram ID ni ko'rish\n"
        "/help — yordam"
    )


# ─── Fayl qabul qilish (admin uchun) ───────────────────────────────────────────

@bot.message_handler(content_types=["document"])
def handle_document(message: types.Message):
    if not is_admin(message.from_user.id):
        bot.send_message(message.chat.id, "Faqat admin xlsx fayl yuklashi mumkin.")
        return

    file_name = (message.document.file_name or "").strip()
    if not (file_name.endswith(".xlsx") or file_name.endswith(".xls")):
        bot.send_message(message.chat.id, "Faqat xlsx yoki xls formatdagi fayllar qabul qilinadi.")
        return

    display_name = file_name.rsplit(".", 1)[0]

    bot.send_message(message.chat.id, f"Fayl yuklanmoqda: {file_name} ...")

    try:
        file_info = bot.get_file(message.document.file_id)
        downloaded = bot.download_file(file_info.file_path)
        questions = parse_xlsx(downloaded)

        if not questions:
            bot.send_message(
                message.chat.id,
                "Faylda savollar topilmadi.\n\n"
                "Format:\n"
                "A: Savol | B: Variant A | C: Variant B | D: Variant C | E: Variant D | F: To'g'ri javob (A/B/C/D)"
            )
            return

        is_update = display_name in quiz_files
        quiz_files[display_name] = questions

        if is_update:
            msg = f"'{display_name}' yangilandi: {len(questions)} ta savol."
        else:
            msg = (
                f"{len(questions)} ta savol muvaffaqiyatli yuklandi!\n"
                f"Fayl nomi: {display_name}\n\n"
                f"Jami yuklangan fayllar: {len(quiz_files)} ta"
            )

        bot.send_message(message.chat.id, msg)

    except Exception as e:
        print(f"Xatolik: {e}")
        bot.send_message(message.chat.id, "Faylni qayta ishlashda xatolik yuz berdi.")


# ─── Fayl tanlash ──────────────────────────────────────────────────────────────

@bot.callback_query_handler(func=lambda call: call.data.startswith("file:"))
def handle_file_select(call: types.CallbackQuery):
    chat_id = call.message.chat.id
    file_name = call.data[len("file:"):]

    if file_name not in quiz_files:
        bot.answer_callback_query(call.id, "Bu fayl topilmadi yoki o'chirilgan.")
        return

    try:
        bot.edit_message_reply_markup(chat_id, call.message.message_id, reply_markup=None)
    except Exception:
        pass

    bot.answer_callback_query(call.id)
    start_quiz(chat_id, file_name)


def start_quiz(chat_id: int, file_name: str):
    """Tanlangan fayl bo'yicha quizni boshlaydi."""
    if chat_id in sessions:
        bot.send_message(chat_id, "Quiz allaqachon davom etmoqda! /stop bilan to'xtating.")
        return

    questions = quiz_files.get(file_name)
    if not questions:
        bot.send_message(chat_id, "Bu faylda savollar topilmadi.")
        return

    # Savollar tartibini aralashtirish + har bir savolning variantlarini aralashtirish
    shuffled_questions = random.sample(questions, len(questions))
    shuffled_questions = [shuffle_question(q) for q in shuffled_questions]

    sessions[chat_id] = {
        "file_name": file_name,
        "questions": shuffled_questions,
        "current": 0,
        "score": 0,
    }

    bot.send_message(chat_id, f"Quiz boshlandi: {file_name}\nJami {len(shuffled_questions)} ta savol. Omad! 🍀")
    send_question(chat_id)


# ─── Javob tugmalarini qayta ishlash ───────────────────────────────────────────

@bot.callback_query_handler(func=lambda call: call.data.startswith("ans:"))
def handle_answer(call: types.CallbackQuery):
    chat_id = call.message.chat.id
    session = sessions.get(chat_id)

    if not session:
        bot.answer_callback_query(call.id, "Quiz topilmadi. /quiz bilan boshlang.")
        return

    try:
        answer_index = int(call.data[len("ans:"):])
    except ValueError:
        bot.answer_callback_query(call.id)
        return

    q = session["questions"][session["current"]]
    is_correct = answer_index == q["correct"]
    letters = ["A", "B", "C", "D"]
    correct_letter = letters[q["correct"]]
    correct_opt = q["options"][q["correct"]]

    if is_correct:
        session["score"] += 1
        bot.answer_callback_query(call.id, "To'g'ri! ✅")
        result_text = "✅ To'g'ri! +1 ball"
    else:
        bot.answer_callback_query(call.id, "Noto'g'ri! ❌")
        result_text = f"❌ Noto'g'ri! To'g'ri javob: {correct_letter}) {correct_opt}"

    try:
        bot.edit_message_reply_markup(chat_id, call.message.message_id, reply_markup=None)
    except Exception:
        pass

    session["current"] += 1
    bot.send_message(chat_id, result_text)

    total = len(session["questions"])

    if session["current"] >= total:
        score = session["score"]
        percent = round((score / total) * 100)

        if percent >= 90:
            emoji = "🏆"
        elif percent >= 70:
            emoji = "🎉"
        elif percent >= 50:
            emoji = "👍"
        else:
            emoji = "📚"

        bot.send_message(
            chat_id,
            f"{emoji} Quiz tugadi!\n\n"
            f"Fayl: {session['file_name']}\n"
            f"Natija: {score}/{total} ({percent}%)\n\n"
            f"Qayta urinish uchun /quiz yuboring."
        )
        sessions.pop(chat_id, None)
    else:
        send_question(chat_id)


# ─── Botni ishga tushirish ─────────────────────────────────────────────────────

if __name__ == "__main__":
    print("Quiz Bot ishga tushdi!")
    print(f"Admin ID: {ADMIN_ID or 'belgilanmagan'}")
    print("To'xtatish uchun Ctrl+C bosing.")

    # Flask serverni alohida threadda ishga tushirish
    Thread(target=run_web).start()

    # Telegram bot polling
    bot.infinity_polling(timeout=10, long_polling_timeout=5)