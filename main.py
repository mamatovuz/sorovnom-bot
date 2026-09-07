import asyncio
import logging
import os
import re
import sqlite3
from datetime import datetime
from typing import Optional

from aiogram import Bot, Dispatcher, F
from aiogram.enums import ParseMode
from aiogram.filters import Command, CommandStart
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.types import (
    BufferedInputFile,
    CallbackQuery,
    ErrorEvent,
    FSInputFile,
    InlineKeyboardButton,
    InlineKeyboardMarkup,
    Message,
)
from aiogram.client.default import DefaultBotProperties

from dotenv import load_dotenv

from io import BytesIO

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


# =========================================================
# CONFIG
# =========================================================

load_dotenv()

BOT_TOKEN = os.getenv("BOT_TOKEN", "").strip()

# Masalan:
# ADMIN_IDS=123456789,987654321
ADMIN_IDS = {
    int(x.strip())
    for x in os.getenv("ADMIN_IDS", "").split(",")
    if x.strip().isdigit()
}

DB_PATH = os.getenv("DB_PATH", "students.db")

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler("bot.log", encoding="utf-8"),
    ],
)
logger = logging.getLogger(__name__)

if not BOT_TOKEN:
    raise SystemExit("❌ BOT_TOKEN topilmadi! .env faylini to'ldiring.")


# =========================================================
# BOT
# =========================================================

bot = Bot(
    token=BOT_TOKEN,
    default=DefaultBotProperties(parse_mode=ParseMode.HTML)
)

dp = Dispatcher()


# =========================================================
# CLASSES
# =========================================================

CLASSES = [
    "1-A", "1-B",
    "2-A", "2-B",
    "3-A", "3-B",
    "4-A", "4-B",
    "5-A", "5-B",
    "6-A", "6-B",
    "7-A", "7-B",
    "8-A", "8-B",
    "9-A", "9-B",
    "10-A", "10-B",
    "11-A", "11-B",
]


# =========================================================
# FSM STATES
# =========================================================

class Registration(StatesGroup):
    choosing_class = State()
    full_name = State()
    birth_date = State()
    document = State()
    phone = State()
    address = State()
    confirmation = State()


class UserEdit(StatesGroup):
    waiting_value = State()


class AdminEdit(StatesGroup):
    waiting_value = State()


class AdminSearch(StatesGroup):
    waiting_query = State()


class Broadcast(StatesGroup):
    waiting_message = State()


# =========================================================
# DATABASE
# =========================================================

def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    conn = get_db()
    cur = conn.cursor()

    cur.execute("""
        CREATE TABLE IF NOT EXISTS students (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            tg_user_id INTEGER NOT NULL,
            class_name TEXT NOT NULL,
            full_name TEXT NOT NULL,
            birth_date TEXT NOT NULL,
            doc_type TEXT NOT NULL,
            doc_value TEXT NOT NULL,
            phone TEXT NOT NULL,
            address TEXT DEFAULT '',
            created_at TEXT NOT NULL
        )
    """)

    # Eski database uchun migration
    cur.execute("PRAGMA table_info(students)")
    columns = [row["name"] for row in cur.fetchall()]

    if "address" not in columns:
        cur.execute(
            "ALTER TABLE students ADD COLUMN address TEXT DEFAULT ''"
        )

    # Telegram ID unique bo'lmasligi kerak.
    # Shuning uchun tg_user_id ga UNIQUE INDEX yaratmaymiz.

    # Qidiruvlarni tezlashtirish
    cur.execute("""
        CREATE INDEX IF NOT EXISTS idx_students_tg_user_id
        ON students(tg_user_id)
    """)

    cur.execute("""
        CREATE INDEX IF NOT EXISTS idx_students_doc_value
        ON students(doc_value)
    """)

    conn.commit()
    conn.close()


# =========================================================
# DATABASE FUNCTIONS
# =========================================================

def db_get_student(student_id: int) -> Optional[sqlite3.Row]:
    conn = get_db()

    row = conn.execute(
        "SELECT * FROM students WHERE id = ?",
        (student_id,)
    ).fetchone()

    conn.close()
    return row


def db_get_user_students(tg_user_id: int):
    conn = get_db()

    rows = conn.execute("""
        SELECT *
        FROM students
        WHERE tg_user_id = ?
        ORDER BY id DESC
    """, (tg_user_id,)).fetchall()

    conn.close()
    return rows


def db_get_all_students():
    conn = get_db()

    rows = conn.execute("""
        SELECT *
        FROM students
        ORDER BY id DESC
    """).fetchall()

    conn.close()
    return rows


def db_get_students_by_class(class_name: str):
    conn = get_db()

    rows = conn.execute("""
        SELECT *
        FROM students
        WHERE class_name = ?
        ORDER BY full_name COLLATE NOCASE
    """, (class_name,)).fetchall()

    conn.close()
    return rows


def db_search_students(query: str):
    conn = get_db()

    like = f"%{query}%"

    rows = conn.execute("""
        SELECT *
        FROM students
        WHERE
            full_name LIKE ?
            OR doc_value LIKE ?
            OR phone LIKE ?
            OR class_name LIKE ?
            OR address LIKE ?
        ORDER BY id DESC
        LIMIT 50
    """, (
        like,
        like,
        like,
        like,
        like
    )).fetchall()

    conn.close()
    return rows


def db_doc_exists(doc_value: str, exclude_id: Optional[int] = None) -> bool:
    conn = get_db()

    if exclude_id is None:
        row = conn.execute("""
            SELECT id
            FROM students
            WHERE doc_value = ?
            LIMIT 1
        """, (doc_value,)).fetchone()
    else:
        row = conn.execute("""
            SELECT id
            FROM students
            WHERE doc_value = ?
              AND id != ?
            LIMIT 1
        """, (doc_value, exclude_id)).fetchone()

    conn.close()

    return row is not None


def db_get_by_doc(doc_value: str, exclude_id: Optional[int] = None):
    """Ushbu JSHSHIR/metrika kimga tegishli ekanini qaytaradi (yoki None)."""
    conn = get_db()

    if exclude_id is None:
        row = conn.execute(
            "SELECT * FROM students WHERE doc_value = ? LIMIT 1",
            (doc_value,)
        ).fetchone()
    else:
        row = conn.execute(
            "SELECT * FROM students WHERE doc_value = ? AND id != ? LIMIT 1",
            (doc_value, exclude_id)
        ).fetchone()

    conn.close()
    return row


def db_create_student(
    tg_user_id: int,
    class_name: str,
    full_name: str,
    birth_date: str,
    doc_type: str,
    doc_value: str,
    phone: str,
    address: str
) -> int:

    conn = get_db()

    cur = conn.execute("""
        INSERT INTO students (
            tg_user_id,
            class_name,
            full_name,
            birth_date,
            doc_type,
            doc_value,
            phone,
            address,
            created_at
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        tg_user_id,
        class_name,
        full_name,
        birth_date,
        doc_type,
        doc_value,
        phone,
        address,
        datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    ))

    student_id = cur.lastrowid

    conn.commit()
    conn.close()

    return student_id


def db_update_field(
    student_id: int,
    field: str,
    value
):
    allowed_fields = {
        "class_name",
        "full_name",
        "birth_date",
        "doc_type",
        "doc_value",
        "phone",
        "address",
    }

    if field not in allowed_fields:
        raise ValueError("Noto'g'ri field")

    conn = get_db()

    conn.execute(
        f"UPDATE students SET {field} = ? WHERE id = ?",
        (value, student_id)
    )

    conn.commit()
    conn.close()


def db_delete_student(student_id: int):
    conn = get_db()

    conn.execute(
        "DELETE FROM students WHERE id = ?",
        (student_id,)
    )

    conn.commit()
    conn.close()


def db_get_total():
    conn = get_db()

    row = conn.execute(
        "SELECT COUNT(*) AS count FROM students"
    ).fetchone()

    conn.close()

    return row["count"]


def db_get_today_count():
    today = datetime.now().strftime("%Y-%m-%d")

    conn = get_db()

    row = conn.execute("""
        SELECT COUNT(*) AS count
        FROM students
        WHERE created_at LIKE ?
    """, (f"{today}%",)).fetchone()

    conn.close()

    return row["count"]


def db_get_unique_users():
    conn = get_db()

    rows = conn.execute("""
        SELECT DISTINCT tg_user_id
        FROM students
    """).fetchall()

    conn.close()

    return [row["tg_user_id"] for row in rows]


# =========================================================
# HELPERS
# =========================================================

def is_admin(user_id: int) -> bool:
    return user_id in ADMIN_IDS


def normalize_phone(text: str) -> str:
    """Telefon raqamini +998XXXXXXXXX ko'rinishiga keltiradi (imkon bo'lsa)."""
    digits = re.sub(r"\D", "", text)

    if digits.startswith("998") and len(digits) == 12:
        return "+" + digits
    if len(digits) == 9:  # 901234567
        return "+998" + digits
    if digits.startswith("8") and len(digits) == 10:  # 8XXXXXXXXXX
        return "+99" + digits
    if 9 <= len(digits) <= 15:
        return "+" + digits

    # Normallashtira olmasak — asl (faqat raqam va +) ko'rinishi
    return re.sub(r"[^\d+]", "", text.strip())


def validate_phone(text: str) -> bool:
    digits = re.sub(r"\D", "", text)
    return 9 <= len(digits) <= 15


def valid_full_name(text: str) -> bool:
    """Kamida 2 so'z va faqat harflar (lotin/kirill) bo'lsin."""
    text = text.strip()
    if len(text.split()) < 2:
        return False
    return bool(
        re.fullmatch(r"[A-Za-zА-Яа-яЁёЎўҚқҒғҲҳ'ʼ`\-\s\.]{5,100}", text)
    )


def normalize_date(text: str):
    """
    Sanani tekshirib normallashtiradi: 'kun.oy.yil' -> 'DD.MM.YYYY' yoki None.
    - / . - va bo'sh joy ajratkichlarini qabul qiladi
    - kelajakdagi sana bo'lmasin, yil 1990..hozir oralig'ida
    """
    cleaned = text.strip().replace("/", ".").replace("-", ".").replace(" ", "")
    try:
        d = datetime.strptime(cleaned, "%d.%m.%Y")
    except ValueError:
        return None

    now = datetime.now()
    if d > now:
        return None
    if not (1990 <= d.year <= now.year):
        return None

    return d.strftime("%d.%m.%Y")


def validate_date(date_text: str) -> bool:
    return normalize_date(date_text) is not None


def extract_document(text: str):
    """
    JSHSHIR:
    14 ta raqam

    Metrika:
    7-10 ta raqam yoki harf/raqam ko'rinishida bo'lishi mumkin.
    """

    value = text.strip().upper()

    digits = re.sub(r"\D", "", value)

    # JSHSHIR
    if len(digits) == 14:
        return "JSHSHIR", digits

    # Metrika (seriya + raqam, masalan: II-AB1234567)
    cleaned = re.sub(r"[\s\-]", "", value)

    if (
        6 <= len(cleaned) <= 20
        and re.fullmatch(r"[A-Z0-9]+", cleaned)
        and any(c.isdigit() for c in cleaned)
    ):
        return "Metrika", cleaned

    return None, None


def validate_jshshir_birth_date(
    doc_type: str,
    doc_value: str,
    birth_date: str
) -> bool:

    if doc_type != "JSHSHIR":
        return True

    try:
        date = datetime.strptime(
            birth_date,
            "%d.%m.%Y"
        )

        # Berilgan oldingi mantiq:
        # JSHSHIRning 2-7 indekslari DDMMYY
        jsh_date = doc_value[1:7]

        expected = date.strftime("%d%m%y")

        return jsh_date == expected

    except Exception:
        return False


# =========================================================
# KEYBOARDS
# =========================================================

def classes_keyboard(prefix: str = "reg_class"):
    rows = []

    for i in range(0, len(CLASSES), 2):
        row = []

        for cls in CLASSES[i:i + 2]:
            row.append(
                InlineKeyboardButton(
                    text=cls,
                    callback_data=f"{prefix}:{cls}"
                )
            )

        rows.append(row)

    return InlineKeyboardMarkup(inline_keyboard=rows)


def main_menu_keyboard():
    buttons = [
        [
            InlineKeyboardButton(
                text="➕ Yangi o'quvchi qo'shish",
                callback_data="new_student"
            )
        ]
    ]

    return InlineKeyboardMarkup(inline_keyboard=buttons)


def reg_back_kb(target: str):
    """Ro'yxatdan o'tishda oldingi bosqichga qaytish tugmasi."""
    return InlineKeyboardMarkup(
        inline_keyboard=[
            [
                InlineKeyboardButton(
                    text="⬅️ Orqaga",
                    callback_data=f"reg_back:{target}"
                )
            ]
        ]
    )


# ---- Ro'yxatdan o'tish savol matnlari (orqaga tugmasi bilan) ----

def reg_prompt_full_name(class_name: str):
    text = (
        f"📚 Sinf: <b>{class_name}</b>\n\n"
        "1️⃣ 👤 O'quvchining <b>F.I.Sh.</b>ini kiriting:\n\n"
        "Masalan: <i>Aliyev Vali Vohid o'g'li</i>"
    )
    return text, reg_back_kb("class")


def reg_prompt_birth_date():
    text = (
        "2️⃣ 🎂 <b>Tug'ilgan sanani</b> kiriting.\n\n"
        "Format: <b>DD.MM.YYYY</b>\n"
        "Masalan: <code>05.12.2008</code>"
    )
    return text, reg_back_kb("full_name")


def reg_prompt_document():
    text = (
        "3️⃣ 🪪 <b>JSHSHIR yoki Metrika</b>ni kiriting.\n\n"
        "• JSHSHIR — 14 ta raqam\n"
        "• Metrika — seriya raqami (masalan: <code>II-AB1234567</code>)"
    )
    return text, reg_back_kb("birth_date")


def reg_prompt_phone():
    text = (
        "4️⃣ 📱 <b>Telefon raqamini</b> kiriting.\n\n"
        "Masalan: <code>+998901234567</code>"
    )
    return text, reg_back_kb("document")


def reg_prompt_address():
    text = (
        "5️⃣ 🏠 <b>Yashash manzilini</b> kiriting.\n\n"
        "Masalan: <i>Toshkent sh., Uzun ko'cha, 23-uy</i>"
    )
    return text, reg_back_kb("phone")


def user_students_keyboard(students):
    rows = []

    for student in students:
        name = student["full_name"]

        if len(name) > 25:
            name = name[:22] + "..."

        rows.append([
            InlineKeyboardButton(
                text=f"✏️ {name} — {student['class_name']}",
                callback_data=f"edit_student:{student['id']}"
            )
        ])

    rows.append([
        InlineKeyboardButton(
            text="➕ Yangi o'quvchi qo'shish",
            callback_data="new_student"
        )
    ])

    return InlineKeyboardMarkup(inline_keyboard=rows)


def student_edit_keyboard(student_id: int, admin: bool = False):
    prefix = "admin_edit_field" if admin else "edit_field"

    return InlineKeyboardMarkup(
        inline_keyboard=[
            [
                InlineKeyboardButton(
                    text="📚 Sinf",
                    callback_data=f"{prefix}:{student_id}:class_name"
                ),
                InlineKeyboardButton(
                    text="👤 F.I.Sh.",
                    callback_data=f"{prefix}:{student_id}:full_name"
                ),
            ],
            [
                InlineKeyboardButton(
                    text="🎂 Tug'ilgan sana",
                    callback_data=f"{prefix}:{student_id}:birth_date"
                )
            ],
            [
                InlineKeyboardButton(
                    text="🪪 JSHSHIR / Metrika",
                    callback_data=f"{prefix}:{student_id}:document"
                )
            ],
            [
                InlineKeyboardButton(
                    text="📱 Telefon",
                    callback_data=f"{prefix}:{student_id}:phone"
                ),
                InlineKeyboardButton(
                    text="🏠 Manzil",
                    callback_data=f"{prefix}:{student_id}:address"
                ),
            ],
            [
                InlineKeyboardButton(
                    text="⬅️ Orqaga",
                    callback_data=(
                        f"admin_back_students"
                        if admin
                        else "back_my_students"
                    )
                )
            ]
        ]
    )


def admin_search_result_keyboard(student_id: int):
    return InlineKeyboardMarkup(
        inline_keyboard=[
            [
                InlineKeyboardButton(
                    text="✏️ Tahrirlash",
                    callback_data=f"admin_select_student:{student_id}"
                ),
                InlineKeyboardButton(
                    text="🗑 O'chirish",
                    callback_data=f"admin_delete:{student_id}"
                )
            ]
        ]
    )


def confirmation_keyboard():
    return InlineKeyboardMarkup(
        inline_keyboard=[
            [
                InlineKeyboardButton(
                    text="✅ Tasdiqlash",
                    callback_data="confirm_registration"
                )
            ],
            [
                InlineKeyboardButton(
                    text="✏️ Qayta kiritish",
                    callback_data="restart_registration"
                )
            ]
        ]
    )


def admin_menu_keyboard():
    return InlineKeyboardMarkup(
        inline_keyboard=[
            [
                InlineKeyboardButton(
                    text="📊 Statistika",
                    callback_data="admin_stats"
                )
            ],
            [
                InlineKeyboardButton(
                    text="🔎 Qidirish",
                    callback_data="admin_search"
                )
            ],
            [
                InlineKeyboardButton(
                    text="📢 Broadcast",
                    callback_data="admin_broadcast"
                )
            ],
            [
                InlineKeyboardButton(
                    text="📥 Excel",
                    callback_data="admin_excel"
                )
            ],
            [
                InlineKeyboardButton(
                    text="🗄 Zaxira (backup)",
                    callback_data="admin_backup"
                )
            ]
        ]
    )


# =========================================================
# FORMATTING
# =========================================================

def format_student(student) -> str:
    address = student["address"] or "—"

    return (
        f"👤 <b>{student['full_name']}</b>\n\n"
        f"📚 <b>Sinf:</b> {student['class_name']}\n"
        f"🎂 <b>Tug'ilgan sana:</b> {student['birth_date']}\n"
        f"🪪 <b>{student['doc_type']}:</b> {student['doc_value']}\n"
        f"📱 <b>Telefon:</b> {student['phone']}\n"
        f"🏠 <b>Manzil:</b> {address}\n\n"
        f"🕐 <b>Ro'yxatdan o'tgan:</b> {student['created_at']}"
    )


def format_registration_preview(data: dict) -> str:
    address = data.get("address") or "—"

    return (
        "📋 <b>Ma'lumotlarni tekshiring</b>\n\n"
        f"📚 <b>Sinf:</b> {data['class_name']}\n"
        f"👤 <b>F.I.Sh.:</b> {data['full_name']}\n"
        f"🎂 <b>Tug'ilgan sana:</b> {data['birth_date']}\n"
        f"🪪 <b>{data['doc_type']}:</b> {data['doc_value']}\n"
        f"📱 <b>Telefon:</b> {data['phone']}\n"
        f"🏠 <b>Manzil:</b> {address}\n\n"
        "Ma'lumotlar to'g'rimi?"
    )


# =========================================================
# /START
# =========================================================

@dp.message(CommandStart())
async def cmd_start(message: Message, state: FSMContext):
    await state.clear()

    students = db_get_user_students(message.from_user.id)

    if not students:
        await message.answer(
            "👋 Assalomu alaykum!\n\n"
            "O'quvchi ma'lumotlarini ro'yxatdan o'tkazish uchun "
            "sinfni tanlang.",
            reply_markup=classes_keyboard()
        )

        await state.set_state(Registration.choosing_class)
        return

    text = (
        f"👋 Assalomu alaykum!\n\n"
        f"👥 Sizning Telegram akkauntingiz orqali "
        f"<b>{len(students)} ta o'quvchi</b> ro'yxatdan o'tgan.\n\n"
        "Kerakli o'quvchini tanlab, uning ma'lumotlarini "
        "tahrirlashingiz mumkin."
    )

    await message.answer(
        text,
        reply_markup=user_students_keyboard(students)
    )


# =========================================================
# /CANCEL  va  /HELP
# =========================================================

@dp.message(Command("cancel"))
async def cmd_cancel(message: Message, state: FSMContext):
    current = await state.get_state()
    await state.clear()

    if current is None:
        await message.answer(
            "ℹ️ Hozir bekor qiladigan jarayon yo'q.\n"
            "Boshlash uchun /start bosing."
        )
    else:
        await message.answer(
            "❌ Jarayon bekor qilindi.\n"
            "Qaytadan boshlash uchun /start bosing."
        )


@dp.message(Command("help"))
async def cmd_help(message: Message):
    text = (
        "ℹ️ <b>Yordam</b>\n\n"
        "Bu bot orqali o'quvchilarni ro'yxatdan o'tkazasiz. "
        "Bitta Telegram akkaunt orqali bir nechta o'quvchini "
        "ro'yxatdan o'tkazishingiz mumkin.\n\n"
        "• /start — boshlash / o'quvchilaringizni boshqarish\n"
        "• /cancel — joriy jarayonni bekor qilish\n"
        "• /help — ushbu yordam\n\n"
    )
    if is_admin(message.from_user.id):
        text += (
            "🛠 <b>Admin buyruqlari:</b>\n"
            "• /admin — admin panel\n"
            "• /stats — statistika\n"
            "• /search yoki /search &lt;so'z&gt; — qidiruv\n\n"
        )
    text += (
        "━━━━━━━━━━━━━━━\n"
        "👨‍💻 <b>Loyiha asoschisi:</b> @Mamatov_ads"
    )
    await message.answer(text)


# =========================================================
# NEW STUDENT
# =========================================================

@dp.callback_query(F.data == "new_student")
async def new_student(callback: CallbackQuery, state: FSMContext):
    await callback.answer()

    await state.clear()

    await state.set_state(Registration.choosing_class)

    await callback.message.edit_text(
        "➕ <b>Yangi o'quvchi qo'shish</b>\n\n"
        "📚 Sinfni tanlang:",
        reply_markup=classes_keyboard()
    )


# =========================================================
# REGISTRATION — CLASS
# =========================================================

@dp.callback_query(
    Registration.choosing_class,
    F.data.startswith("reg_class:")
)
async def registration_class(
    callback: CallbackQuery,
    state: FSMContext
):
    await callback.answer()

    class_name = callback.data.split(":", 1)[1]

    await state.update_data(
        class_name=class_name
    )

    await state.set_state(Registration.full_name)

    text, kb = reg_prompt_full_name(class_name)
    await callback.message.edit_text(text, reply_markup=kb)


# =========================================================
# REGISTRATION — ORQAGA (⬅️)
# =========================================================

@dp.callback_query(F.data.startswith("reg_back:"))
async def registration_back(
    callback: CallbackQuery,
    state: FSMContext
):
    await callback.answer()

    target = callback.data.split(":", 1)[1]
    data = await state.get_data()

    if target == "class":
        await state.set_state(Registration.choosing_class)
        await callback.message.edit_text(
            "📚 Sinfni tanlang:",
            reply_markup=classes_keyboard()
        )

    elif target == "full_name":
        await state.set_state(Registration.full_name)
        text, kb = reg_prompt_full_name(data.get("class_name", "—"))
        await callback.message.edit_text(text, reply_markup=kb)

    elif target == "birth_date":
        await state.set_state(Registration.birth_date)
        text, kb = reg_prompt_birth_date()
        await callback.message.edit_text(text, reply_markup=kb)

    elif target == "document":
        await state.set_state(Registration.document)
        text, kb = reg_prompt_document()
        await callback.message.edit_text(text, reply_markup=kb)

    elif target == "phone":
        await state.set_state(Registration.phone)
        text, kb = reg_prompt_phone()
        await callback.message.edit_text(text, reply_markup=kb)


# =========================================================
# REGISTRATION — FULL NAME
# =========================================================

@dp.message(Registration.full_name)
async def registration_full_name(
    message: Message,
    state: FSMContext
):
    if not message.text:
        await message.answer(
            "⚠️ Iltimos, F.I.Sh.ni matn ko'rinishida yozing."
        )
        return

    full_name = " ".join(message.text.strip().split())

    if not valid_full_name(full_name):
        await message.answer(
            "❌ F.I.Sh. noto'g'ri.\n"
            "Kamida familiya va ism, faqat harflardan iborat bo'lsin.\n"
            "Masalan: <i>Aliyev Vali Vohid o'g'li</i>"
        )
        return

    await state.update_data(
        full_name=full_name
    )

    await state.set_state(Registration.birth_date)

    text, kb = reg_prompt_birth_date()
    await message.answer(text, reply_markup=kb)


# =========================================================
# REGISTRATION — BIRTH DATE
# =========================================================

@dp.message(Registration.birth_date)
async def registration_birth_date(
    message: Message,
    state: FSMContext
):
    if not message.text:
        await message.answer(
            "⚠️ Iltimos, sanani matn ko'rinishida yozing.\n"
            "Masalan: <code>05.12.2008</code>"
        )
        return

    birth_date = normalize_date(message.text)

    if not birth_date:
        await message.answer(
            "❌ Sana noto'g'ri.\n\n"
            "DD.MM.YYYY formatida, kelajakda emas va 1990-yildan keyin kiriting.\n"
            "Masalan: <code>05.12.2008</code>"
        )
        return

    await state.update_data(
        birth_date=birth_date
    )

    await state.set_state(Registration.document)

    text, kb = reg_prompt_document()
    await message.answer(text, reply_markup=kb)


# =========================================================
# REGISTRATION — DOCUMENT
# =========================================================

@dp.message(Registration.document)
async def registration_document(
    message: Message,
    state: FSMContext
):
    if not message.text:
        await message.answer(
            "⚠️ Iltimos, JSHSHIR yoki Metrikani matn ko'rinishida yozing."
        )
        return

    doc_type, doc_value = extract_document(message.text)

    if not doc_type:
        await message.answer(
            "❌ JSHSHIR yoki Metrika formati noto'g'ri.\n\n"
            "JSHSHIR 14 ta raqamdan iborat bo'lishi kerak."
        )
        return

    # JSHSHIR birth date check
    data = await state.get_data()

    if not validate_jshshir_birth_date(
        doc_type,
        doc_value,
        data["birth_date"]
    ):
        await message.answer(
            "❌ JSHSHIRdagi tug'ilgan sana "
            "siz kiritgan tug'ilgan sana bilan mos kelmadi.\n\n"
            "Iltimos, ma'lumotlarni tekshirib qayta kiriting."
        )
        return

    # Duplicate check
    if db_doc_exists(doc_value):
        await message.answer(
            "❌ Bu JSHSHIR/Metrika bilan o'quvchi "
            "allaqachon ro'yxatdan o'tgan.\n\n"
            "Boshqa hujjat ma'lumotini kiriting."
        )
        return

    await state.update_data(
        doc_type=doc_type,
        doc_value=doc_value
    )

    await state.set_state(Registration.phone)

    text, kb = reg_prompt_phone()
    await message.answer(text, reply_markup=kb)


# =========================================================
# REGISTRATION — PHONE
# =========================================================

@dp.message(Registration.phone)
async def registration_phone(
    message: Message,
    state: FSMContext
):
    if not message.text:
        await message.answer(
            "⚠️ Iltimos, telefon raqamini matn ko'rinishida yozing.\n"
            "Masalan: <code>+998901234567</code>"
        )
        return

    phone = normalize_phone(message.text)

    if not validate_phone(phone):
        await message.answer(
            "❌ Telefon raqami noto'g'ri.\n\n"
            "Masalan: <code>+998901234567</code>"
        )
        return

    await state.update_data(
        phone=phone
    )

    await state.set_state(Registration.address)

    text, kb = reg_prompt_address()
    await message.answer(text, reply_markup=kb)


# =========================================================
# REGISTRATION — ADDRESS
# =========================================================

@dp.message(Registration.address)
async def registration_address(
    message: Message,
    state: FSMContext
):
    if not message.text:
        await message.answer(
            "⚠️ Iltimos, manzilni matn ko'rinishida yozing."
        )
        return

    address = message.text.strip()

    if len(address) < 3:
        await message.answer(
            "❌ Manzil juda qisqa.\n"
            "Iltimos, to'liqroq manzil kiriting."
        )
        return

    await state.update_data(
        address=address
    )

    data = await state.get_data()

    await state.set_state(Registration.confirmation)

    await message.answer(
        format_registration_preview(data),
        reply_markup=confirmation_keyboard()
    )


# =========================================================
# REGISTRATION — CONFIRM
# =========================================================

@dp.callback_query(
    Registration.confirmation,
    F.data == "confirm_registration"
)
async def confirm_registration(
    callback: CallbackQuery,
    state: FSMContext
):
    await callback.answer()

    data = await state.get_data()

    student_id = db_create_student(
        tg_user_id=callback.from_user.id,
        class_name=data["class_name"],
        full_name=data["full_name"],
        birth_date=data["birth_date"],
        doc_type=data["doc_type"],
        doc_value=data["doc_value"],
        phone=data["phone"],
        address=data["address"]
    )

    student = db_get_student(student_id)

    await state.clear()

    await callback.message.edit_text(
        "✅ <b>O'quvchi muvaffaqiyatli ro'yxatdan o'tkazildi!</b>\n\n"
        + format_student(student),
        reply_markup=student_edit_keyboard(student_id)
    )

    # Adminlarga xabar
    u = callback.from_user
    username = f"@{u.username}" if u.username else "—"
    admin_text = (
        "🆕 <b>Yangi o'quvchi ro'yxatdan o'tdi</b>\n\n"
        + format_student(student)
        + f"\n\n👤 Yuboruvchi: {u.full_name} ({username})"
        + f"\n🆔 Telegram ID: <code>{u.id}</code>"
    )

    for admin_id in ADMIN_IDS:
        try:
            await bot.send_message(
                admin_id,
                admin_text,
                reply_markup=student_edit_keyboard(
                    student_id,
                    admin=True
                )
            )
        except Exception as e:
            logging.warning(
                f"Admin notification error: {admin_id}: {e}"
            )


# =========================================================
# REGISTRATION — RESTART
# =========================================================

@dp.callback_query(
    Registration.confirmation,
    F.data == "restart_registration"
)
async def restart_registration(
    callback: CallbackQuery,
    state: FSMContext
):
    await callback.answer()

    await state.clear()
    await state.set_state(Registration.choosing_class)

    await callback.message.edit_text(
        "🔄 Ma'lumotlarni qaytadan kiriting.\n\n"
        "📚 Sinfni tanlang:",
        reply_markup=classes_keyboard()
    )


# =========================================================
# USER STUDENT LIST
# =========================================================

@dp.callback_query(F.data == "back_my_students")
async def back_my_students(
    callback: CallbackQuery,
    state: FSMContext
):
    await callback.answer()

    await state.clear()

    students = db_get_user_students(
        callback.from_user.id
    )

    if not students:
        await callback.message.edit_text(
            "Sizda hali ro'yxatdan o'tgan o'quvchi yo'q.",
            reply_markup=main_menu_keyboard()
        )
        return

    await callback.message.edit_text(
        f"👥 Sizning Telegram akkauntingiz orqali "
        f"<b>{len(students)} ta o'quvchi</b> ro'yxatdan o'tgan.\n\n"
        "Kerakli o'quvchini tanlang:",
        reply_markup=user_students_keyboard(students)
    )


# =========================================================
# USER — SELECT STUDENT
# =========================================================

@dp.callback_query(F.data.startswith("edit_student:"))
async def select_user_student(
    callback: CallbackQuery,
    state: FSMContext
):
    await callback.answer()

    try:
        student_id = int(
            callback.data.split(":")[1]
        )
    except (ValueError, IndexError):
        return

    student = db_get_student(student_id)

    if not student:
        await callback.message.edit_text(
            "❌ O'quvchi topilmadi."
        )
        return

    # MUHIM:
    # Faqat shu Telegram akkauntiga tegishli student
    if student["tg_user_id"] != callback.from_user.id:
        await callback.message.answer(
            "❌ Bu o'quvchiga kirish huquqingiz yo'q."
        )
        return

    await state.clear()

    await callback.message.edit_text(
        "👤 <b>O'quvchi ma'lumotlari</b>\n\n"
        + format_student(student)
        + "\n\n✏️ Qaysi ma'lumotni tahrirlash kerak?",
        reply_markup=student_edit_keyboard(student_id)
    )


# =========================================================
# USER — EDIT FIELD
# =========================================================

FIELD_NAMES = {
    "class_name": "📚 Sinf",
    "full_name": "👤 F.I.Sh.",
    "birth_date": "🎂 Tug'ilgan sana",
    "document": "🪪 JSHSHIR / Metrika",
    "phone": "📱 Telefon",
    "address": "🏠 Manzil",
}


@dp.callback_query(F.data.startswith("edit_field:"))
async def user_edit_field(
    callback: CallbackQuery,
    state: FSMContext
):
    await callback.answer()

    parts = callback.data.split(":")

    if len(parts) != 3:
        return

    try:
        student_id = int(parts[1])
    except ValueError:
        return

    field = parts[2]

    student = db_get_student(student_id)

    if not student:
        await callback.message.answer(
            "❌ O'quvchi topilmadi."
        )
        return

    # Ownership check
    if student["tg_user_id"] != callback.from_user.id:
        await callback.message.answer(
            "❌ Bu o'quvchini tahrirlash huquqingiz yo'q."
        )
        return

    if field not in FIELD_NAMES:
        return

    await state.set_state(
        UserEdit.waiting_value
    )

    await state.update_data(
        student_id=student_id,
        field=field
    )

    # Sinf alohida keyboard
    if field == "class_name":
        await callback.message.edit_text(
            "📚 Yangi sinfni tanlang:",
            reply_markup=classes_keyboard(
                prefix=f"edit_class:{student_id}"
            )
        )
        return

    prompts = {
        "full_name":
            "👤 Yangi F.I.Sh.ni kiriting:",

        "birth_date":
            "🎂 Yangi tug'ilgan sanani kiriting.\n\n"
            "Format: <b>DD.MM.YYYY</b>",

        "document":
            "🪪 Yangi JSHSHIR yoki Metrikani kiriting:",

        "phone":
            "📱 Yangi telefon raqamini kiriting:",

        "address":
            "🏠 Yangi manzilni kiriting:",
    }

    await callback.message.edit_text(
        prompts[field]
    )


# =========================================================
# USER — EDIT CLASS
# =========================================================

@dp.callback_query(F.data.startswith("edit_class:"))
async def user_edit_class(
    callback: CallbackQuery,
    state: FSMContext
):
    await callback.answer()

    parts = callback.data.split(":")

    if len(parts) != 3:
        return

    try:
        student_id = int(parts[1])
    except ValueError:
        return

    class_name = parts[2]

    student = db_get_student(student_id)

    if not student:
        return

    if student["tg_user_id"] != callback.from_user.id:
        await callback.message.answer(
            "❌ Ruxsat yo'q."
        )
        return

    db_update_field(
        student_id,
        "class_name",
        class_name
    )

    await state.clear()

    student = db_get_student(student_id)

    await callback.message.edit_text(
        "✅ <b>Sinf yangilandi!</b>\n\n"
        + format_student(student)
        + "\n\n✏️ Yana qaysi ma'lumotni tahrirlash kerak?",
        reply_markup=student_edit_keyboard(student_id)
    )


# =========================================================
# USER — EDIT TEXT VALUE
# =========================================================

@dp.message(UserEdit.waiting_value)
async def user_edit_value(
    message: Message,
    state: FSMContext
):
    data = await state.get_data()

    student_id = data.get("student_id")
    field = data.get("field")

    if not student_id or not field:
        await state.clear()
        return

    student = db_get_student(student_id)

    if not student:
        await state.clear()
        await message.answer(
            "❌ O'quvchi topilmadi."
        )
        return

    if student["tg_user_id"] != message.from_user.id:
        await state.clear()
        await message.answer(
            "❌ Ruxsat yo'q."
        )
        return

    if not message.text:
        await message.answer(
            "⚠️ Iltimos, yangi qiymatni matn ko'rinishida yuboring."
        )
        return

    text = message.text.strip()

    # -----------------------------------------------------
    # FULL NAME
    # -----------------------------------------------------

    if field == "full_name":

        if len(text) < 5:
            await message.answer(
                "❌ F.I.Sh. juda qisqa."
            )
            return

        db_update_field(
            student_id,
            "full_name",
            text
        )

    # -----------------------------------------------------
    # BIRTH DATE
    # -----------------------------------------------------

    elif field == "birth_date":

        if not validate_date(text):
            await message.answer(
                "❌ Sana noto'g'ri.\n"
                "Format: DD.MM.YYYY"
            )
            return

        # JSHSHIR bo'lsa tug'ilgan sana mosligini tekshiramiz
        if not validate_jshshir_birth_date(
            student["doc_type"],
            student["doc_value"],
            text
        ):
            await message.answer(
                "❌ Tug'ilgan sana JSHSHIR bilan mos kelmadi."
            )
            return

        db_update_field(
            student_id,
            "birth_date",
            text
        )

    # -----------------------------------------------------
    # DOCUMENT
    # -----------------------------------------------------

    elif field == "document":

        doc_type, doc_value = extract_document(text)

        if not doc_type:
            await message.answer(
                "❌ JSHSHIR/Metrika noto'g'ri."
            )
            return

        if not validate_jshshir_birth_date(
            doc_type,
            doc_value,
            student["birth_date"]
        ):
            await message.answer(
                "❌ JSHSHIRdagi tug'ilgan sana "
                "o'quvchining tug'ilgan sanasi bilan mos kelmadi."
            )
            return

        if db_doc_exists(
            doc_value,
            exclude_id=student_id
        ):
            await message.answer(
                "❌ Bu JSHSHIR/Metrika boshqa o'quvchiga tegishli."
            )
            return

        db_update_field(
            student_id,
            "doc_type",
            doc_type
        )

        db_update_field(
            student_id,
            "doc_value",
            doc_value
        )

    # -----------------------------------------------------
    # PHONE
    # -----------------------------------------------------

    elif field == "phone":

        phone = normalize_phone(text)

        if not validate_phone(phone):
            await message.answer(
                "❌ Telefon raqami noto'g'ri."
            )
            return

        db_update_field(
            student_id,
            "phone",
            phone
        )

    # -----------------------------------------------------
    # ADDRESS
    # -----------------------------------------------------

    elif field == "address":

        if len(text) < 3:
            await message.answer(
                "❌ Manzil juda qisqa."
            )
            return

        db_update_field(
            student_id,
            "address",
            text
        )

    else:
        await state.clear()
        return

    await state.clear()

    student = db_get_student(student_id)

    await message.answer(
        "✅ <b>Ma'lumot muvaffaqiyatli yangilandi!</b>\n\n"
        + format_student(student)
        + "\n\n✏️ Yana qaysi ma'lumotni tahrirlash kerak?",
        reply_markup=student_edit_keyboard(student_id)
    )


# =========================================================
# ADMIN MENU
# =========================================================

@dp.message(Command("admin"))
async def cmd_admin(
    message: Message,
    state: FSMContext
):
    if not is_admin(message.from_user.id):
        await message.answer(
            "❌ Siz admin emassiz."
        )
        return

    await state.clear()

    await message.answer(
        "🛠 <b>Admin panel</b>\n\n"
        "Kerakli bo'limni tanlang:",
        reply_markup=admin_menu_keyboard()
    )


# =========================================================
# ADMIN STATS
# =========================================================

@dp.message(Command("stats"))
async def cmd_stats(message: Message):
    if not is_admin(message.from_user.id):
        return

    await send_admin_stats(message)


@dp.callback_query(F.data == "admin_stats")
async def admin_stats_callback(
    callback: CallbackQuery
):
    await callback.answer()

    if not is_admin(callback.from_user.id):
        return

    await send_admin_stats(callback.message)


async def send_admin_stats(message: Message):

    total = db_get_total()
    today = db_get_today_count()

    conn = get_db()

    class_rows = conn.execute("""
        SELECT class_name, COUNT(*) AS count
        FROM students
        GROUP BY class_name
        ORDER BY class_name
    """).fetchall()

    doc_rows = conn.execute("""
        SELECT doc_type, COUNT(*) AS count
        FROM students
        GROUP BY doc_type
    """).fetchall()

    users_count = conn.execute("""
        SELECT COUNT(DISTINCT tg_user_id) AS count
        FROM students
    """).fetchone()["count"]

    conn.close()

    text = (
        "📊 <b>Statistika</b>\n\n"
        f"👥 Jami o'quvchilar: <b>{total}</b>\n"
        f"📱 Unikal Telegram akkauntlar: <b>{users_count}</b>\n"
        f"📅 Bugun ro'yxatdan o'tgan: <b>{today}</b>\n\n"
        "📚 <b>Sinf bo'yicha:</b>\n"
    )

    for row in class_rows:
        text += (
            f"• {row['class_name']}: "
            f"<b>{row['count']}</b>\n"
        )

    text += "\n🪪 <b>Hujjat bo'yicha:</b>\n"

    for row in doc_rows:
        text += (
            f"• {row['doc_type']}: "
            f"<b>{row['count']}</b>\n"
        )

    await message.answer(
        text,
        reply_markup=admin_menu_keyboard()
    )


# =========================================================
# ADMIN SEARCH
# =========================================================

@dp.message(Command("search"))
async def cmd_search(
    message: Message,
    state: FSMContext
):
    if not is_admin(message.from_user.id):
        return

    parts = message.text.split(maxsplit=1)

    if len(parts) == 1:
        await state.set_state(
            AdminSearch.waiting_query
        )

        await message.answer(
            "🔎 Qidiruv so'zini kiriting.\n\n"
            "F.I.Sh., JSHSHIR, telefon, sinf yoki manzil bo'yicha qidirishingiz mumkin."
        )
        return

    query = parts[1].strip()

    await send_admin_search_results(
        message,
        query
    )


@dp.callback_query(F.data == "admin_search")
async def admin_search_callback(
    callback: CallbackQuery,
    state: FSMContext
):
    await callback.answer()

    if not is_admin(callback.from_user.id):
        return

    await state.set_state(
        AdminSearch.waiting_query
    )

    await callback.message.edit_text(
        "🔎 Qidiruv so'zini kiriting.\n\n"
        "F.I.Sh., JSHSHIR, telefon, sinf yoki manzil bo'yicha qidirish mumkin."
    )


@dp.message(AdminSearch.waiting_query)
async def admin_search_message(
    message: Message,
    state: FSMContext
):
    if not is_admin(message.from_user.id):
        return

    if not message.text:
        await message.answer(
            "⚠️ Iltimos, qidiruv so'zini matn ko'rinishida yozing."
        )
        return

    query = message.text.strip()

    await state.clear()

    await send_admin_search_results(
        message,
        query
    )


async def send_admin_search_results(
    message: Message,
    query: str
):
    students = db_search_students(query)

    if not students:
        await message.answer(
            f"🔎 <b>{query}</b> bo'yicha natija topilmadi.",
            reply_markup=admin_menu_keyboard()
        )
        return

    await message.answer(
        f"🔎 <b>{query}</b> bo'yicha "
        f"{len(students)} ta natija topildi:"
    )

    for student in students:

        await message.answer(
            format_student(student),
            reply_markup=admin_search_result_keyboard(
                student["id"]
            )
        )


# =========================================================
# ADMIN — SELECT STUDENT
# =========================================================

@dp.callback_query(
    F.data.startswith("admin_select_student:")
)
async def admin_select_student(
    callback: CallbackQuery,
    state: FSMContext
):
    await callback.answer()

    if not is_admin(callback.from_user.id):
        return

    try:
        student_id = int(
            callback.data.split(":")[1]
        )
    except (ValueError, IndexError):
        return

    student = db_get_student(student_id)

    if not student:
        await callback.message.edit_text(
            "❌ O'quvchi topilmadi."
        )
        return

    await state.clear()

    await callback.message.edit_text(
        "👤 <b>O'quvchi</b>\n\n"
        + format_student(student)
        + "\n\n✏️ Qaysi ma'lumotni tahrirlash kerak?",
        reply_markup=student_edit_keyboard(
            student_id,
            admin=True
        )
    )


# =========================================================
# ADMIN — EDIT FIELD
# =========================================================

@dp.callback_query(
    F.data.startswith("admin_edit_field:")
)
async def admin_edit_field(
    callback: CallbackQuery,
    state: FSMContext
):
    await callback.answer()

    if not is_admin(callback.from_user.id):
        return

    parts = callback.data.split(":")

    if len(parts) != 3:
        return

    try:
        student_id = int(parts[1])
    except ValueError:
        return

    field = parts[2]

    student = db_get_student(student_id)

    if not student:
        await callback.message.answer(
            "❌ O'quvchi topilmadi."
        )
        return

    if field not in FIELD_NAMES:
        return

    await state.set_state(
        AdminEdit.waiting_value
    )

    await state.update_data(
        student_id=student_id,
        field=field
    )

    if field == "class_name":

        await callback.message.edit_text(
            "📚 Yangi sinfni tanlang:",
            reply_markup=classes_keyboard(
                prefix=f"admin_edit_class:{student_id}"
            )
        )

        return

    prompts = {
        "full_name":
            "👤 Yangi F.I.Sh.ni kiriting:",

        "birth_date":
            "🎂 Yangi tug'ilgan sanani kiriting.\n\n"
            "Format: <b>DD.MM.YYYY</b>",

        "document":
            "🪪 Yangi JSHSHIR yoki Metrikani kiriting:",

        "phone":
            "📱 Yangi telefon raqamini kiriting:",

        "address":
            "🏠 Yangi manzilni kiriting:",
    }

    await callback.message.edit_text(
        prompts[field]
    )


# =========================================================
# ADMIN — EDIT CLASS
# =========================================================

@dp.callback_query(
    F.data.startswith("admin_edit_class:")
)
async def admin_edit_class(
    callback: CallbackQuery,
    state: FSMContext
):
    await callback.answer()

    if not is_admin(callback.from_user.id):
        return

    parts = callback.data.split(":")

    if len(parts) != 3:
        return

    try:
        student_id = int(parts[1])
    except ValueError:
        return

    class_name = parts[2]

    student = db_get_student(student_id)

    if not student:
        return

    db_update_field(
        student_id,
        "class_name",
        class_name
    )

    await state.clear()

    student = db_get_student(student_id)

    await callback.message.edit_text(
        "✅ <b>Sinf yangilandi!</b>\n\n"
        + format_student(student)
        + "\n\n✏️ Yana qaysi ma'lumotni tahrirlash kerak?",
        reply_markup=student_edit_keyboard(
            student_id,
            admin=True
        )
    )


# =========================================================
# ADMIN — EDIT VALUE
# =========================================================

@dp.message(AdminEdit.waiting_value)
async def admin_edit_value(
    message: Message,
    state: FSMContext
):
    if not is_admin(message.from_user.id):
        return

    data = await state.get_data()

    student_id = data.get("student_id")
    field = data.get("field")

    if not student_id or not field:
        await state.clear()
        return

    student = db_get_student(student_id)

    if not student:
        await state.clear()

        await message.answer(
            "❌ O'quvchi topilmadi."
        )
        return

    if not message.text:
        await message.answer(
            "⚠️ Iltimos, yangi qiymatni matn ko'rinishida yuboring."
        )
        return

    text = message.text.strip()

    if field == "full_name":

        if len(text) < 5:
            await message.answer(
                "❌ F.I.Sh. juda qisqa."
            )
            return

        db_update_field(
            student_id,
            "full_name",
            text
        )

    elif field == "birth_date":

        if not validate_date(text):
            await message.answer(
                "❌ Sana noto'g'ri.\n"
                "Format: DD.MM.YYYY"
            )
            return

        if not validate_jshshir_birth_date(
            student["doc_type"],
            student["doc_value"],
            text
        ):
            await message.answer(
                "❌ Sana JSHSHIR bilan mos kelmadi."
            )
            return

        db_update_field(
            student_id,
            "birth_date",
            text
        )

    elif field == "document":

        doc_type, doc_value = extract_document(text)

        if not doc_type:
            await message.answer(
                "❌ JSHSHIR/Metrika noto'g'ri."
            )
            return

        if not validate_jshshir_birth_date(
            doc_type,
            doc_value,
            student["birth_date"]
        ):
            await message.answer(
                "❌ JSHSHIRdagi tug'ilgan sana "
                "o'quvchining tug'ilgan sanasi bilan mos emas."
            )
            return

        owner = db_get_by_doc(doc_value, exclude_id=student_id)
        if owner:
            await message.answer(
                "❌ Bu JSHSHIR/Metrika boshqa o'quvchiga tegishli:\n\n"
                f"👤 <b>{owner['full_name']}</b>\n"
                f"📚 Sinf: {owner['class_name']}\n"
                f"🆔 ID: {owner['id']}"
            )
            return

        db_update_field(
            student_id,
            "doc_type",
            doc_type
        )

        db_update_field(
            student_id,
            "doc_value",
            doc_value
        )

    elif field == "phone":

        phone = normalize_phone(text)

        if not validate_phone(phone):
            await message.answer(
                "❌ Telefon raqami noto'g'ri."
            )
            return

        db_update_field(
            student_id,
            "phone",
            phone
        )

    elif field == "address":

        if len(text) < 3:
            await message.answer(
                "❌ Manzil juda qisqa."
            )
            return

        db_update_field(
            student_id,
            "address",
            text
        )

    else:
        await state.clear()
        return

    await state.clear()

    student = db_get_student(student_id)

    await message.answer(
        "✅ <b>Ma'lumot yangilandi!</b>\n\n"
        + format_student(student)
        + "\n\n✏️ Yana qaysi ma'lumotni tahrirlash kerak?",
        reply_markup=student_edit_keyboard(
            student_id,
            admin=True
        )
    )


# =========================================================
# ADMIN — BACK
# =========================================================

@dp.callback_query(F.data == "admin_back_students")
async def admin_back_students(
    callback: CallbackQuery,
    state: FSMContext
):
    await callback.answer()

    if not is_admin(callback.from_user.id):
        return

    await state.clear()

    await callback.message.edit_text(
        "🛠 <b>Admin panel</b>\n\n"
        "Kerakli bo'limni tanlang:",
        reply_markup=admin_menu_keyboard()
    )


# =========================================================
# ADMIN — DELETE
# =========================================================

@dp.callback_query(
    F.data.startswith("admin_delete:")
)
async def admin_delete(
    callback: CallbackQuery
):
    await callback.answer()

    if not is_admin(callback.from_user.id):
        return

    try:
        student_id = int(
            callback.data.split(":")[1]
        )
    except (ValueError, IndexError):
        return

    student = db_get_student(student_id)

    if not student:
        await callback.message.edit_text(
            "❌ O'quvchi topilmadi."
        )
        return

    keyboard = InlineKeyboardMarkup(
        inline_keyboard=[
            [
                InlineKeyboardButton(
                    text="⚠️ Ha, o'chirish",
                    callback_data=f"confirm_delete:{student_id}"
                )
            ],
            [
                InlineKeyboardButton(
                    text="❌ Bekor qilish",
                    callback_data=f"cancel_delete:{student_id}"
                )
            ]
        ]
    )

    await callback.message.edit_text(
        "⚠️ <b>O'quvchini o'chirishni tasdiqlaysizmi?</b>\n\n"
        + format_student(student),
        reply_markup=keyboard
    )


@dp.callback_query(
    F.data.startswith("confirm_delete:")
)
async def confirm_delete(
    callback: CallbackQuery
):
    await callback.answer()

    if not is_admin(callback.from_user.id):
        return

    try:
        student_id = int(
            callback.data.split(":")[1]
        )
    except (ValueError, IndexError):
        return

    student = db_get_student(student_id)

    if not student:
        await callback.message.edit_text(
            "❌ O'quvchi topilmadi."
        )
        return

    name = student["full_name"]

    db_delete_student(student_id)

    await callback.message.edit_text(
        f"🗑 <b>{name}</b> o'chirildi."
    )


@dp.callback_query(
    F.data.startswith("cancel_delete:")
)
async def cancel_delete(
    callback: CallbackQuery
):
    await callback.answer()

    if not is_admin(callback.from_user.id):
        return

    try:
        student_id = int(
            callback.data.split(":")[1]
        )
    except (ValueError, IndexError):
        return

    student = db_get_student(student_id)

    if not student:
        await callback.message.edit_text(
            "❌ O'quvchi topilmadi."
        )
        return

    await callback.message.edit_text(
        format_student(student),
        reply_markup=student_edit_keyboard(
            student_id,
            admin=True
        )
    )


# =========================================================
# BROADCAST
# =========================================================

@dp.callback_query(F.data == "admin_broadcast")
async def admin_broadcast(
    callback: CallbackQuery,
    state: FSMContext
):
    await callback.answer()

    if not is_admin(callback.from_user.id):
        return

    await state.set_state(
        Broadcast.waiting_message
    )

    await callback.message.edit_text(
        "📢 <b>Broadcast</b>\n\n"
        "Barcha foydalanuvchilarga yuboriladigan xabarni kiriting.\n\n"
        "⚠️ Bir Telegram akkauntga, unda nechta o'quvchi bo'lishidan qat'i nazar, "
        "xabar faqat bir marta yuboriladi."
    )


@dp.message(Broadcast.waiting_message)
async def broadcast_message(
    message: Message,
    state: FSMContext
):
    if not is_admin(message.from_user.id):
        return

    if not message.text:
        await message.answer(
            "⚠️ Iltimos, broadcast uchun matnli xabar yuboring."
        )
        return

    broadcast_text = message.text

    await state.clear()

    user_ids = db_get_unique_users()

    success = 0
    failed = 0

    await message.answer(
        f"📢 Broadcast boshlandi.\n"
        f"👥 {len(user_ids)} ta Telegram akkauntga yuboriladi."
    )

    for user_id in user_ids:

        try:
            await bot.send_message(
                user_id,
                broadcast_text
            )

            success += 1

        except Exception as e:
            failed += 1

            logging.warning(
                f"Broadcast error {user_id}: {e}"
            )

        await asyncio.sleep(0.05)

    await message.answer(
        "✅ <b>Broadcast tugadi</b>\n\n"
        f"✅ Yuborildi: <b>{success}</b>\n"
        f"❌ Xatolik: <b>{failed}</b>",
        reply_markup=admin_menu_keyboard()
    )


# =========================================================
# EXCEL
# =========================================================

# Ranglar
CLR_TITLE = "1F4E78"      # to'q ko'k (sarlavha)
CLR_HEADER = "2E75B6"     # ko'k (ustun sarlavhalari)
CLR_ROW_ALT = "DDEBF7"    # och ko'k (juft qatorlar)
CLR_BORDER = "9DC3E6"

BASE_HEADERS = [
    "№",
    "F.I.Sh.",
    "Tug'ilgan sana",
    "Hujjat turi",
    "JSHSHIR / Metrika",
    "Telefon",
    "Manzil",
    "Ro'yxatdan o'tgan",
]


def _style_sheet(ws, title_text: str, students, with_class_col: bool = False):
    """Bitta varaqni chiroyli qilib bezaydi."""
    thin = Side(style="thin", color=CLR_BORDER)
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    headers = BASE_HEADERS.copy()
    if with_class_col:
        headers.insert(1, "Sinf")

    n_cols = len(headers)
    last_col = get_column_letter(n_cols)

    # --- Sarlavha qatori (1-qator) ---
    ws.merge_cells(f"A1:{last_col}1")
    c = ws["A1"]
    c.value = title_text
    c.font = Font(name="Calibri", size=15, bold=True, color="FFFFFF")
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.fill = PatternFill("solid", fgColor=CLR_TITLE)
    ws.row_dimensions[1].height = 30

    # --- Ustun sarlavhalari (2-qator) ---
    for col, name in enumerate(headers, 1):
        cell = ws.cell(row=2, column=col, value=name)
        cell.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor=CLR_HEADER)
        cell.alignment = Alignment(
            horizontal="center", vertical="center", wrap_text=True
        )
        cell.border = border
    ws.row_dimensions[2].height = 26

    # --- Ma'lumot qatorlari ---
    for i, s in enumerate(students, start=1):
        r = i + 2  # 3-qatordan boshlanadi
        if with_class_col:
            values = [
                i, s["class_name"], s["full_name"], s["birth_date"],
                s["doc_type"], s["doc_value"], s["phone"],
                s["address"] or "", s["created_at"],
            ]
        else:
            values = [
                i, s["full_name"], s["birth_date"], s["doc_type"],
                s["doc_value"], s["phone"], s["address"] or "",
                s["created_at"],
            ]

        for col, val in enumerate(values, 1):
            cell = ws.cell(row=r, column=col, value=val)
            cell.border = border
            cell.alignment = Alignment(
                horizontal="center" if col == 1 else "left", vertical="center"
            )
            cell.font = Font(name="Calibri", size=11)
            if i % 2 == 0:
                cell.fill = PatternFill("solid", fgColor=CLR_ROW_ALT)

    # --- Ustun kengliklari ---
    if with_class_col:
        widths = [5, 8, 32, 15, 13, 22, 20, 28, 20]
    else:
        widths = [5, 32, 15, 13, 22, 20, 28, 20]
    for col, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(col)].width = w

    ws.freeze_panes = "A3"


def _wb_bytes(wb: Workbook) -> bytes:
    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


def build_excel_class(class_name: str, students) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = class_name.replace("/", "-")[:31]
    title = f"{class_name} sinfi o'quvchilari  •  Jami: {len(students)} ta"
    _style_sheet(ws, title, students, with_class_col=False)
    return _wb_bytes(wb)


def build_excel_all(students) -> bytes:
    """'Hammasi' varag'i + har bir sinf uchun alohida varaq."""
    wb = Workbook()

    # Umumiy varaq
    ws_all = wb.active
    ws_all.title = "Hammasi"
    _style_sheet(
        ws_all,
        f"Barcha o'quvchilar  •  Jami: {len(students)} ta",
        students,
        with_class_col=True,
    )

    # Sinflar bo'yicha guruhlash
    by_class = {}
    for s in students:
        by_class.setdefault(s["class_name"], []).append(s)

    for cls in CLASSES:
        cls_students = by_class.get(cls)
        if not cls_students:
            continue
        ws = wb.create_sheet(title=cls.replace("/", "-")[:31])
        _style_sheet(
            ws,
            f"{cls} sinfi o'quvchilari  •  Jami: {len(cls_students)} ta",
            cls_students,
            with_class_col=False,
        )

    return _wb_bytes(wb)


def excel_menu_keyboard():
    """Sinf bo'yicha yoki barchasini yuklab olish menyusi."""
    conn = get_db()
    counts = dict(
        conn.execute(
            "SELECT class_name, COUNT(*) FROM students GROUP BY class_name"
        ).fetchall()
    )
    conn.close()

    rows = []
    row = []
    for i, cls in enumerate(CLASSES, 1):
        n = counts.get(cls, 0)
        if n == 0:
            continue
        row.append(
            InlineKeyboardButton(
                text=f"📚 {cls} ({n})",
                callback_data=f"xls:{cls}"
            )
        )
        if len(row) == 2:
            rows.append(row)
            row = []
    if row:
        rows.append(row)

    rows.append([
        InlineKeyboardButton(
            text="📦 Barcha sinflar (bitta fayl)",
            callback_data="xls:__ALL__"
        )
    ])
    rows.append([
        InlineKeyboardButton(
            text="⬅️ Admin panel",
            callback_data="admin_back_menu"
        )
    ])
    return InlineKeyboardMarkup(inline_keyboard=rows)


@dp.callback_query(F.data == "admin_excel")
async def admin_excel(callback: CallbackQuery):
    await callback.answer()

    if not is_admin(callback.from_user.id):
        return

    total = db_get_total()
    if total == 0:
        await callback.message.edit_text(
            "📭 Hozircha ma'lumotlar yo'q.",
            reply_markup=admin_menu_keyboard()
        )
        return

    await callback.message.edit_text(
        "📥 <b>Excel yuklab olish</b>\n\n"
        "Kerakli sinfni tanlang yoki barchasini bitta faylda oling:",
        reply_markup=excel_menu_keyboard()
    )


@dp.callback_query(F.data == "admin_back_menu")
async def admin_back_menu(callback: CallbackQuery, state: FSMContext):
    await callback.answer()
    if not is_admin(callback.from_user.id):
        return
    await state.clear()
    await callback.message.edit_text(
        "🛠 <b>Admin panel</b>\n\n"
        "Kerakli bo'limni tanlang:",
        reply_markup=admin_menu_keyboard()
    )


@dp.callback_query(F.data.startswith("xls:"))
async def admin_excel_export(callback: CallbackQuery):
    if not is_admin(callback.from_user.id):
        await callback.answer("⛔️ Ruxsat yo'q", show_alert=True)
        return

    await callback.answer("⏳ Excel tayyorlanmoqda...")

    target = callback.data.split(":", 1)[1]

    if target == "__ALL__":
        students = db_get_all_students()
        if not students:
            await callback.message.answer("📭 Ma'lumotlar yo'q.")
            return
        file_bytes = build_excel_all(students)
        fname = f"barcha_sinflar_{datetime.now():%Y%m%d_%H%M}.xlsx"
        caption = f"📦 <b>Barcha sinflar</b>\n👥 Jami: {len(students)} ta o'quvchi"
    else:
        students = db_get_students_by_class(target)
        if not students:
            await callback.message.answer(
                f"📭 <b>{target}</b> sinfida ma'lumot yo'q."
            )
            return
        file_bytes = build_excel_class(target, students)
        fname = f"{target.replace('-', '')}_{datetime.now():%Y%m%d_%H%M}.xlsx"
        caption = f"📚 <b>{target}</b> sinfi\n👥 Jami: {len(students)} ta o'quvchi"

    await callback.message.answer_document(
        BufferedInputFile(file_bytes, filename=fname),
        caption=caption
    )


# =========================================================
# BACKUP (zaxira nusxa)
# =========================================================

async def send_backup(target_message: Message):
    if not os.path.exists(DB_PATH):
        await target_message.answer("❌ Baza fayli topilmadi.")
        return

    fname = f"backup_{datetime.now():%Y%m%d_%H%M%S}.db"
    total = db_get_total()

    await target_message.answer_document(
        FSInputFile(DB_PATH, filename=fname),
        caption=(
            "🗄 <b>Ma'lumotlar bazasi zaxira nusxasi</b>\n\n"
            f"👥 Jami o'quvchilar: <b>{total}</b>\n"
            f"🕐 {datetime.now():%Y-%m-%d %H:%M:%S}\n\n"
            "⚠️ Bu faylni xavfsiz joyda saqlang."
        )
    )


@dp.message(Command("backup"))
async def cmd_backup(message: Message):
    if not is_admin(message.from_user.id):
        return
    await send_backup(message)


@dp.callback_query(F.data == "admin_backup")
async def admin_backup(callback: CallbackQuery):
    await callback.answer("⏳ Zaxira tayyorlanmoqda...")
    if not is_admin(callback.from_user.id):
        return
    await send_backup(callback.message)


# =========================================================
# RESTORE (bazani yuklab o'tkazish) — admin .db fayl yuboradi
# =========================================================

@dp.message(F.document)
async def restore_db(message: Message):
    """Admin .db faylini yuborsa — mavjud bazani almashtiradi.
    Lokal ma'lumotni Render'ga ko'chirish uchun ishlatiladi."""
    if not is_admin(message.from_user.id):
        return

    doc = message.document
    name = (doc.file_name or "").lower()
    if not name.endswith(".db"):
        await message.answer(
            "ℹ️ Bazani tiklash uchun <b>.db</b> faylini yuboring "
            "(masalan /backup dan olingan fayl)."
        )
        return

    try:
        # Eski bazadan zaxira olib qo'yamiz (ehtiyot uchun)
        if os.path.exists(DB_PATH):
            safety = DB_PATH + ".old"
            try:
                if os.path.exists(safety):
                    os.remove(safety)
                os.replace(DB_PATH, safety)
            except OSError:
                pass

        await bot.download(doc, destination=DB_PATH)

        # Sxema/migratsiyani ta'minlaymiz
        init_db()
        total = db_get_total()

        await message.answer(
            "✅ <b>Baza muvaffaqiyatli tiklandi!</b>\n\n"
            f"👥 Jami o'quvchilar: <b>{total}</b>\n\n"
            "Endi bot shu ma'lumotlar bilan ishlaydi."
        )
        logger.info("Baza tiklandi. Jami: %s", total)
    except Exception as e:
        logger.exception("Restore xatosi: %s", e)
        await message.answer(
            "❌ Bazani tiklashda xatolik yuz berdi. "
            "Fayl to'g'ri .db ekanini tekshiring."
        )


# =========================================================
# UNKNOWN TEXT
# =========================================================

@dp.message()
async def fallback(message: Message, state: FSMContext):
    await state.clear()

    students = db_get_user_students(
        message.from_user.id
    )

    if students:
        await message.answer(
            "Quyidagi o'quvchilardan birini tanlang "
            "yoki yangi o'quvchi qo'shing:",
            reply_markup=user_students_keyboard(
                students
            )
        )
    else:
        await state.set_state(Registration.choosing_class)
        await message.answer(
            "Ro'yxatdan o'tishni boshlash uchun "
            "sinfni tanlang:",
            reply_markup=classes_keyboard()
        )


# =========================================================
# GLOBAL XATO USHLAGICH
# =========================================================

@dp.errors()
async def global_error_handler(event: ErrorEvent):
    logger.exception("Update ishlovida xato: %s", event.exception)

    try:
        update = event.update
        if update.message:
            await update.message.answer(
                "⚠️ Xatolik yuz berdi. Iltimos, qaytadan urinib ko'ring "
                "yoki /start bosing."
            )
        elif update.callback_query:
            await update.callback_query.answer(
                "⚠️ Xatolik yuz berdi. /start bosing.",
                show_alert=True
            )
    except Exception:
        pass

    return True


# =========================================================
# MAIN
# =========================================================

async def start_health_server():
    """Render (yoki boshqa hosting) Web Service uchun oddiy HTTP server.
    PORT muhit o'zgaruvchisi bo'lsagina ishga tushadi."""
    port = os.getenv("PORT")
    if not port:
        return

    from aiohttp import web

    async def health(_request):
        return web.Response(text="Bot ishlayapti ✅")

    app = web.Application()
    app.router.add_get("/", health)

    runner = web.AppRunner(app)
    await runner.setup()
    site = web.TCPSite(runner, "0.0.0.0", int(port))
    await site.start()
    logger.info("Health server %s portda ishga tushdi", port)


async def main():
    init_db()

    logger.info("Database initialized (DB_PATH=%s)", DB_PATH)
    logger.info("Bot ishga tushdi ✅  (Adminlar: %s)", ADMIN_IDS or "yo'q!")

    await start_health_server()
    await bot.delete_webhook(drop_pending_updates=True)
    await dp.start_polling(bot)


if __name__ == "__main__":
    try:
        asyncio.run(main())
    except (KeyboardInterrupt, SystemExit):
        logger.info("Bot to'xtatildi.")