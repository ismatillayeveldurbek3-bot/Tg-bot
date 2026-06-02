import os
import re
import csv
import json
import asyncio
import logging
import sqlite3
from html import escape
from pathlib import Path
from datetime import datetime, timezone, timedelta
from typing import Optional, Iterable

from aiogram import Bot, Dispatcher, F
from aiogram.enums import ParseMode, ChatMemberStatus
from aiogram.filters import Command, StateFilter
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.types import (
    Message, CallbackQuery, InlineKeyboardMarkup, InlineKeyboardButton,
    ReplyKeyboardMarkup, KeyboardButton, ReplyKeyboardRemove, FSInputFile
)
from aiogram.utils.keyboard import InlineKeyboardBuilder
from aiogram.exceptions import TelegramBadRequest, TelegramForbiddenError
from aiogram.client.default import DefaultBotProperties

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
except Exception:  # pragma: no cover
    Workbook = None

try:
    from apscheduler.schedulers.asyncio import AsyncIOScheduler
except Exception:  # pragma: no cover
    AsyncIOScheduler = None

# =====================================================
# CONFIG
# =====================================================
BOT_TOKEN = os.getenv("BOT_TOKEN")
if not BOT_TOKEN:
    raise RuntimeError("BOT_TOKEN env o'zgaruvchisi topilmadi.")

CHANNEL_USERNAME = os.getenv("CHANNEL_USERNAME", "@Qashqadaryo_PMM")
CHANNEL_URL = os.getenv("CHANNEL_URL", "https://t.me/Qashqadaryo_PMM")
FACEBOOK_URL = os.getenv("FACEBOOK_URL", "https://www.facebook.com/share/1E4ZVePTh4/")
INSTAGRAM_URL = os.getenv("INSTAGRAM_URL", "https://www.instagram.com/pedagogikmahorat")
ADMIN_IDS = [int(x) for x in os.getenv("ADMIN_IDS", "5298063089,7361393654").replace(";", ",").split(",") if x.strip().isdigit()]

DATA_DIR = Path(os.getenv("DATA_DIR", "/app/data"))
DATA_DIR.mkdir(parents=True, exist_ok=True)
DB_PATH = DATA_DIR / "votes.db"
BACKUP_DIR = DATA_DIR / "backups"
BACKUP_DIR.mkdir(parents=True, exist_ok=True)
UZ_TZ = timezone(timedelta(hours=5), name="Asia/Tashkent")

logging.basicConfig(level=os.getenv("LOG_LEVEL", "INFO"), format="%(asctime)s | %(levelname)s | %(name)s | %(message)s")
log = logging.getLogger("rating-bot")

# =====================================================
# INITIAL DATA / BACKWARD COMPATIBILITY
# =====================================================
SUBJECTS = {
    "s1": {"name": "Tillarni o'qitish metodikasi", "old_key": "tillarni_oqitish_metodikasi", "teachers": {
        "tom_1": "Norov Otajon Shomurodovich", "tom_2": "Abdixolikov Abdulazizxon Abduvohob o'g'li", "tom_3": "Azimova Nigora Anvar qizi", "tom_4": "Abatov Doston Ro'zimurod o'g'li", "tom_5": "Jalilova Komila Abdullayevna", "tom_6": "Oqboyeva Zulfiya Bobonazarovna", "tom_7": "Sevastyanova Nadejda Aleksandrovna", "tom_8": "Xidirova Feruza To'rayevna", "tom_9": "Ergasheva Dilorom Muradilloyevna"}},
    "s2": {"name": "Pedagogika, psixologiya va ta'lim menejmenti", "old_key": "pedagogika_psixologiya_va_talim_menejmenti", "teachers": {
        "pptm_1": "Umarov Lutfillo Murodilloyevich", "pptm_2": "Baratova Nasiba Turobovna", "pptm_3": "Bekmurodova Dilnoza Pirimovna", "pptm_4": "Meyliyev Lobar Nurmatovna", "pptm_5": "Ochilov Og'abek Narzullayevich", "pptm_6": "Shoniyozova Dilafruz Sabirovna", "pptm_7": "Yaratov Xamidjon Muxtorovich", "pptm_8": "Nazarov Asliddin Faxriddin o'g'li", "pptm_9": "Ergasheva Dilafruz Ergamqulovna", "pptm_10": "Soatov Asadulloh Jabborovich"}},
    "s3": {"name": "Aniq va tabiiy fanlar", "old_key": "aniq_va_tabiiy_fanlar", "teachers": {
        "atf_1": "Jobborov Farhod Bo'riyevich", "atf_2": "Karimova Habiba Abduraxmonovna", "atf_3": "Quldoshova Maftuna Jumanzar qizi", "atf_4": "Mallaev Xamro Ro'ziboyevich", "atf_5": "Mamatov Bekzod Farxotovich", "atf_6": "Pardaeva Muqaddas Zafar qizi", "atf_7": "Parmanov Jahongir Rayhonovich", "atf_8": "Rahmatullayev Erkin Shokirovich", "atf_9": "Suyarov Zoir Shojmardonovich", "atf_10": "Tursunova Maftuna Sulton qizi", "atf_11": "Umarov Ibrohimxon Norxuja o'g'li", "atf_12": "Chariev Rashid Ravshanovich", "atf_13": "Elmurodov Sherdil Ergashyevich", "atf_14": "Eshmonov Laziz Norxo'rja o'g'li", "atf_15": "Karaeva Dilfuzaxon Mamasharipovna", "atf_16": "Salomova Madina Sodiq qizi"}},
    "s4": {"name": "Amaliy va ijtimoiy fanlar", "old_key": "amaliy_va_ijtimoiy_fanlar", "teachers": {
        "aif_1": "Yo'ldashev Bekmirza Elmurodovich", "aif_2": "Jabboborov Laziz Hamza o'g'li", "aif_3": "Nurmatov Samandar Fayratovich", "aif_4": "Batoshov Inatillo Kungirovich", "aif_5": "Rajabov Ruslan Bozorovich", "aif_6": "Sanaev Azamat Alponovich", "aif_7": "Shamsiev Jahongir Qulmurod o'g'li", "aif_8": "Xudoyberdiev Axrorboy Nabi o'g'li", "aif_9": "Xasanova Gulnora Qorshanbiyevna", "aif_10": "Eshnazarova Maziya Allanazarovna"}},
    "s5": {"name": "Maktabgacha, boshlang'ich va maxsus ta'lim", "old_key": "maktabgacha_boshlangich_va_maxsus_talim", "teachers": {
        "mbmt_1": "Irisova Sayyora Rajabovna", "mbmt_2": "Azizova Dilnoz Yo'ldoshevna", "mbmt_3": "G'oyimov Umar Eshmurodovich", "mbmt_4": "Ziyotova Madina Mansur qizi", "mbmt_5": "Karimova Umida Sharopovna", "mbmt_6": "Qarshiyeva Guzal Alimardonovna", "mbmt_7": "Qurbanova Xusnora Xudoyberdi qizi", "mbmt_8": "Rajabova Xurshida Hakimovna", "mbmt_9": "Razzaqova Dilnoza Akramovna", "mbmt_10": "Sadinova Marjona Akmal qizi", "mbmt_11": "Shaxmurodova Dilxaxon Almardanovna", "mbmt_12": "Ergasheva Xusniya Mirzoxid qizi", "mbmt_13": "Zaripova Muslima Qurbonovna"}},
}
OLD_TO_NEW_SUBJECT = {v["old_key"]: k for k, v in SUBJECTS.items()}

conn = sqlite3.connect(DB_PATH, check_same_thread=False)
conn.row_factory = sqlite3.Row
db_lock = asyncio.Lock()
bot = Bot(BOT_TOKEN, default=DefaultBotProperties(parse_mode=ParseMode.HTML))
dp = Dispatcher(storage=MemoryStorage())

# =====================================================
# FSM
# =====================================================
class Register(StatesGroup):
    first_name = State(); last_name = State(); phone = State(); after = State()
class ComplaintFSM(StatesGroup):
    department = State(); teacher = State(); text = State()
class TeacherFSM(StatesGroup):
    add_department = State(); add_name = State(); edit_name = State(); student_count = State(); search = State()
class DepartmentFSM(StatesGroup):
    add_name = State(); edit_name = State()
class AdminFSM(StatesGroup):
    broadcast = State(); admin_add = State(); admin_remove = State()

# =====================================================
# HELPERS
# =====================================================
def now_str() -> str: return datetime.now(UZ_TZ).strftime("%Y-%m-%d %H:%M:%S")
def is_admin(uid: int) -> bool: return uid in ADMIN_IDS or str(uid) in get_setting("extra_admins", "").split(",")
def safe(s: Optional[str]) -> str: return escape(s or "")
def slug(text: str, prefix: str = "id") -> str:
    base = re.sub(r"[^a-zA-Z0-9_]+", "_", (text or "").lower()).strip("_")[:24]
    return f"{prefix}_{base or int(datetime.now().timestamp())}"
def setting_bool(key: str, default="1") -> bool: return get_setting(key, default) == "1"

async def answer_cb(call: CallbackQuery, text: str = ""):
    try: await call.answer(text)
    except TelegramBadRequest: pass

async def edit_or_send(target, text: str, reply_markup=None):
    if isinstance(target, CallbackQuery):
        try: await target.message.edit_text(text, reply_markup=reply_markup)
        except TelegramBadRequest: await target.message.answer(text, reply_markup=reply_markup)
    else:
        await target.answer(text, reply_markup=reply_markup)

# =====================================================
# DATABASE + MIGRATIONS
# =====================================================
def col_exists(table: str, col: str) -> bool:
    cur = conn.execute(f"PRAGMA table_info({table})")
    return any(r[1] == col for r in cur.fetchall())

def table_exists(table: str) -> bool:
    return conn.execute("SELECT name FROM sqlite_master WHERE type='table' AND name=?", (table,)).fetchone() is not None

def get_setting(key: str, default: str = "") -> str:
    row = conn.execute("SELECT value FROM settings WHERE key=?", (key,)).fetchone() if table_exists("settings") else None
    return row[0] if row else default

def set_setting(key: str, value: str):
    conn.execute("INSERT INTO settings(key,value) VALUES(?,?) ON CONFLICT(key) DO UPDATE SET value=excluded.value", (key, value)); conn.commit()

def init_db():
    c = conn.cursor()
    c.execute("CREATE TABLE IF NOT EXISTS settings(key TEXT PRIMARY KEY, value TEXT)")
    c.execute("CREATE TABLE IF NOT EXISTS user_prefs(user_id INTEGER PRIMARY KEY, script TEXT DEFAULT 'latin', access_granted INTEGER DEFAULT 0)")
    c.execute("""CREATE TABLE IF NOT EXISTS users(
        telegram_id INTEGER PRIMARY KEY, first_name TEXT, last_name TEXT, fullname TEXT,
        username TEXT, phone TEXT, registered_at TEXT, updated_at TEXT, is_blocked INTEGER DEFAULT 0)""")
    c.execute("""CREATE TABLE IF NOT EXISTS departments(
        department_key TEXT PRIMARY KEY, name TEXT NOT NULL, sort_order INTEGER DEFAULT 0, is_active INTEGER DEFAULT 1, created_at TEXT)""")
    c.execute("""CREATE TABLE IF NOT EXISTS teachers(
        teacher_key TEXT NOT NULL, department_key TEXT NOT NULL, name TEXT NOT NULL,
        student_count INTEGER DEFAULT 0, is_active INTEGER DEFAULT 1, created_at TEXT,
        PRIMARY KEY(department_key, teacher_key))""")
    c.execute("""CREATE TABLE IF NOT EXISTS ratings(
        id INTEGER PRIMARY KEY AUTOINCREMENT, user_id INTEGER NOT NULL, department_key TEXT NOT NULL,
        teacher_key TEXT NOT NULL, rating INTEGER NOT NULL CHECK(rating BETWEEN 1 AND 5),
        fullname TEXT, phone TEXT, username TEXT, rated_at TEXT,
        UNIQUE(user_id, department_key, teacher_key))""")
    c.execute("""CREATE TABLE IF NOT EXISTS complaints(
        id INTEGER PRIMARY KEY AUTOINCREMENT, user_id INTEGER NOT NULL, teacher_id TEXT,
        department_key TEXT, teacher_key TEXT, fullname TEXT, phone TEXT, username TEXT,
        complaint_text TEXT, message_text TEXT, created_at TEXT, status TEXT DEFAULT 'Yangi')""")
    # legacy tables preserved
    c.execute("""CREATE TABLE IF NOT EXISTS votes(
        user_id INTEGER PRIMARY KEY, full_name TEXT, username TEXT, subject_key TEXT NOT NULL, teacher_key TEXT NOT NULL, voted_at TEXT)""")
    conn.commit()

    for col, ddl in [("student_count", "ALTER TABLE teachers ADD COLUMN student_count INTEGER DEFAULT 0"), ("is_active", "ALTER TABLE teachers ADD COLUMN is_active INTEGER DEFAULT 1")]:
        if not col_exists("teachers", col): c.execute(ddl)
    for col, ddl in [("phone", "ALTER TABLE users ADD COLUMN phone TEXT"), ("is_blocked", "ALTER TABLE users ADD COLUMN is_blocked INTEGER DEFAULT 0")]:
        if not col_exists("users", col): c.execute(ddl)
    for col, ddl in [("department_key", "ALTER TABLE complaints ADD COLUMN department_key TEXT"), ("teacher_key", "ALTER TABLE complaints ADD COLUMN teacher_key TEXT"), ("teacher_id", "ALTER TABLE complaints ADD COLUMN teacher_id TEXT"), ("phone", "ALTER TABLE complaints ADD COLUMN phone TEXT"), ("complaint_text", "ALTER TABLE complaints ADD COLUMN complaint_text TEXT"), ("status", "ALTER TABLE complaints ADD COLUMN status TEXT DEFAULT 'Yangi'")]:
        if not col_exists("complaints", col): c.execute(ddl)
    conn.commit()

    # seed departments/teachers without overwriting admin changes
    if conn.execute("SELECT COUNT(*) FROM departments").fetchone()[0] == 0:
        for i, (dkey, data) in enumerate(SUBJECTS.items()):
            c.execute("INSERT OR IGNORE INTO departments(department_key,name,sort_order,created_at) VALUES(?,?,?,?)", (dkey, data["name"], i, now_str()))
            for tkey, tname in data["teachers"].items():
                c.execute("INSERT OR IGNORE INTO teachers(teacher_key,department_key,name,created_at) VALUES(?,?,?,?)", (tkey, dkey, tname, now_str()))
    # migrate old db_subjects/db_teachers if they exist
    if table_exists("db_subjects"):
        for r in conn.execute("SELECT subject_key, subject_name, sort_order FROM db_subjects"):
            dkey = OLD_TO_NEW_SUBJECT.get(r[0], r[0])
            c.execute("INSERT OR IGNORE INTO departments(department_key,name,sort_order,created_at) VALUES(?,?,?,?)", (dkey, r[1], r[2] or 0, now_str()))
    if table_exists("db_teachers"):
        for r in conn.execute("SELECT teacher_key, subject_key, teacher_name FROM db_teachers"):
            dkey = OLD_TO_NEW_SUBJECT.get(r[1], r[1])
            c.execute("INSERT OR IGNORE INTO teachers(teacher_key,department_key,name,created_at) VALUES(?,?,?,?)", (r[0], dkey, r[2], now_str()))
    # migrate old user names from votes/teacher_ratings
    for table in ["votes", "teacher_ratings"]:
        if table_exists(table):
            cols = [x[1] for x in conn.execute(f"PRAGMA table_info({table})")]
            if "user_id" in cols:
                name_col = "full_name" if "full_name" in cols else "fullname" if "fullname" in cols else None
                uname_col = "username" if "username" in cols else None
                select = f"SELECT user_id{','+name_col if name_col else ''}{','+uname_col if uname_col else ''} FROM {table}"
                for row in conn.execute(select):
                    uid = row[0]; full = row[1] if name_col else ""; uname = row[2] if name_col and uname_col else (row[1] if uname_col and not name_col else "")
                    c.execute("INSERT OR IGNORE INTO users(telegram_id, fullname, username, registered_at, updated_at) VALUES(?,?,?,?,?)", (uid, full, uname, now_str(), now_str()))
    # migrate like/dislike table into 1-5 ratings
    if table_exists("teacher_ratings"):
        cols = [x[1] for x in conn.execute("PRAGMA table_info(teacher_ratings)")]
        if "rating" in cols:
            for r in conn.execute("SELECT * FROM teacher_ratings"):
                dkey = OLD_TO_NEW_SUBJECT.get(r["subject_key"], r["subject_key"]) if "subject_key" in r.keys() else r["department_key"]
                val = r["rating"]
                try: iv = int(val)
                except Exception: iv = 5 if val == "like" else 1
                c.execute("""INSERT OR IGNORE INTO ratings(user_id,department_key,teacher_key,rating,fullname,username,rated_at)
                             VALUES(?,?,?,?,?,?,?)""", (r["user_id"], dkey, r["teacher_key"], iv, r["full_name"] if "full_name" in r.keys() else "", r["username"] if "username" in r.keys() else "", r["rated_at"] if "rated_at" in r.keys() else now_str()))
    # normalize old subject keys
    for old, new in OLD_TO_NEW_SUBJECT.items():
        c.execute("UPDATE votes SET subject_key=? WHERE subject_key=?", (new, old))
        c.execute("UPDATE ratings SET department_key=? WHERE department_key=?", (new, old))
    # indexes
    indexes = [
        "CREATE INDEX IF NOT EXISTS idx_ratings_teacher ON ratings(department_key,teacher_key)",
        "CREATE INDEX IF NOT EXISTS idx_ratings_user ON ratings(user_id)",
        "CREATE INDEX IF NOT EXISTS idx_complaints_teacher ON complaints(department_key,teacher_key)",
        "CREATE INDEX IF NOT EXISTS idx_complaints_status ON complaints(status)",
        "CREATE INDEX IF NOT EXISTS idx_teachers_department ON teachers(department_key,is_active)",
    ]
    for q in indexes: c.execute(q)
    if get_setting("voting_open", "") == "": set_setting("voting_open", "1")
    if get_setting("daily_backup", "") == "": set_setting("daily_backup", "1")
    conn.commit()

# =====================================================
# DATA ACCESS
# =====================================================
def ensure_user_obj(m: Message):
    u = m.from_user
    full = " ".join([x for x in [u.first_name, u.last_name] if x])
    conn.execute("""INSERT INTO users(telegram_id,first_name,last_name,fullname,username,registered_at,updated_at)
        VALUES(?,?,?,?,?,?,?) ON CONFLICT(telegram_id) DO UPDATE SET username=excluded.username, updated_at=excluded.updated_at""",
        (u.id, u.first_name, u.last_name, full, u.username, now_str(), now_str()))
    conn.execute("INSERT OR IGNORE INTO user_prefs(user_id) VALUES(?)", (u.id,)); conn.commit()

def user_registered(uid: int) -> bool:
    r = conn.execute("SELECT first_name,last_name,phone FROM users WHERE telegram_id=?", (uid,)).fetchone()
    return bool(r and r[0] and r[1] and r[2])

def get_user(uid: int): return conn.execute("SELECT * FROM users WHERE telegram_id=?", (uid,)).fetchone()
def departments(active=True):
    q = "SELECT * FROM departments" + (" WHERE is_active=1" if active else "") + " ORDER BY sort_order,name"
    return conn.execute(q).fetchall()
def teachers(dkey: str, active=True):
    q = "SELECT * FROM teachers WHERE department_key=?" + (" AND is_active=1" if active else "") + " ORDER BY name"
    return conn.execute(q, (dkey,)).fetchall()
def all_teachers(active=True):
    q = """SELECT t.*, d.name AS department_name FROM teachers t JOIN departments d ON d.department_key=t.department_key"""
    if active: q += " WHERE t.is_active=1 AND d.is_active=1"
    q += " ORDER BY d.sort_order,t.name"
    return conn.execute(q).fetchall()
def dep_name(dkey):
    r = conn.execute("SELECT name FROM departments WHERE department_key=?", (dkey,)).fetchone(); return r[0] if r else dkey
def teacher_name(dkey,tkey):
    r = conn.execute("SELECT name FROM teachers WHERE department_key=? AND teacher_key=?", (dkey,tkey)).fetchone(); return r[0] if r else tkey

def save_rating(uid:int, dkey:str, tkey:str, rating:int):
    u=get_user(uid); full=(u["fullname"] if u else "") or " ".join(filter(None,[u["first_name"] if u else "", u["last_name"] if u else ""]))
    conn.execute("""INSERT INTO ratings(user_id,department_key,teacher_key,rating,fullname,phone,username,rated_at)
        VALUES(?,?,?,?,?,?,?,?) ON CONFLICT(user_id,department_key,teacher_key) DO UPDATE SET rating=excluded.rating, fullname=excluded.fullname, phone=excluded.phone, username=excluded.username, rated_at=excluded.rated_at""",
        (uid,dkey,tkey,rating,full,u["phone"] if u else "",u["username"] if u else "",now_str()))
    conn.commit()

def teacher_stats(dkey:str,tkey:str):
    t=conn.execute("SELECT * FROM teachers WHERE department_key=? AND teacher_key=?",(dkey,tkey)).fetchone()
    rows=conn.execute("SELECT rating,COUNT(*) c FROM ratings WHERE department_key=? AND teacher_key=? GROUP BY rating",(dkey,tkey)).fetchall()
    dist={i:0 for i in range(1,6)}
    total=0; weighted=0
    for r in rows:
        dist[int(r[0])]=r[1]; total+=r[1]; weighted+=int(r[0])*r[1]
    avg=weighted/total if total else 0
    sc=(t["student_count"] or 0) if t else 0
    participation=(total/sc*100) if sc else 0
    final=avg*(total/sc) if sc else 0
    return {"avg":avg,"total":total,"student_count":sc,"participation":participation,"final":final,"dist":dist}

def department_stats(dkey:str):
    ts=teachers(dkey)
    total_teachers=len(ts)
    total_votes=0; weighted_sum=0; student_sum=0; finals=[]
    for t in ts:
        st=teacher_stats(dkey,t["teacher_key"])
        total_votes += st["total"]
        weighted_sum += st["avg"]*st["total"]
        student_sum += st["student_count"]
        finals.append(st["final"])
    avg=weighted_sum/total_votes if total_votes else 0
    participation=total_votes/student_sum*100 if student_sum else 0
    final=sum(finals)/len(finals) if finals else 0
    return {"teachers":total_teachers,"votes":total_votes,"avg":avg,"participation":participation,"final":final,"student_sum":student_sum}

# =====================================================
# KEYBOARDS
# =====================================================
def ik(rows: list[list[tuple[str,str]]]) -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text=t, callback_data=d) for t,d in row] for row in rows])
def main_kb(): return ik([[('📝 Baholash','rate:start')],[('📢 Shikoyat','complaint:start'),('🏆 Reyting','rating:top')],[('ℹ️ Ma’lumot','info')]])
def admin_kb(): return ik([[('📊 Dashboard','admin:dash')],[('👨‍🏫 O‘qituvchilar','admin:teachers'),('⭐ Statistika','admin:tstats')],[('🏛 Kafedralar','admin:dept_rating'),('📢 Shikoyatlar','admin:complaints')],[('📁 Export / Backup','admin:export'),('⚙️ Sozlamalar','admin:settings')],[('🏠 User menyu','home')]])
def deps_kb(prefix:str, back='home'):
    rows=[[ (d['name'][:40], f'{prefix}:{d["department_key"]}') ] for d in departments()]
    rows.append([('⬅️ Orqaga',back)]); return ik(rows)
def teachers_kb(dkey:str,prefix:str,back:str):
    rows=[[ (t['name'][:45], f'{prefix}:{dkey}:{t["teacher_key"]}') ] for t in teachers(dkey)]
    rows.append([('⬅️ Orqaga',back)]); return ik(rows)
def rating_stars_kb(dkey,tkey):
    return ik([[('⭐ 1',f'rate:save:{dkey}:{tkey}:1'),('⭐⭐ 2',f'rate:save:{dkey}:{tkey}:2')],[('⭐⭐⭐ 3',f'rate:save:{dkey}:{tkey}:3')],[('⭐⭐⭐⭐ 4',f'rate:save:{dkey}:{tkey}:4'),('⭐⭐⭐⭐⭐ 5',f'rate:save:{dkey}:{tkey}:5')],[('⬅️ O‘qituvchilar',f'rate:dep:{dkey}')]])
def phone_kb(): return ReplyKeyboardMarkup(keyboard=[[KeyboardButton(text='📱 Telefon raqamni yuborish', request_contact=True)]], resize_keyboard=True, one_time_keyboard=True)

def admin_teachers_kb(): return ik([[('➕ Qo‘shish','tm:add'),('📋 Ro‘yxat','tm:list')],[('🔍 Qidirish','tm:search'),('👥 O‘quvchilar soni','tm:students')],[('⬅️ Admin','admin:menu')]])
def admin_settings_kb(): return ik([[('👨‍🏫 O‘qituvchilar','admin:teachers'),('🏛 Kafedralar','set:deps')],[('📢 Shikoyatlar','set:complaints'),('🗳 Baholash','set:voting')],[('📊 Reyting','set:rating'),('🔐 Adminlar','set:admins')],[('📦 Backup','set:backup')],[('⬅️ Admin','admin:menu')]])

# =====================================================
# TEXTS
# =====================================================
MOTIVATION = """<b>Sizning fikringiz ta’lim sifatini oshirish uchun muhim.</b>\n\nHar bir berilgan baho o‘qituvchilar faoliyatini xolis tahlil qilish, kuchli jihatlarni aniqlash va rivojlantirilishi kerak bo‘lgan yo‘nalishlarni ko‘rish uchun xizmat qiladi.\n\nIltimos, baholashda adolatli, xolis va mas’uliyatli bo‘ling. Sizning javobingiz umumiy reyting va sifat ko‘rsatkichlariga ta’sir qiladi."""
def home_text(): return "<b>Bosh menyu</b>\nKerakli bo‘limni tanlang."
def info_text(): return "<b>Ma’lumot</b>\n\nBot o‘qituvchilar faoliyatini 1–5 yulduzli tizimda baholash, reytinglarni ko‘rish va shikoyat yuborish uchun mo‘ljallangan.\n\nFormula: <code>final_score = average_rating × (total_votes / student_count)</code>"

def teacher_stats_text(dkey,tkey):
    st=teacher_stats(dkey,tkey); dist='\n'.join([f"{i}⭐: {st['dist'][i]} ta" for i in range(1,6)])
    return f"<b>O‘qituvchi statistikasi</b>\n\n<b>O‘qituvchi:</b> {safe(teacher_name(dkey,tkey))}\n<b>Kafedra:</b> {safe(dep_name(dkey))}\n\nO‘rtacha baho: <b>{st['avg']:.2f}</b>\nOvozlar soni: <b>{st['total']}</b>\nO‘quvchilar soni: <b>{st['student_count']}</b>\nQatnashuv: <b>{st['participation']:.1f}%</b>\nFinal score: <b>{st['final']:.3f}</b>\n\n<b>Rating distribution</b>\n{dist}"

def top_rating_text(limit=20):
    data=[]
    for t in all_teachers():
        st=teacher_stats(t['department_key'],t['teacher_key'])
        data.append((st['final'],st['avg'],st['total'],t))
    data.sort(key=lambda x:(x[0],x[2],x[1]), reverse=True)
    lines=["🏆 <b>O‘qituvchilar reytingi</b>\n"]
    for i,(final,avg,total,t) in enumerate(data[:limit],1):
        lines.append(f"{i}. <b>{safe(t['name'])}</b>\n{safe(t['department_name'])}\n⭐ {avg:.2f} • {total} baho • score: <b>{final:.3f}</b>")
    return '\n\n'.join(lines)

def departments_rating_text():
    data=[]
    for d in departments():
        st=department_stats(d['department_key']); data.append((st['final'],d,st))
    data.sort(key=lambda x:x[0], reverse=True)
    lines=["🏛 <b>Kafedralar reytingi</b>\n"]
    for i,(score,d,st) in enumerate(data,1):
        lines.append(f"{i}. <b>{safe(d['name'])}</b>\nO‘qituvchi: {st['teachers']} • Baholar: {st['votes']}\nO‘rtacha: {st['avg']:.2f} • Qatnashuv: {st['participation']:.1f}%\nFinal score: <b>{st['final']:.3f}</b>")
    return '\n\n'.join(lines)

# =====================================================
# EXPORT / BACKUP
# =====================================================
def style_ws(ws):
    if Workbook is None: return
    header_fill=PatternFill("solid", fgColor="1F4E78"); header_font=Font(color="FFFFFF", bold=True)
    thin=Side(style="thin", color="D9E2F3")
    for cell in ws[1]: cell.fill=header_fill; cell.font=header_font; cell.alignment=Alignment(horizontal="center"); cell.border=Border(bottom=thin)
    for col in ws.columns:
        max_len=max(len(str(c.value or "")) for c in col)
        ws.column_dimensions[get_column_letter(col[0].column)].width=min(max(max_len+2,12),45)
    ws.freeze_panes="A2"

def export_all_excel() -> Path:
    if Workbook is None:
        path=DATA_DIR/"export.csv"; return path
    wb=Workbook(); wb.remove(wb.active)
    sheets = {
        "Users": ("SELECT telegram_id,first_name,last_name,fullname,username,phone,registered_at FROM users", ["Telegram ID","Ism","Familiya","FISH","Username","Telefon","Ro‘yxat"]),
        "Departments": ("SELECT department_key,name,sort_order,is_active FROM departments", ["Key","Nomi","Tartib","Active"]),
        "Teachers": ("SELECT department_key,teacher_key,name,student_count,is_active FROM teachers", ["Kafedra","Teacher key","O‘qituvchi","O‘quvchilar soni","Active"]),
        "Ratings": ("SELECT user_id,department_key,teacher_key,rating,fullname,phone,username,rated_at FROM ratings ORDER BY rated_at DESC", ["User ID","Kafedra","Teacher key","Baho","FISH","Telefon","Username","Sana"]),
        "Complaints": ("SELECT id,user_id,department_key,teacher_key,fullname,phone,complaint_text,status,created_at FROM complaints ORDER BY id DESC", ["ID","User ID","Kafedra","Teacher key","FISH","Telefon","Shikoyat","Status","Sana"]),
    }
    for name,(query,headers) in sheets.items():
        ws=wb.create_sheet(name); ws.append(headers)
        for r in conn.execute(query): ws.append(list(r))
        style_ws(ws)
    ws=wb.create_sheet("Teacher Stats"); ws.append(["Kafedra","O‘qituvchi","Average","Votes","Students","Participation %","Final score","1⭐","2⭐","3⭐","4⭐","5⭐"])
    for t in all_teachers(active=False):
        st=teacher_stats(t['department_key'],t['teacher_key'])
        ws.append([dep_name(t['department_key']),t['name'],round(st['avg'],2),st['total'],st['student_count'],round(st['participation'],2),round(st['final'],4),st['dist'][1],st['dist'][2],st['dist'][3],st['dist'][4],st['dist'][5]])
    style_ws(ws)
    path=DATA_DIR/f"full_export_{datetime.now(UZ_TZ).strftime('%Y%m%d_%H%M%S')}.xlsx"; wb.save(path); return path

def make_backup() -> Path:
    import zipfile
    export_path=export_all_excel()
    backup=BACKUP_DIR/f"backup_{datetime.now(UZ_TZ).strftime('%Y%m%d_%H%M%S')}.zip"
    conn.commit()
    with zipfile.ZipFile(backup,"w",zipfile.ZIP_DEFLATED) as z:
        if DB_PATH.exists(): z.write(DB_PATH, "votes.db")
        if export_path.exists(): z.write(export_path, export_path.name)
    return backup

# =====================================================
# ACCESS / SUBSCRIPTION
# =====================================================
async def check_sub(uid:int) -> bool:
    if not setting_bool("subscription_required", "1"): return True
    try:
        m=await bot.get_chat_member(CHANNEL_USERNAME, uid)
        return m.status in {ChatMemberStatus.MEMBER, ChatMemberStatus.ADMINISTRATOR, ChatMemberStatus.CREATOR}
    except Exception:
        return False

def sub_kb(): return ik([[('📢 Telegram kanal', 'url:noop')],[('✅ Tekshirish','check_sub')]])

# =====================================================
# USER HANDLERS
# =====================================================
@dp.message(Command("start"))
async def start(m: Message, state:FSMContext):
    await state.clear(); ensure_user_obj(m)
    if not await check_sub(m.from_user.id):
        kb=InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text='📢 Telegram kanal', url=CHANNEL_URL)],[InlineKeyboardButton(text='✅ Tekshirish', callback_data='check_sub')]])
        return await m.answer("Botdan foydalanish uchun Telegram kanalga obuna bo‘ling.", reply_markup=kb)
    await m.answer(home_text(), reply_markup=main_kb())

@dp.callback_query(F.data == 'check_sub')
async def cb_check(call: CallbackQuery):
    await answer_cb(call)
    if await check_sub(call.from_user.id): await edit_or_send(call, home_text(), main_kb())
    else: await call.answer("Avval kanalga obuna bo‘ling.", show_alert=True)

@dp.message(Command("admin"))
async def admin_cmd(m:Message):
    ensure_user_obj(m)
    if not is_admin(m.from_user.id): return await m.answer("Ruxsat yo‘q.")
    await m.answer("<b>Admin panel</b>", reply_markup=admin_kb())

@dp.callback_query(F.data == 'home')
async def home_cb(call:CallbackQuery): await answer_cb(call); await edit_or_send(call, home_text(), main_kb())
@dp.callback_query(F.data == 'info')
async def info_cb(call:CallbackQuery): await answer_cb(call); await edit_or_send(call, info_text(), ik([[('⬅️ Orqaga','home')]]))

async def start_registration(target, state:FSMContext, after:str):
    await state.set_state(Register.first_name); await state.update_data(after=after)
    await edit_or_send(target, "Baholashdan oldin ro‘yxatdan o‘ting.\n\nIsmingizni kiriting:")

@dp.message(Register.first_name)
async def reg_first(m:Message,state:FSMContext):
    await state.update_data(first_name=m.text.strip()); await state.set_state(Register.last_name); await m.answer("Familiyangizni kiriting:")
@dp.message(Register.last_name)
async def reg_last(m:Message,state:FSMContext):
    await state.update_data(last_name=m.text.strip()); await state.set_state(Register.phone); await m.answer("Telefon raqamingizni contact button orqali yuboring:", reply_markup=phone_kb())
@dp.message(Register.phone)
async def reg_phone(m:Message,state:FSMContext):
    if not m.contact or m.contact.user_id != m.from_user.id: return await m.answer("Iltimos, pastdagi contact tugmasi orqali o‘z telefon raqamingizni yuboring.", reply_markup=phone_kb())
    data=await state.get_data(); full=f"{data['first_name']} {data['last_name']}"
    conn.execute("""INSERT INTO users(telegram_id,first_name,last_name,fullname,username,phone,registered_at,updated_at)
        VALUES(?,?,?,?,?,?,?,?) ON CONFLICT(telegram_id) DO UPDATE SET first_name=excluded.first_name,last_name=excluded.last_name,fullname=excluded.fullname,username=excluded.username,phone=excluded.phone,updated_at=excluded.updated_at""",
        (m.from_user.id,data['first_name'],data['last_name'],full,m.from_user.username,m.contact.phone_number,now_str(),now_str()))
    conn.commit(); after=data.get('after','rate') ; await state.clear()
    await m.answer("Ro‘yxatdan o‘tish yakunlandi.", reply_markup=ReplyKeyboardRemove())
    if after=='complaint': await m.answer("Kafedrani tanlang:", reply_markup=deps_kb('complaint:dep','home'))
    else: await m.answer(MOTIVATION, reply_markup=deps_kb('rate:dep','home'))

@dp.callback_query(F.data == 'rate:start')
async def rate_start(call:CallbackQuery,state:FSMContext):
    await answer_cb(call)
    if not setting_bool('voting_open','1'): return await call.answer('Baholash hozircha yopilgan.', show_alert=True)
    if not user_registered(call.from_user.id): return await start_registration(call,state,'rate')
    await edit_or_send(call, MOTIVATION, deps_kb('rate:dep','home'))
@dp.callback_query(F.data.startswith('rate:dep:'))
async def rate_dep(call:CallbackQuery):
    await answer_cb(call); dkey=call.data.split(':',2)[2]
    await edit_or_send(call, f"<b>{safe(dep_name(dkey))}</b>\nO‘qituvchini tanlang:", teachers_kb(dkey,'rate:teacher','rate:start'))
@dp.callback_query(F.data.startswith('rate:teacher:'))
async def rate_teacher(call:CallbackQuery):
    await answer_cb(call); _,_,dkey,tkey=call.data.split(':',3)
    old=conn.execute("SELECT rating FROM ratings WHERE user_id=? AND department_key=? AND teacher_key=?",(call.from_user.id,dkey,tkey)).fetchone()
    txt=f"<b>{safe(teacher_name(dkey,tkey))}</b>\n\nBahoni tanlang." + (f"\nHozirgi baho: {old[0]}⭐" if old else "")
    await edit_or_send(call, txt, rating_stars_kb(dkey,tkey))
@dp.callback_query(F.data.startswith('rate:save:'))
async def rate_save(call:CallbackQuery):
    await answer_cb(call); _,_,dkey,tkey,val=call.data.split(':',4)
    if not user_registered(call.from_user.id): return await call.answer('Avval ro‘yxatdan o‘ting.', show_alert=True)
    save_rating(call.from_user.id,dkey,tkey,int(val))
    await edit_or_send(call, f"✅ Baho saqlandi: <b>{val}⭐</b>\n\n{safe(teacher_name(dkey,tkey))}", ik([[('Yana baholash','rate:start'),('🏆 Reyting','rating:top')],[('🏠 Menyu','home')]]))

@dp.callback_query(F.data == 'rating:top')
async def rating_top(call:CallbackQuery): await answer_cb(call); await edit_or_send(call, top_rating_text(), ik([[('🏛 Kafedralar','admin:dept_rating_public')],[('⬅️ Orqaga','home')]]))
@dp.callback_query(F.data == 'admin:dept_rating_public')
async def public_deps(call:CallbackQuery): await answer_cb(call); await edit_or_send(call, departments_rating_text(), ik([[('⬅️ Reyting','rating:top')]]))

@dp.callback_query(F.data == 'complaint:start')
async def complaint_start(call:CallbackQuery,state:FSMContext):
    await answer_cb(call)
    if not user_registered(call.from_user.id): return await start_registration(call,state,'complaint')
    await state.set_state(ComplaintFSM.department); await edit_or_send(call,"Kafedrani tanlang:", deps_kb('complaint:dep','home'))
@dp.callback_query(F.data.startswith('complaint:dep:'))
async def complaint_dep(call:CallbackQuery,state:FSMContext):
    await answer_cb(call); dkey=call.data.split(':',2)[2]; await state.update_data(dkey=dkey); await state.set_state(ComplaintFSM.teacher)
    await edit_or_send(call, "O‘qituvchini tanlang:", teachers_kb(dkey,'complaint:teacher','complaint:start'))
@dp.callback_query(F.data.startswith('complaint:teacher:'))
async def complaint_teacher(call:CallbackQuery,state:FSMContext):
    await answer_cb(call); _,_,dkey,tkey=call.data.split(':',3); await state.update_data(dkey=dkey,tkey=tkey); await state.set_state(ComplaintFSM.text)
    await edit_or_send(call, f"<b>{safe(teacher_name(dkey,tkey))}</b> ustidan shikoyat matnini yozing:")
@dp.message(ComplaintFSM.text)
async def complaint_text(m:Message,state:FSMContext):
    data=await state.get_data(); u=get_user(m.from_user.id); text=(m.text or '').strip()
    if len(text)<5: return await m.answer("Shikoyat matni juda qisqa.")
    conn.execute("""INSERT INTO complaints(user_id,teacher_id,department_key,teacher_key,fullname,phone,username,complaint_text,message_text,created_at,status)
        VALUES(?,?,?,?,?,?,?,?,?,?,?)""", (m.from_user.id,data['tkey'],data['dkey'],data['tkey'],u['fullname'] if u else '',u['phone'] if u else '',u['username'] if u else '',text,text,now_str(),'Yangi'))
    conn.commit(); await state.clear(); await m.answer("✅ Shikoyatingiz saqlandi.", reply_markup=main_kb())

# =====================================================
# ADMIN HANDLERS
# =====================================================
def admin_required(func):
    async def wrapper(event, *args, **kwargs):
        uid=event.from_user.id
        if not is_admin(uid):
            if isinstance(event, CallbackQuery): await event.answer('Ruxsat yo‘q.', show_alert=True)
            else: await event.answer('Ruxsat yo‘q.')
            return
        return await func(event,*args,**kwargs)
    return wrapper

@dp.callback_query(F.data == 'admin:menu')
@admin_required
async def admin_menu(call:CallbackQuery): await answer_cb(call); await edit_or_send(call,"<b>Admin panel</b>",admin_kb())
@dp.callback_query(F.data == 'admin:dash')
@admin_required
async def admin_dash(call:CallbackQuery):
    await answer_cb(call)
    users=conn.execute('SELECT COUNT(*) FROM users').fetchone()[0]; ratings=conn.execute('SELECT COUNT(*) FROM ratings').fetchone()[0]; comps=conn.execute('SELECT COUNT(*) FROM complaints').fetchone()[0]
    txt=f"📊 <b>Dashboard</b>\n\nFoydalanuvchilar: <b>{users}</b>\nBaholar: <b>{ratings}</b>\nShikoyatlar: <b>{comps}</b>\nBaholash: <b>{'Ochiq' if setting_bool('voting_open') else 'Yopiq'}</b>"
    await edit_or_send(call,txt,ik([[('⬅️ Admin','admin:menu')]]))
@dp.callback_query(F.data == 'admin:teachers')
@admin_required
async def adm_teachers(call:CallbackQuery): await answer_cb(call); await edit_or_send(call,"👨‍🏫 <b>O‘qituvchilar</b>",admin_teachers_kb())
@dp.callback_query(F.data == 'tm:list')
@admin_required
async def tm_list(call:CallbackQuery): await answer_cb(call); await edit_or_send(call,"Kafedrani tanlang:",deps_kb('tm:listdep','admin:teachers'))
@dp.callback_query(F.data.startswith('tm:listdep:'))
@admin_required
async def tm_listdep(call:CallbackQuery):
    await answer_cb(call); dkey=call.data.split(':',2)[2]
    lines=[f"<b>{safe(dep_name(dkey))}</b>\n"]
    for t in teachers(dkey,False): lines.append(f"• {safe(t['name'])} — o‘quvchi: {t['student_count'] or 0} {'✅' if t['is_active'] else '❌'}")
    await edit_or_send(call,'\n'.join(lines)[:4000],ik([[('⬅️ Orqaga','tm:list')]]))
@dp.callback_query(F.data == 'tm:add')
@admin_required
async def tm_add(call:CallbackQuery,state:FSMContext): await answer_cb(call); await state.set_state(TeacherFSM.add_department); await edit_or_send(call,"Qaysi kafedraga qo‘shiladi?",deps_kb('tm:adddep','admin:teachers'))
@dp.callback_query(F.data.startswith('tm:adddep:'))
@admin_required
async def tm_adddep(call:CallbackQuery,state:FSMContext): await answer_cb(call); dkey=call.data.split(':',2)[2]; await state.update_data(dkey=dkey); await state.set_state(TeacherFSM.add_name); await edit_or_send(call,"Yangi o‘qituvchi F.I.Sh ni kiriting:")
@dp.message(TeacherFSM.add_name)
@admin_required
async def tm_addname(m:Message,state:FSMContext):
    data=await state.get_data(); name=m.text.strip(); key=slug(name,'t')
    conn.execute("INSERT OR IGNORE INTO teachers(teacher_key,department_key,name,created_at) VALUES(?,?,?,?)",(key,data['dkey'],name,now_str())); conn.commit(); await state.clear(); await m.answer("✅ O‘qituvchi qo‘shildi.",reply_markup=admin_teachers_kb())
@dp.callback_query(F.data == 'tm:students')
@admin_required
async def tm_students(call:CallbackQuery): await answer_cb(call); await edit_or_send(call,"Kafedrani tanlang:",deps_kb('tm:studentsdep','admin:teachers'))
@dp.callback_query(F.data.startswith('tm:studentsdep:'))
@admin_required
async def tm_studentsdep(call:CallbackQuery): await answer_cb(call); dkey=call.data.split(':',2)[2]; await edit_or_send(call,"O‘qituvchini tanlang:",teachers_kb(dkey,'tm:studentteacher','tm:students'))
@dp.callback_query(F.data.startswith('tm:studentteacher:'))
@admin_required
async def tm_studentteacher(call:CallbackQuery,state:FSMContext):
    await answer_cb(call); _,_,dkey,tkey=call.data.split(':',3); await state.update_data(dkey=dkey,tkey=tkey); await state.set_state(TeacherFSM.student_count)
    st=teacher_stats(dkey,tkey); action='✏️ O‘zgartirish' if st['student_count'] else '➕ Kiritish'
    await edit_or_send(call,f"{safe(teacher_name(dkey,tkey))}\nHozirgi o‘quvchilar soni: <b>{st['student_count']}</b>\n\n{action}: yangi sonni kiriting.")
@dp.message(TeacherFSM.student_count)
@admin_required
async def tm_student_save(m:Message,state:FSMContext):
    if not m.text.isdigit(): return await m.answer("Faqat raqam kiriting.")
    data=await state.get_data(); conn.execute("UPDATE teachers SET student_count=? WHERE department_key=? AND teacher_key=?",(int(m.text),data['dkey'],data['tkey'])); conn.commit(); await state.clear(); await m.answer("✅ Saqlandi.",reply_markup=admin_teachers_kb())
@dp.callback_query(F.data == 'tm:search')
@admin_required
async def tm_search(call:CallbackQuery,state:FSMContext): await answer_cb(call); await state.set_state(TeacherFSM.search); await edit_or_send(call,"Qidirish uchun ism/familiya kiriting:")
@dp.message(TeacherFSM.search)
@admin_required
async def tm_search_msg(m:Message,state:FSMContext):
    q=f"%{m.text.strip()}%"; rows=conn.execute("SELECT t.*,d.name dep FROM teachers t JOIN departments d ON d.department_key=t.department_key WHERE t.name LIKE ? LIMIT 30",(q,)).fetchall(); await state.clear()
    if not rows: return await m.answer("Topilmadi.",reply_markup=admin_teachers_kb())
    await m.answer('\n\n'.join([f"<b>{safe(r['name'])}</b>\n{safe(r['dep'])}\nO‘quvchi: {r['student_count'] or 0}" for r in rows]),reply_markup=admin_teachers_kb())

@dp.callback_query(F.data == 'admin:tstats')
@admin_required
async def adm_tstats(call:CallbackQuery): await answer_cb(call); await edit_or_send(call,"Kafedrani tanlang:",deps_kb('stats:dep','admin:menu'))
@dp.callback_query(F.data.startswith('stats:dep:'))
@admin_required
async def stats_dep(call:CallbackQuery): await answer_cb(call); dkey=call.data.split(':',2)[2]; await edit_or_send(call,"O‘qituvchini tanlang:",teachers_kb(dkey,'stats:teacher','admin:tstats'))
@dp.callback_query(F.data.startswith('stats:teacher:'))
@admin_required
async def stats_teacher(call:CallbackQuery):
    await answer_cb(call); _,_,dkey,tkey=call.data.split(':',3)
    kb=ik([[('📢 Shikoyatlar',f'complaints:teacher:{dkey}:{tkey}')],[('✏️ Tahrirlash',f'edit_teacher:{dkey}:{tkey}'),('🗑 O‘chirish',f'del_teacher:{dkey}:{tkey}')],[('⬅️ Orqaga',f'stats:dep:{dkey}')]])
    await edit_or_send(call,teacher_stats_text(dkey,tkey),kb)
@dp.callback_query(F.data.startswith('complaints:teacher:'))
@admin_required
async def complaints_teacher(call:CallbackQuery):
    await answer_cb(call); _,_,dkey,tkey=call.data.split(':',3)
    rows=conn.execute("SELECT * FROM complaints WHERE department_key=? AND teacher_key=? ORDER BY id DESC LIMIT 20",(dkey,tkey)).fetchall()
    txt=f"📢 <b>{safe(teacher_name(dkey,tkey))} ustidan shikoyatlar</b>\n\n" + ('Hali mavjud emas.' if not rows else '\n\n'.join([f"#{r['id']} • {safe(r['status'])} • {safe(r['created_at'])}\n{safe(r['complaint_text'] or r['message_text'])}" for r in rows]))
    await edit_or_send(call,txt[:4000],ik([[('⬅️ Orqaga',f'stats:teacher:{dkey}:{tkey}')]]))
@dp.callback_query(F.data.startswith('del_teacher:'))
@admin_required
async def del_teacher(call:CallbackQuery): await answer_cb(call); _,dkey,tkey=call.data.split(':',2); conn.execute("UPDATE teachers SET is_active=0 WHERE department_key=? AND teacher_key=?",(dkey,tkey)); conn.commit(); await edit_or_send(call,"✅ O‘qituvchi o‘chirildi (arxivlandi).",admin_teachers_kb())

@dp.callback_query(F.data == 'admin:dept_rating')
@admin_required
async def adm_dep_rating(call:CallbackQuery): await answer_cb(call); await edit_or_send(call,departments_rating_text(),ik([[('⬅️ Admin','admin:menu')]]))
@dp.callback_query(F.data == 'admin:complaints')
@admin_required
async def adm_complaints(call:CallbackQuery):
    await answer_cb(call); rows=conn.execute("SELECT * FROM complaints ORDER BY id DESC LIMIT 30").fetchall()
    txt="📢 <b>Shikoyatlar</b>\n\n" + ('Mavjud emas.' if not rows else '\n\n'.join([f"#{r['id']} • {safe(r['status'])}\n{safe(r['fullname'])} | {safe(r['phone'])}\n{safe(dep_name(r['department_key']))} / {safe(teacher_name(r['department_key'],r['teacher_key']))}\n{safe(r['complaint_text'] or r['message_text'])}" for r in rows]))
    await edit_or_send(call,txt[:4000],ik([[('⬅️ Admin','admin:menu')]]))
@dp.callback_query(F.data == 'admin:export')
@admin_required
async def adm_export(call:CallbackQuery): await answer_cb(call); await edit_or_send(call,"📁 <b>Export / Backup</b>",ik([[('Excel export','export:excel'),('Backup ZIP','export:backup')],[('⬅️ Admin','admin:menu')]]))
@dp.callback_query(F.data == 'export:excel')
@admin_required
async def export_excel_cb(call:CallbackQuery):
    await answer_cb(call,'Tayyorlanmoqda...'); path=export_all_excel(); await call.message.answer_document(FSInputFile(path), caption='Excel export tayyor.')
@dp.callback_query(F.data == 'export:backup')
@admin_required
async def export_backup_cb(call:CallbackQuery):
    await answer_cb(call,'Backup tayyorlanmoqda...'); path=make_backup(); await call.message.answer_document(FSInputFile(path), caption='Backup ZIP tayyor.')
@dp.callback_query(F.data == 'admin:settings')
@admin_required
async def adm_settings(call:CallbackQuery): await answer_cb(call); await edit_or_send(call,"⚙️ <b>Sozlamalar</b>",admin_settings_kb())
@dp.callback_query(F.data == 'set:voting')
@admin_required
async def set_voting(call:CallbackQuery):
    await answer_cb(call); cur=setting_bool('voting_open'); set_setting('voting_open','0' if cur else '1')
    await edit_or_send(call,f"Baholash holati: <b>{'Ochiq' if not cur else 'Yopiq'}</b>",admin_settings_kb())
@dp.callback_query(F.data.startswith('set:'))
@admin_required
async def set_placeholders(call:CallbackQuery):
    await answer_cb(call); section=call.data.split(':',1)[1]
    await edit_or_send(call,f"⚙️ <b>{section}</b> sozlamalari tayyor. Zarur amallar admin paneldagi tegishli bo‘limlar orqali boshqariladi.",admin_settings_kb())

# legacy/admin export commands
@dp.message(Command('backup'))
@admin_required
async def backup_cmd(m:Message): path=make_backup(); await m.answer_document(FSInputFile(path), caption='Backup ZIP')
@dp.message(Command('export'))
@admin_required
async def export_cmd(m:Message): path=export_all_excel(); await m.answer_document(FSInputFile(path), caption='Excel export')

# fallback buttons from old bot
@dp.message(F.text.in_({'📝 O‘qituvchini baholash','📝 Baholash'}))
async def msg_rate(m:Message,state:FSMContext): ensure_user_obj(m); await start_registration(m,state,'rate') if not user_registered(m.from_user.id) else await m.answer(MOTIVATION,reply_markup=deps_kb('rate:dep','home'))
@dp.message(F.text.in_({'📢 Shikoyat yuborish','📢 Shikoyat'}))
async def msg_complaint(m:Message,state:FSMContext): ensure_user_obj(m); await start_registration(m,state,'complaint') if not user_registered(m.from_user.id) else await m.answer('Kafedrani tanlang:',reply_markup=deps_kb('complaint:dep','home'))

@dp.errors()
async def errors_handler(event):
    log.exception("Update error: %s", event.exception)
    return True

async def daily_backup_job():
    if setting_bool('daily_backup','1'):
        try: make_backup(); log.info('Daily backup created')
        except Exception: log.exception('Daily backup failed')

async def main():
    init_db()
    if AsyncIOScheduler:
        scheduler=AsyncIOScheduler(timezone='Asia/Tashkent')
        scheduler.add_job(daily_backup_job, 'cron', hour=3, minute=0)
        scheduler.start()
    log.info('Bot started')
    await dp.start_polling(bot, allowed_updates=dp.resolve_used_update_types())

if __name__ == '__main__':
    asyncio.run(main())
