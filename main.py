
# Talaba Bot - Mustaqil ish, Referat, Kurs ishi va boshqalar
# Python + aiogram 3 + OpenRouter (bepul) + Render uchun tayyor
# Ishga tushirish: pip install -r requirements.txt  -> python main.py

import asyncio, os, json, sqlite3
from datetime import datetime
from aiogram import Bot, Dispatcher, F
from aiogram.types import Message, CallbackQuery, InlineKeyboardMarkup, InlineKeyboardButton, FSInputFile
from aiogram.filters import CommandStart, Command
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
import aiohttp
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

# ===== SOZLAMALAR - .env dan olinadi, Render da oson o'zgaradi =====
BOT_TOKEN = os.getenv("BOT_TOKEN", "TOKENINGNI_BUYERGA_QOY")
OPENROUTER_KEY = os.getenv("OPENROUTER_KEY", "")
ADMIN_ID = int(os.getenv("ADMIN_ID", "0"))  # sening Telegram ID ing
CARD_NUMBER = os.getenv("CARD_NUMBER", "8600 XXXX XXXX XXXX - Isming")
CARD_OWNER = os.getenv("CARD_OWNER", "Isming")

# Narxlar - admin botda o'zgartirishi mumkin, bazada saqlanadi
DEFAULT_PRICES = {"5-7": 10000, "10-15": 20000, "20+": 30000}

# ===== BAZA =====
db = sqlite3.connect("bot.db", check_same_thread=False)
cur = db.cursor()
cur.execute("CREATE TABLE IF NOT EXISTS users (user_id INTEGER PRIMARY KEY, name TEXT, balance INTEGER DEFAULT 0, free_used INTEGER DEFAULT 0, created TEXT)")
cur.execute("CREATE TABLE IF NOT EXISTS prices (key TEXT PRIMARY KEY, value INTEGER)")
cur.execute("CREATE TABLE IF NOT EXISTS payments (id INTEGER PRIMARY KEY AUTOINCREMENT, user_id INTEGER, amount INTEGER, status TEXT, created TEXT)")
for k,v in DEFAULT_PRICES.items():
    cur.execute("INSERT OR IGNORE INTO prices VALUES (?,?)", (k,v))
db.commit()

def get_price(key):
    cur.execute("SELECT value FROM prices WHERE key=?", (key,))
    r = cur.fetchone()
    return r[0] if r else DEFAULT_PRICES[key]

def get_user(uid):
    cur.execute("SELECT * FROM users WHERE user_id=?", (uid,))
    r = cur.fetchone()
    if not r:
        cur.execute("INSERT INTO users VALUES (?,?,?,?,?)", (uid, "", 0, 0, datetime.now().isoformat()))
        db.commit()
        return (uid, "", 0, 0, "")
    return r

def add_balance(uid, amount):
    cur.execute("UPDATE users SET balance=balance+? WHERE user_id=?", (amount, uid))
    db.commit()

# ===== AI - OpenRouter bepul DeepSeek R1 =====
async def ask_ai(prompt: str) -> str:
    if not OPENROUTER_KEY:
        return "AI kalit qo'yilmagan. Admin OPENROUTER_KEY ni sozlasin."
    url = "https://openrouter.ai/api/v1/chat/completions"
    headers = {"Authorization": f"Bearer {OPENROUTER_KEY}", "Content-Type": "application/json"}
    payload = {
        "model": "deepseek/deepseek-r1:free",
        "messages": [
            {"role": "system", "content": "Sen universitet o'qituvchisisan. O'zbek tilida, ilmiy uslubda, aniq, takrorlarsiz, talaba tilida tushunarli yozasan. Raqamlarni o'ylab topmaysan."},
            {"role": "user", "content": prompt}
        ]
    }
    async with aiohttp.ClientSession() as s:
        async with s.post(url, json=payload, headers=headers) as resp:
            data = await resp.json()
            try:
                return data["choices"][0]["message"]["content"]
            except:
                return f"AI da xatolik: {data}"

def make_docx(title, text, user_info):
    doc = Document()
    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.font.size = Pt(14)
    # Titul
    for line in [user_info.get("univer",""), user_info.get("fan",""), "", "", "MUSTAQIL ISH", f"Mavzu: {title}", "", "", f"Talaba: {user_info.get('name','')}", f"Guruh: {user_info.get('group','')}", f"O'qituvchi: {user_info.get('teacher','')}"]:
        p = doc.add_paragraph(line)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_page_break()
    for para in text.split("\n"):
        if para.strip():
            doc.add_paragraph(para.strip())
    path = f"ish_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
    doc.save(path)
    return path

# ===== HOLATLAR =====
class Order(StatesGroup):
    lang = State(); work_type = State(); univer = State(); fan = State(); mavzu = State()
    name = State(); group = State(); teacher = State(); hajm = State(); confirm = State()

class TopUp(StatesGroup):
    amount = State(); check = State()

class AdminPrice(StatesGroup):
    key = State(); value = State()

bot = Bot(token=BOT_TOKEN)
dp = Dispatcher()

def main_kb():
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="📝 Ish yozdirish", callback_data="order")],
        [InlineKeyboardButton(text="💰 Balansni to'ldirish", callback_data="topup"), InlineKeyboardButton(text="👤 Balansim", callback_data="mybalance")],
        [InlineKeyboardButton(text="ℹ️ Yordam", callback_data="help")],
    ])

@dp.message(CommandStart())
async def start(m: Message):
    get_user(m.from_user.id)
    await m.answer("Assalomu alaykum! 🎓\n\nMen Talaba Botman - Mustaqil ish, Referat, Kurs ishi, Slayd, Konspekt yozib beraman.\n\n1-ish - BEPUL! Keyingilari balansdan.\n\nBalansni to'ldirib, ishingizni yozdiring.", reply_markup=main_kb())

@dp.callback_query(F.data=="mybalance")
async def mybal(c: CallbackQuery):
    u = get_user(c.from_user.id)
    await c.message.answer(f"💰 Balansingiz: {u[2]} so'm\n🎁 Bepul ish: {'ishlatilgan' if u[3] else 'ishlatilmagan - 1 ta bepul bor!'}", reply_markup=main_kb())
    await c.answer()

@dp.callback_query(F.data=="order")
async def order_start(c: CallbackQuery, state: FSMContext):
    await state.set_state(Order.lang)
    kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="🇺🇿 O'zbekcha", callback_data="lang_uz")],
        [InlineKeyboardButton(text="🇷🇺 Ruscha", callback_data="lang_ru")],
        [InlineKeyboardButton(text="🇬🇧 Inglizcha", callback_data="lang_en")],
    ])
    await c.message.answer("1️⃣ Tilni tanlang:", reply_markup=kb)
    await c.answer()

@dp.callback_query(F.data.startswith("lang_"))
async def lang_set(c: CallbackQuery, state: FSMContext):
    await state.update_data(lang=c.data.replace("lang_",""))
    await state.set_state(Order.work_type)
    kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="Mustaqil ish", callback_data="wt_Mustaqil ish")],
        [InlineKeyboardButton(text="Referat", callback_data="wt_Referat")],
        [InlineKeyboardButton(text="Kurs ishi", callback_data="wt_Kurs ishi")],
        [InlineKeyboardButton(text="Slayd", callback_data="wt_Slayd")],
        [InlineKeyboardButton(text="Konspekt", callback_data="wt_Konspekt")],
        [InlineKeyboardButton(text="Laboratoriya ishi", callback_data="wt_Laboratoriya ishi")],
    ])
    await c.message.answer("2️⃣ Ish turini tanlang:", reply_markup=kb)
    await c.answer()

@dp.callback_query(F.data.startswith("wt_"))
async def wt_set(c: CallbackQuery, state: FSMContext):
    await state.update_data(work_type=c.data[3:])
    await state.set_state(Order.univer)
    await c.message.answer("3️⃣ Universitet nomini yozing:")
    await c.answer()

@dp.message(Order.univer)
async def univer_set(m: Message, state: FSMContext):
    await state.update_data(univer=m.text)
    await state.set_state(Order.fan)
    await m.answer("4️⃣ Fan nomini yozing:")

@dp.message(Order.fan)
async def fan_set(m: Message, state: FSMContext):
    await state.update_data(fan=m.text)
    await state.set_state(Order.mavzu)
    await m.answer("5️⃣ Mavzuni yozing:")

@dp.message(Order.mavzu)
async def mavzu_set(m: Message, state: FSMContext):
    await state.update_data(mavzu=m.text)
    await state.set_state(Order.name)
    await m.answer("6️⃣ Ism Familiyangiz:")

@dp.message(Order.name)
async def name_set(m: Message, state: FSMContext):
    await state.update_data(name=m.text)
    await state.set_state(Order.group)
    await m.answer("7️⃣ Guruhingiz:")

@dp.message(Order.group)
async def group_set(m: Message, state: FSMContext):
    await state.update_data(group=m.text)
    await state.set_state(Order.teacher)
    await m.answer("8️⃣ O'qituvchi ismi:")

@dp.message(Order.teacher)
async def teacher_set(m: Message, state: FSMContext):
    await state.update_data(teacher=m.text)
    await state.set_state(Order.hajm)
    kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text=f"5-7 bet - {get_price('5-7')} so'm", callback_data="hajm_5-7")],
        [InlineKeyboardButton(text=f"10-15 bet - {get_price('10-15')} so'm", callback_data="hajm_10-15")],
        [InlineKeyboardButton(text=f"20+ bet - {get_price('20+')} so'm", callback_data="hajm_20+")],
    ])
    await m.answer("9️⃣ Hajmni tanlang:", reply_markup=kb)

@dp.callback_query(F.data.startswith("hajm_"))
async def hajm_set(c: CallbackQuery, state: FSMContext):
    hajm = c.data.replace("hajm_","")
    data = await state.get_data()
    u = get_user(c.from_user.id)
    is_free = (u[3]==0)
    price = 0 if is_free else get_price(hajm)
    await state.update_data(hajm=hajm, price=price, is_free=is_free)
    txt = f"📋 Tekshiring:\n\nIsh: {data['work_type']}\nFan: {data['fan']}\nMavzu: {data['mavzu']}\nHajm: {hajm}\nNarx: {'BEPUL 🎁' if is_free else f'{price} so\\'m'}\nBalansingiz: {u[2]} so'm"
    kb = InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text="✅ Tasdiqlash, yozishni boshlash", callback_data="confirm_yes")]])
    await c.message.answer(txt, reply_markup=kb)
    await c.answer()

@dp.callback_query(F.data=="confirm_yes")
async def confirm(c: CallbackQuery, state: FSMContext):
    data = await state.get_data()
    u = get_user(c.from_user.id)
    if not data.get("is_free"):
        if u[2] < data["price"]:
            await c.message.answer(f"❌ Balans yetmaydi. Balansingiz: {u[2]} so'm, kerak: {data['price']} so'm. Avval balansni to'ldiring.", reply_markup=main_kb())
            await state.clear()
            await c.answer()
            return
        cur.execute("UPDATE users SET balance=balance-? WHERE user_id=?", (data["price"], c.from_user.id))
        db.commit()
    else:
        cur.execute("UPDATE users SET free_used=1 WHERE user_id=?", (c.from_user.id,))
        db.commit()
    await c.message.answer("⏳ Ishingiz yozilmoqda... 1-2 daqiqa kuting, tayyor bo'lishi bilan yuboraman.")
    prompt = f"{data['lang']} tilida {data['work_type']} yoz. Fan: {data['fan']}. Mavzu: {data['mavzu']}. Hajm: {data['hajm']} bet. Tuzilma: Kirish, 2-3 bob, Xulosa, Adabiyotlar. Ilmiy uslubda, batafsil yoz."
    text = await ask_ai(prompt)
    path = make_docx(data['mavzu'], text, data)
    await c.message.answer_document(FSInputFile(path), caption=f"✅ Tayyor! {data['work_type']} - {data['mavzu']}\n\nYana ish yozdirish uchun /start ni bosing.")
    await state.clear()
    await c.answer()

# ===== BALANS TO'LDIRISH - QO'LDA TEKSHIRISH =====
@dp.callback_query(F.data=="topup")
async def topup_start(c: CallbackQuery, state: FSMContext):
    await state.set_state(TopUp.amount)
    await c.message.answer(f"💳 Balansni to'ldirish:\n\nKarta: {CARD_NUMBER}\nEgasi: {CARD_OWNER}\n\nQancha to'ldirmoqchisiz? Summani yozing (masalan: 50000):")
    await c.answer()

@dp.message(TopUp.amount)
async def topup_amount(m: Message, state: FSMContext):
    try:
        amt = int(m.text.replace(" ","").replace(".",""))
    except:
        await m.answer("Iltimos, faqat raqam yozing. Masalan: 50000")
        return
    await state.update_data(amount=amt)
    await state.set_state(TopUp.check)
    await m.answer(f"{amt} so'mni kartaga tashlab, chekni shu yerga rasm qilib yuboring. Admin tekshirib, balansingizni to'ldiradi.")

@dp.message(TopUp.check, F.photo)
async def topup_check(m: Message, state: FSMContext):
    data = await state.get_data()
    cur.execute("INSERT INTO payments (user_id, amount, status, created) VALUES (?,?,?,?)", (m.from_user.id, data["amount"], "pending", datetime.now().isoformat()))
    db.commit()
    pid = cur.lastrowid
    await m.answer("✅ Chekingiz qabul qilindi. Admin tekshirib, balansingizni to'ldiradi. Biroz kuting.")
    if ADMIN_ID:
        kb = InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text="✅ Tasdiqlash", callback_data=f"approve_{pid}_{m.from_user.id}_{data['amount']}")]])
        await bot.send_photo(ADMIN_ID, m.photo[-1].file_id, caption=f"Yangi to'lov!\nFoydalanuvchi: {m.from_user.id} @{m.from_user.username}\nSumma: {data['amount']} so'm", reply_markup=kb)
    await state.clear()

@dp.callback_query(F.data.startswith("approve_"))
async def approve(c: CallbackQuery):
    _, pid, uid, amt = c.data.split("_")
    add_balance(int(uid), int(amt))
    cur.execute("UPDATE payments SET status='approved' WHERE id=?", (pid,))
    db.commit()
    await bot.send_message(int(uid), f"✅ Balansingiz to'ldirildi: +{amt} so'm. Endi ish yozdirishingiz mumkin.", reply_markup=main_kb())
    await c.message.answer("✅ Tasdiqlandi.")
    await c.answer()

# ===== ADMIN - NARXNI BOTDA O'ZGARTIRISH =====
@dp.message(Command("admin"))
async def admin(m: Message):
    if m.from_user.id != ADMIN_ID:
        await m.answer("Siz admin emassiz.")
        return
    kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text=f"5-7 bet: {get_price('5-7')}", callback_data="price_5-7")],
        [InlineKeyboardButton(text=f"10-15 bet: {get_price('10-15')}", callback_data="price_10-15")],
        [InlineKeyboardButton(text=f"20+ bet: {get_price('20+')}", callback_data="price_20+")],
        [InlineKeyboardButton(text="📊 Statistika", callback_data="stats")],
    ])
    await m.answer("⚙️ Admin panel:", reply_markup=kb)

@dp.callback_query(F.data.startswith("price_"))
async def price_edit(c: CallbackQuery, state: FSMContext):
    key = c.data.replace("price_","")
    await state.update_data(price_key=key)
    await state.set_state(AdminPrice.value)
    await c.message.answer(f"{key} uchun yangi narxni yozing:")
    await c.answer()

@dp.message(AdminPrice.value)
async def price_save(m: Message, state: FSMContext):
    data = await state.get_data()
    try:
        v = int(m.text)
    except:
        await m.answer("Faqat raqam yozing.")
        return
    cur.execute("UPDATE prices SET value=? WHERE key=?", (v, data["price_key"]))
    db.commit()
    await m.answer(f"✅ Yangi narx saqlandi: {data['price_key']} = {v} so'm")
    await state.clear()

@dp.callback_query(F.data=="stats")
async def stats(c: CallbackQuery):
    cur.execute("SELECT COUNT(*), SUM(balance) FROM users")
    u = cur.fetchone()
    cur.execute("SELECT COUNT(*), SUM(amount) FROM payments WHERE status='approved'")
    p = cur.fetchone()
    await c.message.answer(f"📊 Statistika:\nFoydalanuvchilar: {u[0]}\nJami balans: {u[1] or 0} so'm\nTasdiqlangan to'lovlar: {p[0]} ta, {p[1] or 0} so'm")
    await c.answer()

async def main():
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
