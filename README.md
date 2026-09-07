# 🎓 Sorovnoma Bot

Maktab o'quvchilarini ro'yxatga oluvchi Telegram bot (Python + aiogram 3 + SQLite).

## Imkoniyatlar

- Bitta Telegram akkauntdan **bir nechta o'quvchini** ro'yxatdan o'tkazish
- JSHSHIR **takrorlanmasligi** (bir xil JSHSHIR ikkinchi marta qabul qilinmaydi)
- Har bir o'quvchini alohida **tahrirlash** (faqat kerakli maydonni) va **o'chirish**
- Tug'ilgan sana / telefon / F.I.Sh. **validatsiyasi**, JSHSHIR ichidagi sanani tekshirish
- Admin panel: 📊 statistika, 🔎 qidiruv, 📢 broadcast, 📥 chiroyli Excel (sinf bo'yicha), 🗄 zaxira
- Yangi ariza kelganda **adminga bildirishnoma**

## Buyruqlar

| Buyruq | Kim uchun | Vazifa |
|--------|-----------|--------|
| `/start` | Hamma | Boshlash / o'quvchilarni boshqarish |
| `/cancel` | Hamma | Joriy jarayonni bekor qilish |
| `/help` | Hamma | Yordam |
| `/admin` | Admin | Admin panel |
| `/stats` | Admin | Statistika |
| `/search [so'z]` | Admin | Qidiruv |
| `/backup` | Admin | Bazani (`.db`) yuklab olish |
| `/restore` | Admin | `.db` fayl yuborib bazani tiklash |

## Lokal ishga tushirish

```bash
pip install -r requirements.txt
cp env.example .env      # .env ichini to'ldiring (BOT_TOKEN, ADMIN_IDS)
python main.py
```

## Render'ga joylash

1. Ushbu repo'ni GitHub'ga push qiling.
2. Render → **New → Blueprint** → repo'ni tanlang (`render.yaml` avtomatik o'qiladi).
   Yoki **New → Web Service** → `python main.py`.
3. **Environment** bo'limida quyidagilarni kiriting:
   - `BOT_TOKEN` — @BotFather tokeni
   - `ADMIN_IDS` — admin ID (masalan `7903688837`)
4. Deploy.

> ⚠️ `.env` va `school.db` GitHub'ga **chiqmaydi** (`.gitignore`). Sirlar faqat Render dashboard'da.

### 📦 Lokal ma'lumotni Render'ga ko'chirish

Server bo'sh baza bilan ishga tushadi. Hozirgi ma'lumotni o'tkazish:

1. **Lokal** botda (kompyuterda) admin bo'lib `/backup` yuboring → `.db` fayl keladi.
2. Faylni saqlang.
3. **Render'dagi** botga o'sha `.db` faylni oddiy hujjat sifatida yuboring (`/restore`).
4. Bot: *"✅ Baza tiklandi, jami N o'quvchi"* deydi. Tamom.

### 🔒 Ma'lumot RESET bo'lmasligi (muhim)

- **Free plan:** fayl-tizim vaqtinchalik — har redeploy/uyquda baza **RESET bo'ladi**.
  Yechim: vaqti-vaqti bilan `/backup` qilib turing; kerak bo'lsa `/restore`.
- **Doimiy saqlash uchun (tavsiya):** Render **Starter (paid)** plan + **Disk**:
  1. `render.yaml` da `plan: starter`, `DB_PATH: /var/data/school.db`, `disk` blokini oching.
  2. Shunda baza diskda saqlanadi va **hech qachon reset bo'lmaydi**.

## Fayllar

- `main.py` — botning barcha kodi
- `requirements.txt` — kutubxonalar
- `render.yaml` — Render blueprint
- `env.example` — `.env` namunasi
