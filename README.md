# 🤖 Doim ishlaydigan Telegram bot — /hisobot (rasm + Excel)

Dashboard **yopiq bo'lsa ham** ishlaydi. Bulutda doim yoniq turadi.

- Guruhda **`/hisobot`** → **kechagi** kunlik hisobot (🖼️ rasm + 📎 Excel)
- Guruhda **`/hisobot bugun`** → **bugungi** hisobot (rasm + Excel)
- Har kuni **09:00 (Toshkent)** → kechagi hisobot avtomatik `DAILY_CHAT_ID` guruhiga

Ma'lumot Google Sheets'dan olinadi, hisob-kitob dashboard bilan bir xil. Chiqim **filial kesmida** (qaysi filial nimaga) ko'rsatiladi.

---

## 📁 Fayllar
- `bot.mjs` — butun bot (bitta fayl)
- `package.json`, `Dockerfile`, `.gitignore`

---

## 🔑 Kerakli sozlamalar (Environment Variables)

| Nomi | Majburiy | Izoh |
|------|:---:|------|
| `TELEGRAM_BOT_TOKEN` | ✅ | @BotFather bergan token |
| `DAILY_CHAT_ID` | ⬜ | 09:00 avtomatik yuboriladigan guruh id (`-100…`). Bo'sh bo'lsa — faqat /hisobot ishlaydi |
| `DAILY_TIME` | ⬜ | Avtomatik vaqt, Toshkent, standart `09:00` |
| `XARAJAT_SHEET_ID` | ⬜ | Kodda standart bor; boshqa jadval bo'lsa qo'ying |
| `PRIXOD_SHEET_ID` | ⬜ | Kodda standart bor; boshqa jadval bo'lsa qo'ying |

> `/hisobot` **qaysi guruhda yozilsa, o'sha guruhga** javob beradi — u uchun chat_id shart emas. `DAILY_CHAT_ID` faqat avtomatik 09:00 uchun kerak.

---

## 🚂 Variant A — Railway (eng oson)

1. Kodni GitHub repoga yuklang (`bot.mjs`, `package.json`, `Dockerfile`).
2. [railway.app](https://railway.app) → **New Project → Deploy from GitHub repo** → repoingizni tanlang.
3. Railway `Dockerfile` ni o'zi topib quradi.
4. **Variables** bo'limiga yuqoridagi env larni qo'shing (kamida `TELEGRAM_BOT_TOKEN`, `DAILY_CHAT_ID`).
5. Deploy tugagach — bot ishga tushadi. Guruhda `/hisobot` deb sinang.

> Railway'da bepul kredit bor; tugasa arzon "Hobby" reja yetadi (bot juda kam resurs ishlatadi).

---

## 🎨 Variant B — Render

1. Kodni GitHub repoga yuklang.
2. [render.com](https://render.com) → **New → Web Service** → repo tanlang.
3. **Runtime: Docker** (Dockerfile o'zi topiladi). Instance type: **Free** yoki **Starter**.
4. **Environment** → env larni qo'shing.
5. **Create Web Service**. Loglarda `🤖 Bot ishga tushdi` chiqadi.

> Render Free reja faolsizlikda "uxlab" qolishi mumkin — health server bor, lekin doim uzluksiz ishlashi uchun **Starter** (arzon) tavsiya etiladi. Railway bu jihatdan qulayroq.

---

## 🧪 Lokal sinash (ixtiyoriy)
```bash
npm install
TELEGRAM_BOT_TOKEN=xx: node bot.mjs
# guruhda /hisobot deb yozing
```

---

## ⚠️ Muhim

1. **@BotFather → /setprivacy → Disable** qiling (yoki botni guruhga admin qiling). Aks holda bot guruhdagi `/hisobot` ni ko'rmaydi. (`/hisobot@BotNomi` ko'rinishi privacy yoniq bo'lsa ham ishlaydi.)
2. **Faqat bitta joy** `getUpdates` ni o'qisin. Shu bot ishlab tursa:
   - dashboarddagi "🤖 /hisobot tinglash" tugmasini **yoqmang**,
   - avvalgi GitHub Actions cron'ini **o'chiring** (bu botning o'zida 09:00 bor).
   Ikki joy birga o'qisa — "Conflict" xatosi chiqadi.
3. Ikkala Google Sheet **"havolaga ega har kim ko'ra oladi"** bo'lsin.
4. Token rasmda/ochiq joyda ko'ringan bo'lsa — @BotFather `/revoke` bilan yangilang.
