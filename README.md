# 📅 Telegram Dars Jadvali Boti

Telegram guruhga har kuni belgilangan vaqtda (masalan, soat 07:30 da) kunlik dars jadvalini avtomatik yuboruvchi Python boti.

Dars jadvalini o'zgartirish juda oson — shunchaki to'ldirilgan **Excel (`.xlsx`)** faylini botga yuborasiz va bot uni avtomatik saqlab oladi!

---

## ✨ Asosiy imkoniyatlari

1. ⏰ **Kunlik avtomatik xabar:** Har kuni ertalab (soat 07:30 da) bugungi darslar ro'yxatini guruhga yuboradi.
2. 📊 **Excel orqali boshqarish:** Dars jadvalini Excel (`.xlsx`) orqali kiritish va yangilash.
3. 📥 **Tayyor namuna:** `/namuna` buyrug'i orqali chiroyli dizayndagi Excel shablonini yuklab olish.
4. 🗓 **Foydali buyruqlar:**
   - `/bugun` — Bugungi darslar ro'yxati
   - `/ertaga` — Ertangi kun darslari ro'yxati
   - `/hafta` — Butun haftalik to'liq jadval
   - `/send_now` — Guruhga jadvalni zudlik bilan yuborish (faqat adminlar uchun)
   - `/id` — Chat va foydalanuvchi ID sini aniqlash
5. 🛡 **Xavfsiz boshqaruv:** Jadvalni faqat ruxsat berilgan adminlar yangilay oladi.

---

## 🚀 O'rnatish va Ishga tushirish

### 1. Bog'liqliklarni o'rnatish

Terminalda quyidagi buyruqni bajaring:
```bash
pip install -r requirements.txt
```

### 2. Sozlamalarni kiritish (.env)

Loyihadagi `.env` faylini oching va quyidagi parametrlarni to'ldiring:

```ini
# 1. @BotFather dan olingan bot tokeni:
BOT_TOKEN=1234567890:ABCdefGHIjklMNOpqrSTUvwxyz

# 2. Dars jadvali yuboriladigan guruh ID raqami (masalan: -1001234567890)
GROUP_ID=-1001234567890

# 3. O'zingizning Telegram ID raqamingiz (vergul bilan bir nechta yozish mumkin):
ADMIN_IDS=123456789

# 4. Jadval har kuni soat nechada yuborilsin (Format: HH:MM):
SEND_TIME=07:30

# 5. Vaqt mintaqasi:
TIMEZONE=Asia/Tashkent

# 6. Dam olish yoki dars yo'q kunlarda ham "Bugun dars yo'q" deb xabar borsinmi?
NOTIFY_ON_OFF_DAYS=false
```

---

## 🔍 Guruh va Admin ID sini qanday bilish mumkin?

1. Botni ishga tushiring: `python main.py`
2. Botga shaxsiy xabarda `/id` deb yozing — bot sizning **User ID** ingizni ko'rsatadi (buni `ADMIN_IDS` ga qo'ying).
3. Botni dars jadvali yuborilishi kerak bo'lgan Telegram guruhga qo'shing va guruhda `/id` deb yozing — bot guruhning manfiy ID sini (masalan: `-1002345678901`) ko'rsatadi (buni `GROUP_ID` ga qo'ying).
4. Bot guruhga xabar yuborishi uchun unga guruhda xabar yozish huquqi berilgan bo'lishi kerak.

---

## 📝 Excel jadvalini qanday kiritish mumkin?

1. Botga shaxsiy xabarda `/namuna` buyrug'ini yuboring.
2. Bot sizga `jadval_namuna.xlsx` faylini yuboradi.
3. Faylni ochib, o'z guruhingiz dars jadvalini kiriting:
   - **Kun:** Dushanba, Seshanba, Chorshanba, Payshanba, Juma, Shanba
   - **Vaqt:** 08:30 - 09:50
   - **Fan:** Matematika
   - **O'qituvchi:** Dots. Karimov A.
   - **Xona:** 305-xona
   - **Dars turi:** Ma'ruza / Amaliyot
4. Faylni saqlab, uni botga to'g'ridan-to'g'ri yuboring.
5. Bot faylni tekshiradi va jadvalni saqlaydi!

---

## 🏃‍♂️ Botni ishga tushirish

```bash
python main.py
```
Bot fon rejimida ishlaydi va har kuni belgilangan vaqtda jadvalni avtomatik guruhga joylaydi.
