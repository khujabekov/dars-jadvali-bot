import os
import sys
import shutil
import logging
import asyncio
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
from pathlib import Path

from aiogram import Bot, Dispatcher, F
from aiogram.types import Message, FSInputFile, ContentType, ChatMemberUpdated, CallbackQuery
from aiogram.filters import Command, CommandStart, ChatMemberUpdatedFilter, IS_MEMBER, IS_NOT_MEMBER, ADMINISTRATOR
from aiogram.enums import ParseMode
from aiogram.client.default import DefaultBotProperties

import config
from excel_parser import format_day_schedule, format_week_schedule, parse_schedule_excel, convert_xls_to_xlsx
from scheduler import setup_scheduler, send_daily_schedule
from sample_generator import create_sample_excel
from keyboards import get_schedule_keyboard

# Windows konsoli uchun UTF-8 kodlashni yoqish
if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

# Logging sozlamalari
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - [%(levelname)s] - %(name)s - %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)]
)
logger = logging.getLogger("DarsJadvaliBot")

# Bot va Dispatcher yaratish
dp = Dispatcher()

# Barcha xabarlarni log qilish
@dp.message.outer_middleware()
async def log_all_messages(handler, event: Message, data):
    chat_title = getattr(event.chat, 'title', None)
    logger.info(f"Yangi xabar: chat_id={event.chat.id}, type={event.chat.type}, title={chat_title!r}, user_id={event.from_user.id}, text={event.text!r}")
    return await handler(event, data)

# Bot guruhga qo'shilganda yoki admin qilinganda avtomatik guruhni aniqlash
@dp.my_chat_member()
async def on_my_chat_member(event: ChatMemberUpdated, bot: Bot):
    status = event.new_chat_member.status
    logger.info(f"Bot a'zoligi yangilandi: chat_id={event.chat.id}, chat_title={event.chat.title!r}, status={status}")
    if status in ("administrator", "member"):
        config.GROUP_ID = event.chat.id
        try:
            env_path = config.BASE_DIR / ".env"
            if env_path.exists():
                import re
                content = env_path.read_text(encoding="utf-8")
                if re.search(r"^GROUP_ID=.*$", content, flags=re.MULTILINE):
                    new_content = re.sub(r"^GROUP_ID=.*$", f"GROUP_ID={event.chat.id}", content, flags=re.MULTILINE)
                else:
                    new_content = content + f"\nGROUP_ID={event.chat.id}\n"
                env_path.write_text(new_content, encoding="utf-8")
                logger.info(f"Avtomatik yangilandi: GROUP_ID={event.chat.id}")
        except Exception as e:
            logger.error(f".env faylni yangilashda xatolik: {e}")

        try:
            await bot.send_message(
                chat_id=event.chat.id,
                text=(
                    f"✅ <b>Bot guruhga ulandi!</b>\n\n"
                    f"Guruh: <b>{event.chat.title}</b>\n"
                    f"Jadval har kuni soat {config.SEND_HOUR:02d}:{config.SEND_MINUTE:02d} da yuboriladi."
                ),
                parse_mode=ParseMode.HTML
            )
        except Exception as e:
            logger.error(f"Xabar yuborishda xatolik: {e}")

def is_admin(user_id: int) -> bool:
    """Foydalanuvchi admin ekanligini tekshiradi."""
    if not config.ADMIN_IDS:
        return True
    return user_id in config.ADMIN_IDS

@dp.message(CommandStart())
async def cmd_start(message: Message):
    admin_note = ""
    if is_admin(message.from_user.id):
        admin_note = (
            "\n\n<b>Admin buyruqlari:</b>\n"
            "• Excel (.xlsx) yuborish — jadvalni yangilash\n"
            "• /send_now — Jadvalni hozir guruhga yuborish"
        )

    text = (
        f"Assalomu alaykum, <b>{message.from_user.full_name}</b>!\n\n"
        "<b>Buyruqlar:</b>\n"
        "• /bugun — Bugungi dars jadvali\n"
        "• /ertaga — Ertangi dars jadvali\n"
        "• /hafta — Haftalik dars jadvali\n"
        "• /namuna — Excel shablonini olish\n"
        "• /id — Chat va foydalanuvchi ID si"
        f"{admin_note}"
    )
    await message.answer(text, parse_mode=ParseMode.HTML)

@dp.message(Command("id"))
async def cmd_id(message: Message):
    text = (
        f"👤 <b>Sizning ID:</b> <code>{message.from_user.id}</code>\n"
        f"💬 <b>Chat ID:</b> <code>{message.chat.id}</code>"
    )
    await message.answer(text, parse_mode=ParseMode.HTML)

@dp.message(Command("namuna", "template"))
async def cmd_namuna(message: Message):
    if not config.TEMPLATE_FILE.exists():
        create_sample_excel(config.TEMPLATE_FILE)
        
    doc = FSInputFile(config.TEMPLATE_FILE, filename="jadval_namuna.xlsx")
    caption = "📥 Dars jadvali Excel namunasi.\nTo'ldirib, botga yuborishingiz mumkin."
    await message.answer_document(doc, caption=caption, parse_mode=ParseMode.HTML)

@dp.message(Command("bugun", "today"))
async def cmd_bugun(message: Message):
    try:
        tz = ZoneInfo(config.TIMEZONE)
    except Exception:
        tz = ZoneInfo("Asia/Tashkent")
        
    now = datetime.now(tz)
    weekday = now.weekday()
    text = format_day_schedule(weekday, now.date())
    kb = None if message.chat.type in ("group", "supergroup") else get_schedule_keyboard("today")
    await message.answer(text, reply_markup=kb, parse_mode=ParseMode.HTML)

@dp.message(Command("ertaga", "tomorrow"))
async def cmd_ertaga(message: Message):
    try:
        tz = ZoneInfo(config.TIMEZONE)
    except Exception:
        tz = ZoneInfo("Asia/Tashkent")
        
    now = datetime.now(tz) + timedelta(days=1)
    weekday = now.weekday()
    text = format_day_schedule(weekday, now.date())
    kb = None if message.chat.type in ("group", "supergroup") else get_schedule_keyboard("tomorrow")
    await message.answer(text, reply_markup=kb, parse_mode=ParseMode.HTML)

@dp.message(Command("hafta", "week"))
async def cmd_hafta(message: Message):
    text = format_week_schedule()
    kb = None if message.chat.type in ("group", "supergroup") else get_schedule_keyboard("week")
    await message.answer(text, reply_markup=kb, parse_mode=ParseMode.HTML)

@dp.callback_query(F.data.startswith("view:"))
async def on_view_schedule_callback(callback: CallbackQuery):
    view_type = callback.data.split(":")[1]
    try:
        tz = ZoneInfo(config.TIMEZONE)
    except Exception:
        tz = ZoneInfo("Asia/Tashkent")
        
    now = datetime.now(tz)
    
    if view_type == "today":
        weekday = now.weekday()
        text = format_day_schedule(weekday, now.date())
        kb = get_schedule_keyboard("today")
    elif view_type == "tomorrow":
        tomorrow = now + timedelta(days=1)
        weekday = tomorrow.weekday()
        text = format_day_schedule(weekday, tomorrow.date())
        kb = get_schedule_keyboard("tomorrow")
    elif view_type == "week":
        text = format_week_schedule()
        kb = get_schedule_keyboard("week")
    else:
        await callback.answer()
        return

    try:
        await callback.message.edit_text(text, reply_markup=kb, parse_mode=ParseMode.HTML)
    except Exception:
        pass
    await callback.answer()

@dp.message(Command("setgroup", "set_group"))
async def cmd_set_group(message: Message):
    if message.chat.type not in ("group", "supergroup"):
        await message.answer("⚠️ Bu buyruqni dars jadvali yuborilishi kerak bo'lgan Telegram guruh ichida yozishingiz kerak!")
        return

    if not is_admin(message.from_user.id):
        await message.answer("❌ Bu buyruq faqat bot adminlari uchun!", parse_mode=ParseMode.HTML)
        return

    group_id = message.chat.id
    config.GROUP_ID = group_id

    # .env faylini avtomatik yangilash
    try:
        env_path = config.BASE_DIR / ".env"
        if env_path.exists():
            import re
            content = env_path.read_text(encoding="utf-8")
            if re.search(r"^GROUP_ID=.*$", content, flags=re.MULTILINE):
                new_content = re.sub(r"^GROUP_ID=.*$", f"GROUP_ID={group_id}", content, flags=re.MULTILINE)
            else:
                new_content = content + f"\nGROUP_ID={group_id}\n"
            env_path.write_text(new_content, encoding="utf-8")
    except Exception as e:
        logger.error(f".env faylni yangilashda xatolik: {e}")

    await message.answer(
        f"✅ Dars jadvali guruhi sozlandi: <b>{message.chat.title}</b>\n"
        f"Jadval har kuni soat {config.SEND_HOUR:02d}:{config.SEND_MINUTE:02d} da yuboriladi.",
        parse_mode=ParseMode.HTML
    )

@dp.message(Command("send_now"))
async def cmd_send_now(message: Message, bot: Bot):
    if not is_admin(message.from_user.id):
        await message.answer("❌ Bu buyruq faqat adminlar uchun.")
        return
        
    await message.answer("⏳ Jadval yuborilmoqda...")
    target_id = message.chat.id if message.chat.type in ("group", "supergroup") else config.GROUP_ID
    success, res_msg = await send_daily_schedule(bot, target_chat_id=target_id, force_send=True)
    if success:
        await message.answer(f"✅ {res_msg}")
    else:
        await message.answer(f"❌ {res_msg}")

@dp.message(F.new_chat_members)
async def on_new_chat_members(message: Message, bot: Bot):
    bot_info = await bot.get_me()
    for member in message.new_chat_members:
        if member.id == bot_info.id:
            await message.answer(
                f"Assalomu alaykum!\n"
                f"Guruh ID: <code>{message.chat.id}</code>\n"
                f"Jadvalni sozlash uchun guruhda /setgroup buyrug'ini yuboring.",
                parse_mode=ParseMode.HTML
            )
            break

@dp.message(F.document)
async def handle_excel_upload(message: Message, bot: Bot):
    doc = message.document
    file_name = doc.file_name or ""
    lower_name = file_name.lower()
    
    if not (lower_name.endswith(".xlsx") or lower_name.endswith(".xls")):
        if message.chat.type == "private":
            await message.answer("⚠️ Iltimos, dars jadvali uchun <b>.xlsx</b> yoki <b>.xls</b> (Excel) formatidagi fayl yuboring.", parse_mode=ParseMode.HTML)
        return

    # Adminlik tekshiruvi
    if not is_admin(message.from_user.id):
        await message.answer("🚫 Kechirasiz, faqat bot adminlari dars jadvalini yangilashi mumkin.", parse_mode=ParseMode.HTML)
        return

    msg = await message.answer("⏳ Fayl yuklab olinmoqda va tekshirilmoqda...")
    
    ext = ".xls" if lower_name.endswith(".xls") else ".xlsx"
    temp_path = config.DATA_DIR / f"temp_{doc.file_id}{ext}"
    try:
        # Faylni yuklab olish
        await bot.download(doc, destination=temp_path)
        
        # Tekshirib ko'rish
        parsed, error = parse_schedule_excel(temp_path)
        if error:
            await msg.edit_text(f"❌ <b>Faylda xatolik aniqlandi:</b>\n{error}\n\nJadval yangilanmadi.", parse_mode=ParseMode.HTML)
            return
            
        total_lessons = sum(len(lessons) for lessons in parsed.values())
        if total_lessons == 0:
            await msg.edit_text("⚠️ Faylda birorta ham dars topilmadi. Ustun nomlari va ma'lumotlarni tekshiring.", parse_mode=ParseMode.HTML)
            return
            
        # Asosiy faylga saqlash
        if ext == ".xls":
            convert_xls_to_xlsx(temp_path, config.SCHEDULE_FILE)
        else:
            shutil.move(temp_path, config.SCHEDULE_FILE)
        
        # Hisobot tayyorlash
        summary_lines = []
        from excel_parser import UZBEK_DAYS
        for day_num, lessons in parsed.items():
            if lessons:
                summary_lines.append(f"• <b>{UZBEK_DAYS[day_num]}:</b> {len(lessons)} ta dars")
                
        summary_text = "\n".join(summary_lines)
        
        await msg.edit_text(
            f"✅ <b>Dars jadvali saqlandi!</b>\n\n"
            f"Jami: {total_lessons} ta dars\n"
            f"{summary_text}",
            parse_mode=ParseMode.HTML
        )
        logger.info(f"Yangi dars jadvali yuklandi. Yuklagan: {message.from_user.id} ({message.from_user.full_name})")
        
    except Exception as e:
        logger.exception("Faylni qayta ishlashda xatolik:")
        await msg.edit_text(f"❌ Faylni saqlashda kutilmagan xatolik yuz berdi:\n<code>{e}</code>", parse_mode=ParseMode.HTML)
    finally:
        if temp_path.exists():
            temp_path.unlink()

async def main():
    if not config.BOT_TOKEN:
        print("\n" + "!" * 60)
        print("DIQQAT: .env faylida BOT_TOKEN ko'rsatilmagan!")
        print("1. .env faylini oching.")
        print("2. BOT_TOKEN=... qatoriga @BotFather dan olingan tokenni yozing.")
        print("3. Qayta ishga tushiring: python main.py")
        print("!" * 60 + "\n")
        return

    # Namunaviy fayllarni tayyorlab qo'yish
    if not config.TEMPLATE_FILE.exists():
        create_sample_excel(config.TEMPLATE_FILE)
    if not config.SCHEDULE_FILE.exists():
        shutil.copy(config.TEMPLATE_FILE, config.SCHEDULE_FILE)

    bot = Bot(
        token=config.BOT_TOKEN,
        default=DefaultBotProperties(parse_mode=ParseMode.HTML)
    )

    # Scheduler ni ishga tushirish
    scheduler = setup_scheduler(bot)
    scheduler.start()

    bot_info = await bot.get_me()
    print(f"\n[OK] Bot ishga tushdi: @{bot_info.username}")
    print(f"[OK] Dars jadvali har kuni soat {config.SEND_HOUR:02d}:{config.SEND_MINUTE:02d} da guruhga yuboriladi.")
    print("To'xtatish uchun: Ctrl + C\n")

    try:
        await dp.start_polling(
            bot,
            allowed_updates=["message", "my_chat_member", "chat_member", "callback_query"]
        )
    finally:
        scheduler.shutdown()
        await bot.session.close()

if __name__ == "__main__":
    try:
        asyncio.run(main())
    except (KeyboardInterrupt, SystemExit):
        logger.info("Bot to'xtatildi.")
