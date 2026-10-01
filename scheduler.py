import logging
from datetime import datetime
from zoneinfo import ZoneInfo
from aiogram import Bot
from apscheduler.schedulers.asyncio import AsyncIOScheduler
from apscheduler.triggers.cron import CronTrigger

import config
from excel_parser import format_day_schedule, parse_schedule_excel

logger = logging.getLogger(__name__)

async def send_daily_schedule(bot: Bot, target_chat_id: int | str = None, force_send: bool = False):
    """
    Guruhga kunlik dars jadvalini yuboradi.
    target_chat_id ko'rsatilmasa, config.GROUP_ID ishlatiladi.
    force_send=True bo'lsa, dam olish kunida ham xabar jo'natadi.
    """
    chat_id = target_chat_id or config.GROUP_ID
    if not chat_id:
        logger.warning("GROUP_ID belgilanmagan! Xabar yuborilmadi. .env faylni tekshiring.")
        return False, "GROUP_ID ko'rsatilmagan. Iltimos, .env faylida GROUP_ID ni sozlang."

    try:
        tz = ZoneInfo(config.TIMEZONE)
    except Exception:
        tz = ZoneInfo("Asia/Tashkent")
        
    now = datetime.now(tz)
    weekday = now.weekday() # 0 = Dushanba, ..., 6 = Yakshanba
    today_date = now.date()
    
    # Jadvalni tekshirish
    schedule, error = parse_schedule_excel(config.SCHEDULE_FILE)
    if error:
        logger.error(f"Dars jadvalini o'qishda xatolik: {error}")
        if force_send:
            return False, f"Xatolik: {error}"
        return False, error
        
    lessons = schedule.get(weekday, [])
    
    # Agar yakshanba yoki dars bo'lmasa va dam olish kunlarida yuborish o'chirilgan bo'lsa:
    if not lessons and not config.NOTIFY_ON_OFF_DAYS and not force_send:
        logger.info(f"Bugun ({today_date}, {weekday}-kun) dars yo'q va NOTIFY_ON_OFF_DAYS=false. Xabar yuborilmadi.")
        return True, "Bugun dars yo'q, xabar yuborilmadi."

    message_text = format_day_schedule(weekday, today_date, schedule=schedule)
    
    try:
        await bot.send_message(
            chat_id=chat_id,
            text=message_text,
            parse_mode="HTML"
        )
        logger.info(f"Kunlik dars jadvali guruhga ({chat_id}) muvaffaqiyatli yuborildi.")
        return True, "Jadval guruhga yuborildi!"
    except Exception as e:
        logger.error(f"Guruhga xabar yuborishda xatolik yuz berdi: {e}")
        err_str = str(e)
        if "chat not found" in err_str:
            return False, "Guruh topilmadi. Botni guruhga qo'shib, guruh ichida /setgroup buyrug'ini yuboring."
        elif "have no rights to send a message" in err_str or "restricted" in err_str or "not enough rights" in err_str:
            return False, "Botga guruhda xabar yozish ruxsati berilmagan."
        elif "bot was kicked" in err_str:
            return False, "Bot guruhdan chiqarilgan. Qayta qo'shing."
        return False, f"Xatolik: {e}"


def setup_scheduler(bot: Bot) -> AsyncIOScheduler:
    """APScheduler ni sozlaydi va ishga tushiradi."""
    scheduler = AsyncIOScheduler()
    
    try:
        tz = ZoneInfo(config.TIMEZONE)
    except Exception:
        tz = ZoneInfo("Asia/Tashkent")
        
    trigger = CronTrigger(
        hour=config.SEND_HOUR,
        minute=config.SEND_MINUTE,
        timezone=tz
    )
    
    scheduler.add_job(
        send_daily_schedule,
        trigger=trigger,
        kwargs={"bot": bot, "force_send": False},
        id="daily_schedule_job",
        replace_existing=True
    )
    
    logger.info(
        f"Scheduler faollashtirildi: Har kuni soat {config.SEND_HOUR:02d}:{config.SEND_MINUTE:02d} da ({config.TIMEZONE}) yuboriladi."
    )
    return scheduler
