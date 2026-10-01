from aiogram.types import InlineKeyboardMarkup, InlineKeyboardButton

def get_schedule_keyboard(current_view: str = "today") -> InlineKeyboardMarkup:
    """
    Jadval ko'rinishlari o'rtasida o'tish uchun qulay va oddiy tugmalar.
    """
    buttons = [
        [
            InlineKeyboardButton(text="📅 Bugun", callback_data="view:today"),
            InlineKeyboardButton(text="➡️ Ertaga", callback_data="view:tomorrow"),
            InlineKeyboardButton(text="📋 Hafta", callback_data="view:week")
        ]
    ]
    return InlineKeyboardMarkup(inline_keyboard=buttons)
