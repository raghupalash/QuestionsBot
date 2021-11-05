from openpyxl import load_workbook
from telegram import InlineKeyboardButton, InlineKeyboardMarkup, ParseMode

def open_workbook(update, context):
    try:
        wb = load_workbook(f"custom/excel_sheet.xlsx")
    except:
        context.bot.send_message(chat_id=update.effective_chat.id, text="Send a sheet first!")
        return None
    return wb

def show_sheets(wb, update, context):
    keyboard = []
    for sheet in wb.worksheets:
        keyboard.append([InlineKeyboardButton(str(sheet.title), callback_data=f"sheet_{sheet.title}")])
    reply_markup = InlineKeyboardMarkup(keyboard)
    context.bot.send_message(
        chat_id=update.effective_chat.id, 
        text="These are your sheets:",
        reply_markup=reply_markup,
    )
    return

def show_groups(query, context):
    text = "Groups Selected:" + " "
    group_list = context.user_data.get("groups")
    if not len(group_list):
        text += "**None**"
    else:
        text += ", ".join([group for group in group_list])
    groups = set(["group a", "group b", "group c"]).symmetric_difference(set(group_list))
    keyboard = []
    for group in groups:
        keyboard.append([InlineKeyboardButton(group, callback_data=f"group_{group}")])
    if len(group_list):
        keyboard.append([InlineKeyboardButton("Done.", callback_data=f"done")])
    reply_markup = InlineKeyboardMarkup(keyboard)
    query.edit_message_text(
        text=text, 
        reply_markup=reply_markup,
        parse_mode=ParseMode.MARKDOWN
    )
