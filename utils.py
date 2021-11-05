from openpyxl import load_workbook
from telegram import InlineKeyboardButton, InlineKeyboardMarkup, ParseMode
import datetime

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
        if sheet.title not in ["Schedule", "Groups"]:
            keyboard.append([InlineKeyboardButton(str(sheet.title), callback_data=f"sheet_{sheet.title}")])
    reply_markup = InlineKeyboardMarkup(keyboard)
    context.bot.send_message(
        chat_id=update.effective_chat.id, 
        text="These are your sheets:",
        reply_markup=reply_markup,
    )
    return

def show_groups(wb, query, context):
    sheet = wb["Groups"]
    group_list = [item.value for item in sheet["A"]]
    text = "Groups Selected:" + " "
    selected_groups = context.user_data.get("groups")
    if not len(selected_groups):
        text += "**None**"
    else:
        text += ", ".join([group for group in selected_groups])
    print(group_list)
    groups = set(group_list).symmetric_difference(set(selected_groups))
    keyboard = []
    for group in groups:
        keyboard.append([InlineKeyboardButton(group, callback_data=f"group_{group}")])
    if len(selected_groups):
        keyboard.append([InlineKeyboardButton("Done.", callback_data=f"done")])
    reply_markup = InlineKeyboardMarkup(keyboard)
    query.edit_message_text(
        text=text, 
        reply_markup=reply_markup,
        parse_mode=ParseMode.MARKDOWN
    )

def validate_date(update, context):
    text = update.message.text
    date = [int(x) for x in text.split("-")]
    # An error might occur while running questions if the date has been set in the past
    try:
        datetime.datetime(date[0], date[1], date[2])
    except:
        update.message.reply_text("Wrong format, use YYYY-MM-DD instead.")
        return None
    context.bot.send_message(
        chat_id=update.effective_chat.id,
        text="Write a time (in HH:MM:SS format):"
    )
    return text

def validate_time(update, context):
    text = update.message.text
    timeformat = "%H:%M:%S"
    try:
        # Check if 0 problem happens here!
        datetime.datetime.strptime(text, timeformat)
    except:
        update.message.reply_text("Wrong format, use HH:MM:SS instead.")
        return None
    return text

def fill_database(wb, context):
    sheet = wb["Schedule"]
    data = [
        context.user_data["sheet"],
        ", ".join(context.user_data["groups"]),
        context.user_data["date"],
        context.user_data["time"]
    ]
    sheet.append(data)
    wb.save(filename="custom/excel_sheet.xlsx")
    return wb

def save_data(update, context):
    # Update workbook
    wb = open_workbook(update, context)
    if not wb:
        return
    if fill_database(wb, context):
        context.bot.send_message(chat_id=update.effective_chat.id, text="Data updated!")
    else:
        context.bot.send_message("Oops, something went wrong, try again!")