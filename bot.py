# We are going to use this file for main bot logic and holding data.
# For sending messages and storing data, we will use utils.py
from telegram import Update, Message, ParseMode, ReplyKeyboardMarkup
from telegram.ext import (
    Updater, 
    CommandHandler, 
    ConversationHandler,
    MessageHandler,
    CallbackQueryHandler, 
    CallbackContext,
    Filters,
)
import logging
from openpyxl import load_workbook
from utils import show_sheets, show_groups, open_workbook, validate_date, validate_time, fill_database, save_data

# Enable logging
logging.basicConfig(
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s', level=logging.DEBUG
)

ADMIN_ID = 1107423707
  
ONE, TWO, THREE = range(3)
def incoming_document(update, context):
    file_type = update.message.document.mime_type
    if file_type == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet":
        # writing to a custom file
        with open(f"custom/excel_sheet.xlsx", 'wb') as f:
            context.bot.get_file(update.message.document).download(out=f)
        
        keyboard = [["/start"]]
        reply_markup = ReplyKeyboardMarkup(
            keyboard,
            one_time_keyboard=False,
            resize_keyboard=True,
            input_field_placeholder="start"
        )
        context.bot.send_message(chat_id=update.effective_chat.id, 
            text="Sheet uploaded, press /start to update the sheet.", reply_markup=reply_markup)
    else:
        context.bot.send_message(chat_id=update.effective_chat.id,
            text="Only ms office spreadsheet allowed with extension '.xlsx'")

def start(update: Update, context:CallbackContext):
    # Allow only admin
    if update.message.from_user.id != ADMIN_ID:
        context.bot.send_message(chat_id=update.effective_chat.id, text="Sorry, I'm not allowed to talk to you!")
        return ConversationHandler.END
    # Load worksheet
    wb = open_workbook(update, context)
    if not wb:
        return ConversationHandler.END
    show_sheets(wb, update, context)
    return ONE

def groups(update: Update, context: CallbackContext):
    query = update.callback_query
    query.answer()

    data = query.data.split("_")
    if data[0] == "sheet":
        context.user_data["sheet"] = data[1]
        context.user_data["groups"] = []
    elif data[0] == "group":
        # Avoid repeatition while adding into list
        if data[1] not in context.user_data.get("groups"):
            context.user_data["groups"].append(data[1])
    else:
        # User pressed "done"
        query.edit_message_text(text="Write the date (in YYYY-MM-DD format):") # Not a good design
        return TWO

    # wb = open_workbook(update, context)
    # if not wb:
    #     return ConversationHandler.END
    show_groups(query, context)
    
def schedule_date(update: Update, context: CallbackContext):
    date = validate_date(update, context)
    if not date:
        return TWO
    context.user_data["date"] = date
    return THREE

def schedule_time(update: Update, context: CallbackContext):
    time = validate_time(update, context)
    print(time)
    if not time:
        return THREE
    context.user_data["time"] = time
    save_data(update, context)
    return ConversationHandler.END

def add_group(update: Update, context: CallbackContext):
    if update.message.from_user.id != ADMIN_ID:
        update.message.reply_text("Sorry, only the bot owner can add the bot in the database.")
        return
    # Save group info
    group_title = update.message.chat.title
    group_id = update.message.chat.id
    wb = open_workbook(update, context)
    if not wb:
        return
    sheet = wb["Groups"]
    for item in sheet['B']:
        if item.value == group_id:
            update.message.reply_text("The group is already in the database")
            return
    data = [group_title, group_id]
    sheet.append(data)
    wb.save(filename="custom/excel_sheet.xlsx")
    update.message.reply_text("Group added to the database.")

def cancel():
    ConversationHandler.END

def main():
    updater = Updater(token="2042937645:AAFnQDvY7UrVNxhW8J0eRsC7ZTctWQ8M6Ds")

    dispatcher = updater.dispatcher

    start_handler = ConversationHandler(
        entry_points=[CommandHandler("start", start)],
        states={
            ONE: [CallbackQueryHandler(groups)],
            TWO: [MessageHandler(Filters.text, schedule_date)],
            THREE: [MessageHandler(Filters.text, schedule_time)],
        },
        fallbacks={CommandHandler("cancel", cancel)}
    )

    dispatcher.add_handler(MessageHandler(Filters.document, incoming_document))
    dispatcher.add_handler(CommandHandler('add', add_group))
    dispatcher.add_handler(start_handler)

    updater.start_polling()

    updater.idle()

if __name__ == "__main__":
    main()