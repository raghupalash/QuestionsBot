from telegram import Update, Message, ParseMode, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    Updater, 
    CommandHandler, 
    ConversationHandler,
    CallbackQueryHandler, 
    CallbackContext,
)
from openpyxl import load_workbook

ADMIN_ID = "1107423707"

def start(update: Update, context: CallbackContext):
    # Allow only admin
    if update.message.from_user.id != ADMIN_ID:
        context.bot.send_message(chat_id=update.effective_chat.id, text="Sorry, I'm not allowed to talk to you!")
        return
    # Load worksheet
    try:
        wb = load_workbook(f"custom/file{update.effective_chat.id}.xlsx")
    except:
        context.bot.send_message(chat_id=update.effective_chat.id, text="Send a sheet with links first!")
        return
    show_sheets(wb, update, context)
    
def show_sheets(wb, update, context):
    keyboard = []
    for sheet in wb.worksheets:
        keyboard.append(InlineKeyboardButton(sheet.title, callback_data=f"sheet_{sheet.title}"))
    reply_markup = InlineKeyboardMarkup(keyboard)
    context.bot.send_message(
        chat_id=update.effective_chat.id, 
        text="These are your sheets:",
        reply_markup=reply_markup,
    )
    return

def show_groups():
    groups = ["group a", "group b", "group c"]
    keyboard = []
    for i, group in enumerate(groups):
        keyboard.append(InlineKeyboardButton(f"{i}. {group}", callback_data=f"group_{group}"))
    context.bot.send_message(chat_id=udpate.effective_chat..id, )

def groups(update: Update, context: CallbackContext):

    

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

def main():
    updater = Updater(token="1982910780:AAGdFTAqcud7pC1p7zZdcVtxh-UgO177-xM")

    dispatcher = updater.dispatcher

    start_handler = CommandHandler('start', start)
    dispatcher.add_handler(MessageHandler(Filters.document, incoming_document))

    dispatcher.add_handler(start_handler)
    dispatcher.add_handler(button_handler)

    updater.start_polling()

    updater.idle()

if __name__ == "__main__":
    main()