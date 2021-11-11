# We are going to use this file for main bot logic and holding data.
# For sending messages and storing data, we will use utils.py
from telegram import Update, Message, ParseMode, ReplyKeyboardMarkup, ParseMode, message
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
from telegram.inline.inlinekeyboardbutton import InlineKeyboardButton
from telegram.inline.inlinekeyboardmarkup import InlineKeyboardMarkup
from utils import (
    check_admin,
    show_sheets, 
    show_groups, 
    open_workbook, 
    validate_date, 
    validate_time, 
    save_data, 
    create_datetime,
    group_ids_by_title,
    send_message_to_ids,
    in_run_time,
    collect_garbage,
    send_report,
    get_question_number
)
from credentials import admin_id, token
import time
import datetime
import pytz

# Enable logging
logging.basicConfig(
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s', level=logging.INFO
)

ADMIN_ID = admin_id
TOKEN = token
  
ONE, TWO, THREE = range(3)


def add_sheet_options(update, context):
    if not check_admin(update, context):
        ConversationHandler.END
    # in open workbook error message, say "use /addsheet to add sheet"
    keyboard = [
        [InlineKeyboardButton("questions sheet", callback_data="question_sheet")],
        [InlineKeyboardButton("attendance sheet", callback_data="attendance_sheet")],
    ]
    reply_markup = InlineKeyboardMarkup(keyboard)
    message = "Which sheet do you want to add?"
    update.message.reply_text(text=message, reply_markup=reply_markup)
    return ONE

def ask_for_sheet(update, context):
    query = update.callback_query
    query.answer()
    context.user_data["sheet_type"] = query.data
    message = "Alright, send me the sheet:"
    query.edit_message_text(text=message)
    return TWO

def incoming_document(update, context):
    document = update.message.document
    mime_type = document.mime_type
    if mime_type == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet":
        # writing to a custom file
        if context.user_data["sheet_type"] == "question_sheet":
            file_location = "custom/excel_sheet.xlsx"
        else:
            file_location = f"custom/attendance_sheet.xlsx"
        with open(file_location, 'wb') as f:
            context.bot.get_file(update.message.document).download(out=f)
        
        keyboard = [["/start"]]
        reply_markup = ReplyKeyboardMarkup(
            keyboard,
            one_time_keyboard=False,
            resize_keyboard=True,
            input_field_placeholder="start"
        )
        context.bot.send_message(chat_id=update.effective_chat.id, 
            text="Sheet uploaded, press /start to view sheets from.", reply_markup=reply_markup)
    else:
        context.bot.send_message(chat_id=update.effective_chat.id,
            text="Only ms office spreadsheet allowed with extension '.xlsx'")
    ConversationHandler.END

def start(update: Update, context:CallbackContext):
    # Allow only pms
    if update.message.chat.type == "group":
        return
    # Allow only admin
    if not check_admin(update, context):
        return ConversationHandler.END
    # Check if there's an attendance sheet at all.
    try:
        load_workbook("custom/attendance_sheet.xlsx")
    except:
        message = "Upload attendance spreadsheet file using /addsheet!"
        context.bot.send_message(chat_id=update.effective_chat.id, text=message)
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
        query.edit_message_text(text="Write the date you want to post questions (in YYYY-MM-DD format):") # Not a good design
        return TWO

    wb = open_workbook(update, context)
    if not wb:
        return ConversationHandler.END
    show_groups(wb, query, update, context)
    
def schedule_date(update: Update, context: CallbackContext):
    date = validate_date(update, context)
    if not date:
        return TWO
    context.user_data["date"] = date
    return THREE

def schedule_time(update: Update, context: CallbackContext):
    time = validate_time(update, context)
    if not time:
        return THREE
    context.user_data["time"] = time
    save_data(update, context)
    return ConversationHandler.END

def add_group(update: Update, context: CallbackContext):
    if update.message.chat.type == "private":
        return
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

def cancel(update, context):
    ConversationHandler.END

def set_jobs(update: Update, context: CallbackContext):
    if update.message.chat.type == "group":
        return
    if not check_admin(update, context):
        return
    sheet = open_workbook(update, context)["Schedule"]
    for row in sheet.iter_rows():
        time = create_datetime(row, 2, 3)
        context.job_queue.run_once(test, time, context=row[0].row)
    update.message.reply_text("You are all set!!! The questions are scheduled to run at the date and time you have chosen.")

def test(context: CallbackContext):
    # Runs the full list of questions according to time
    # Then checks the history and sees who answered at the right times
    # Makes a report and send it to the group admin
    # Deletes the row from the sheet.
    job = context.job
    bot = context.bot
    row_index = job.context
    session_start = datetime.datetime.now(pytz.timezone('Asia/Kolkata'))
    wb = load_workbook("custom/excel_sheet.xlsx")
    schedule = wb["Schedule"]
    for row in schedule.iter_rows():
        if row[0].row == row_index:
            questions_sheet_title = row[0].value
            group_list = row[1].value.strip(", ")

    questions = wb[questions_sheet_title]
    group_ids = group_ids_by_title(wb, group_list)

    # Send the question and sleep for the time limit
    for row in questions.iter_rows(min_row=3):
        if None in [row[1].value, row[2].value]:
            break
        send_message_to_ids(bot, group_ids, row[1].value)
        time.sleep(float(row[2].value) * 60)
    send_message_to_ids(bot, group_ids, message="time's up!")

    # start is saved in schedule and end can be taken right now
    send_report(
        wb=wb,
        context=context,
        chat_history=chat_history_sheet, 
        questions=questions, 
        group_ids=group_ids, 
        session_start=session_start
    )

    # Delete the schedule from the database
    schedule.delete_rows(row_index)
    wb.save(filename="custom/excel_sheet.xlsx")

def handle_user_responses(update: Update, context:CallbackContext):
    if update.message.chat.type != "group":
        return
    # Check if message is a reply to another message.
    reply_to = update.message.reply_to_message
    if not reply_to:
        return

    wb_main = load_workbook("custom/excel_sheet.xlsx")
    date_time = datetime.datetime.now(pytz.timezone('Asia/Kolkata'))
    schedule = in_run_time(wb_main, date_time)
    # Check if we currently are in a schedule
    if not schedule:
        return
    
    question_sheet = wb_main[schedule["sheet"]]
    question_number = get_question_number(question_sheet, reply_to.text)

    # Add these questions in attendance sheet.
    # Here we operate on the understanding that the schedule is deleted after it has been executed.
    wb_attendance = load_workbook("custom/attendance_sheet")

    # Find the username


def main():
    updater = Updater(token=TOKEN)
    
    dispatcher = updater.dispatcher

    # Garbage collect older schedules
    collect_garbage()

    start_handler = ConversationHandler(
        entry_points=[CommandHandler("start", start)],
        states={
            ONE: [CallbackQueryHandler(groups)],
            TWO: [MessageHandler(Filters.text, schedule_date)],
            THREE: [MessageHandler(Filters.text, schedule_time)],
        },
        fallbacks=[CommandHandler("cancel", cancel)]
    )

    document_handler = ConversationHandler(
        entry_points=[CommandHandler("addsheet", add_sheet_options)],
        states={
            ONE: [CallbackQueryHandler(ask_for_sheet)],
            TWO: [MessageHandler(Filters.document, incoming_document)]
        },
        fallbacks=[CommandHandler("cancel", cancel)]
    )

    dispatcher.add_handler(document_handler)
    dispatcher.add_handler(CommandHandler("add", add_group))
    dispatcher.add_handler(CommandHandler("set", set_jobs))
    dispatcher.add_handler(start_handler)
    dispatcher.add_handler(MessageHandler(Filters.text, handle_user_responses))

    updater.start_polling()

    updater.idle()

if __name__ == "__main__":
    main()