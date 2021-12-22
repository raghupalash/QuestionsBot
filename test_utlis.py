from openpyxl import load_workbook
from utils import create_datetime, fill_database, open_workbook, show_sheets, test
from telegram.ext import ConversationHandler
from datetime import datetime, timedelta
"""
These functions are implemented for making testing the bot easier for us.
start_admin_manual for manual mode, and start_admin_auto for automatic.

What's going to happend is, I will give a special commnand to the bot in the pm, and it will
make a schedule for a given sheet, given group, time = time.now() + 10 seconds. Then it's going to set the schedule.
"""
def start_admin(update, context):
    # Load worksheet
    context.user_data["sheet"] = "Basketball"
    context.user_data["groups"] = ["okokokok"]
    context.user_data["date"] = datetime.today().strftime('%Y-%m-%d')
    time = datetime.now() + timedelta(seconds=5)
    context.user_data["time"] = time.strftime("%H:%M:%S")

    wb_main = open_workbook(update, context)
    wb_attendance = load_workbook("custom/attendance_sheet.xlsx")

    fill_database(wb_main, wb_attendance, context)

    # Set the schedule
    sheet = wb_main["Schedule"]
    for row in sheet.iter_rows():
        time = create_datetime(row, 2, 3)
        context.job_queue.run_once(test, time, context=row[4].value, name=f"job{row[0].row}")
    update.message.reply_text("You are all set!!! The questions are scheduled to run at the date and time you have chosen.")
