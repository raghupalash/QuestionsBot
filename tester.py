from telethon import TelegramClient, events
from datetime import date, datetime, timedelta

# Remember to use your own values from my.telegram.org!
api_id = 14308098
api_hash = '8023e064e5ff0cc369b78e202e684fa6'
client = TelegramClient('raghupalash', api_id, api_hash)


@client.on(events.NewMessage)
async def my_event_handler(event):
    print(event)
    if "YYYY-MM-DD" in event.raw_text:
        event.reply(str(date.today()))
    elif "HH:MM:SS" in event.raw_text:
        event.reply(str(datetime.now().time()).split(".")[0])
    elif "Data updated!" in event.raw_text:
        event.reply("/set")
    

with client:
    client.start()
    client.run_until_disconnected()
