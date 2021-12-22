- Fill attendance sheet as shown in the videos.

- Don't disturb the first three sheets in the excel file(Schedule, Groups, count)

- While the bot is running, don't open any excel sheets inside the custom directory.

- To add the group to the bot's database, add the bot to the group and then use the command '/add'.

- To make the bot register a job, you will have to run /set. So if you make a new schedule, it won't run
unless you run /set. Also if you make a schedule and close the bot program, when you open the program again
you will have to run /set so that bot can set the jobs it needs to run.

- **IMPORTANT** after you run the bot the bot, you can only use /set command once, if you have already used /set
but now want to add another schedule, add the schedule as you normally would, turn off the bot, 
start it again and then use the set command.

- If you accidently use the /set command twice, then don't worry, just shut off the bot, run it again and do /set.

- It should be noted that adding a new question sheet or attendance sheet, is going to replace the previous sheet, 
so that only one sheet will exist at a time in bot's database.

- You can't run two tests at the exact same time. Also don't use /set command while a schedule is running.

- Always keep a backup copy of your data, in case bot crashes and the database is corrupted.

- Each user should have a username, otherwise logic will be affected.

- For the questions you don't won't to be included in report, write 0 as question number.

Process to make an executable - 
1. Install latest version of python from here - https://www.python.org/downloads/ 
(you can follow this tutorial - https://www.youtube.com/watch?v=uDbDIhR76H4)
2. open command line, change directory to the bot files directory.
3. in the command line run 'pip install -r requirements.txt' (this will install some libraries).
4. Open the credentials.py file and fill your bot token in the space provided.
    - Also put your user_id in the given space.
    - Also put your timezone in the given space.
5. You can either run the bot as it is, or you can make an executable.
 - To run bot as it is, in the project directory, open command line and enter - 'python bot.py'
 - if you want to make an executable, in the command line, enter 'pyinstaller -F bot.py'. This will make some folders
 in the current directory
 just open the folder called 'dist', cut the execuatable file and paste it to the root folder(our bot files folder)

 Your bot is complete. If you run into any problem, feel free to contact me.
