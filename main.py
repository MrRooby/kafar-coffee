from discord.ext import tasks
from datetime    import datetime
from database    import Database
from excel       import Spredsheet
from dyspobot    import DyspoBot
from util        import load_token

user_dyspo = {}
user_messages = {}

kafarDB = Database()
spredsheet = Spredsheet()
dyspoBot = DyspoBot()

bot_tick_minutes = 1


###############################################################################################################
############################################## Bot startup ####################################################
###############################################################################################################

dyspoBot.run(load_token())

@dyspoBot.event
async def on_ready():
  send_dyspo       .start()
  create_spredsheet.start()
  print(f"Logged in as {dyspoBot.user}")


###############################################################################################################
############################################## Bot commands ###################################################
###############################################################################################################

@dyspoBot.command()
async def dyspo(ctx):
   await dyspoBot.send_select_menus(ctx.author)


@dyspoBot.command()
async def excel(ctx):
    await dyspoBot.send_spredsheet(ctx, spredsheet)


@dyspoBot.command()
async def notifon(ctx):
    await dyspoBot.turn_on_notif_for_user(ctx, kafarDB)


@dyspoBot.command()
async def notifoff(ctx):
    await dyspoBot.turn_off_notif_for_user(ctx, kafarDB)


@dyspoBot.command()
async def wykres(ctx):
    await dyspoBot.send_gantt_chart(ctx, kafarDB)


###############################################################################################################
############################################### Bot loops #####################################################
###############################################################################################################

@tasks.loop(minutes = 1) # The form will be sent every Wednesday at 4 PM
async def send_dyspo():
    now = datetime.now()

    if now.weekday() == 2 and now.hour == 15:
        await dyspoBot.send_dyspo_to_users(kafarDB)


@tasks.loop(minutes = 1) # Create new excel spredsheet at the start of the week
async def create_spredsheet():
    now = datetime.now()

    if (now.weekday() == 0 and now.hour == 0):
        await dyspoBot.create_new_spredsheet(spredsheet)