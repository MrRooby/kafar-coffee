import discord
from discord.ext import commands

from ui import *
from util import *
from ganttchart import create_gantt_chart


class DyspoBot:
  def __init__(self, bot_tick_minutes: float):
    """Creates an initialized bot object.
    
    Bot is given permissions for messages, guilds,
    message_content and members.

    Parameters:
      bot_tick_minutes: Determines how often bot will check
                        if it is time to send dyspo reminder 
                        and when to create new spredsheet
    
    Returns:
      Bot: Initialized bot object.
    """
    self.bot_tick_minutes = bot_tick_minutes
    intents = discord.Intents.default()
    intents.messages = True
    intents.guilds = True
    intents.message_content = True
    intents.members = True

    self = commands.Bot(command_prefix='!', intents=intents)


  async def notify_users(self, users, database: Database):
    for user_id in database.notified_users():
      user = await self.fetch_user(user_id[0])
      await user.send(f"{users.name} wypełnił dyspo.")


  async def send_select_menus(self, user, spredsheet: Spredsheet, user_messages):
    start_date, end_date = get_next_week_dates()
    
    view1 = WorkHoursView1(spredsheet)
    view2 = WorkHoursView2(spredsheet)
    
    date = "[Dyspo " + start_date.strftime('%d') + "-" + end_date.strftime('%d.%m') + "]\n"
    
    message1 = await user.send(date, view=view1)
    message2 = await user.send(view=view2)
    
    # Store the messages to delete later
    user_messages[user.name] = [message1, message2]

  
  async def turn_on_notif_for_user(self, ctx, database: Database):
    if database.is_notified(ctx.author.id):
      await ctx.send("Powiadomienia już są włączone!")
    else:
      database.add_to_notified_users(ctx.author.id)
      await ctx.send(":green_circle: Włączono powiadomienia!")


  async def turn_off_notif_for_user(self, ctx, database: Database):
    if database.is_notified(ctx.author.id):
      database.delete_from_notified_users(ctx.author.id)
      await ctx.send(":red_circle: Wyłączono powiadomienia!")
    else:
      await ctx.send("Powiadomienia już są wyłączone!")

  
  async def send_spredsheet(self, ctx, spredsheet: Spredsheet):
    if os.path.exists(spredsheet.current_workbook_name()):
      await ctx.send(file=discord.File(spredsheet.current_workbook_name()))
    else:
      await ctx.send("No dyspo sheet found")

  
  async def send_gantt_chart(self, ctx, database: Database):
    start_of_next_week, end_of_next_week = get_next_week_dates()

    switcher = {
      1: 'MON',
      2: 'TUE',
      3: 'WED',
      4: 'THU',
      5: 'FRI',
      6: 'SAT',
      7: 'SUN'
    }
    
    for day in range(1, 8):
      create_gantt_chart(start_of_next_week, end_of_next_week, day, database)
      await ctx.send(file=discord.File(f"charts/{switcher[day]}.png"))


  async def send_dyspo_to_users(self, database: Database):
    start_of_next_week, end_of_next_week = get_next_week_dates()

    for guild in self.guilds:
      for member in guild.members:
        if ( not database.is_form_sent_record_exists(member.id, start_of_next_week, end_of_next_week) 
            and not member.bot 
            and discord.utils.get(member.roles, name="beboki") 
            and database.dyspo_record_exists(member.id, start_of_next_week, end_of_next_week) == False ):
          await self.send_select_menus(member)
          database.add_form_sent_record(member.id, start_of_next_week, end_of_next_week)


  async def create_new_spredsheet(spredsheet: Spredsheet):
    spredsheet.init_spredsheet()
    print("New dyspo sheet created")