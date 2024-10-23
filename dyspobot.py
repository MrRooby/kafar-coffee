import discord
from discord.ext import commands, tasks
import os
from dotenv import load_dotenv
from datetime import datetime, timedelta
import xlsxwriter
from openpyxl import load_workbook, Workbook
from openpyxl.chart import BarChart, Reference, Series
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font


# Global variables
user_dyspo = {}
user_messages = {}
thick_right_border = Border(left=Side(style='none'), 
                    right=Side(style='thick'), 
                    top=Side(style='none'), 
                    bottom=Side(style='none'))

thick_right_and_left_border = Border(left=Side(style='thick'), 
                    right=Side(style='thick'), 
                    top=Side(style='none'), 
                    bottom=Side(style='none'))

notified_users = [] # TODO: You need to save these data structures so they don't reset after reboot

fill_colours = {
  "Monday": PatternFill(start_color="FF9999", end_color="FF9999", fill_type="solid"),  # Desaturated Red
  "Tuesday": PatternFill(start_color="FFCC99", end_color="FFCC99", fill_type="solid"),  # Desaturated Orange
  "Wednesday": PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid"),  # Desaturated Yellow
  "Thursday": PatternFill(start_color="99FF99", end_color="99FF99", fill_type="solid"),  # Desaturated Green
  "Friday": PatternFill(start_color="99CCFF", end_color="99CCFF", fill_type="solid"),  # Desaturated Light Blue
  "Saturday": PatternFill(start_color="9999FF", end_color="9999FF", fill_type="solid"),  # Desaturated Blue
  "Sunday": PatternFill(start_color="CC99FF", end_color="CC99FF", fill_type="solid")   # Desaturated Purple
}


# Function to load the token from .env file
def load_token():
  load_dotenv()
  return os.getenv('TOKEN')

# Function to generate the dates for the current dyspo
def get_next_week_dates():
  today = datetime.today()
  start_of_next_week = today + timedelta(days=(7 - today.weekday()))
  end_of_next_week = start_of_next_week + timedelta(days=6)
  return start_of_next_week, end_of_next_week

# Function to initialize the bot with its permissions
def initialize_bot():
  intents = discord.Intents.default()
  intents.messages = True
  intents.guilds = True
  intents.message_content = True
  intents.members = True

  bot = commands.Bot(command_prefix='!', intents=intents)
  return bot


def current_workbook_name():
  start_of_next_week, end_of_next_week = get_next_week_dates()
  return f"dyspozycje/dyspo[{start_of_next_week.strftime('%d')}-{end_of_next_week.strftime('%d.%m')}].xlsx"

# Function to create a new spreadsheet for the next week
def make_spredsheet():
  global current_workbook, current_worksheet
  file_name = current_workbook_name()
  current_workbook = xlsxwriter.Workbook(file_name)
  current_worksheet = current_workbook.add_worksheet()
  days_of_week = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
  for row, day in enumerate(days_of_week):
    current_worksheet.write(row + 1, 0, day)
  #   current_worksheet
  # current
  current_workbook.close()


def update_spredsheet(user, dyspo):
  global num_of_dyspo_from_users
  global current_worksheet, current_workbook
  global thick_right_border
  global fill_colours

  if not os.path.exists(current_workbook_name()):
    make_spredsheet()
  current_workbook = load_workbook(current_workbook_name())
  current_worksheet = current_workbook.active
  
  user_col = None
  for col in range(2, current_worksheet.max_column + 1):
    if current_worksheet.cell(row = 1, column = col).value == user:
      user_col = col
      break

  row = 1
  
  if user_col is None:
    user_col = current_worksheet.max_column + 1
    current_worksheet.merge_cells(start_row = row, start_column = user_col, end_row = row, end_column = user_col + 1)
    
    current_worksheet.cell(row = row, column = user_col).value = user
    current_worksheet.cell(row=row, column=user_col).border = thick_right_and_left_border

  row += 1

  for day, hours in dyspo.items():
    if len(hours) == 2:
      current_worksheet.cell(row = row, column = user_col, value=hours[0])
      current_worksheet.cell(row = row, column = user_col + 1, value=hours[1])
      current_worksheet.cell(row = row, column = user_col + 1).border = thick_right_border
    else:
      current_worksheet.cell(row = row, column = user_col, value=hours[0])
      current_worksheet.cell(row = row, column = user_col + 1, value=hours[0])
      current_worksheet.cell(row = row, column = user_col + 1).border = thick_right_border

    current_worksheet.cell(row=row, column=user_col).fill = fill_colours[day]
    current_worksheet.cell(row=row, column=user_col + 1).fill = fill_colours[day]

    row += 1

  # create_gantt_chart(current_workbook, current_worksheet)

  current_workbook.save(current_workbook_name())


def add_to_spredsheet(user_id, day, hours):
  global num_of_dyspo_from_users

  if(user_id not in user_dyspo):
    user_dyspo[user_id] = {}

  # Handle special cases for "DO + 30min" and "OD + 30min"
  adjusted_hours = []
  do_30min_selected = False
  od_30min_selected = False

  for hour in hours:
    if hour == "DO + 30min":
      do_30min_selected = True
    elif hour == "OD + 30min":
      od_30min_selected = True
    else:
      adjusted_hours.append(hour)

  if do_30min_selected and adjusted_hours:
    adjusted_hours[0] = adjust_time(adjusted_hours[0], 30)
  if od_30min_selected and len(adjusted_hours) > 1:
    adjusted_hours[1] = adjust_time(adjusted_hours[1], 30)

  user_dyspo[user_id][day] = adjusted_hours


def adjust_time(time_str, minutes):
  hour, minute = map(int, time_str.split(":"))
  minute += minutes
  if minute >= 60:
    hour += 1
    minute -= 60
  return f"{hour:02d}:{minute:02d}"


def create_gantt_chart(workbook, worksheet):
  chart = BarChart()
  chart.type = "bar"
  chart.style = 10
  chart.title = "Gantt Chart"
  chart.y_axis.title = 'Days'
  chart.x_axis.title = 'Hours'

  data = Reference(worksheet, min_col=2, min_row=1, max_col=worksheet.max_column, max_row=worksheet.max_row)
  categories = Reference(worksheet, min_col=1, min_row=2, max_row=worksheet.max_row)
  chart.add_data(data, titles_from_data=True)
  chart.set_categories(categories)

  worksheet.add_chart(chart, "B10")


async def notify_users(users):
  global notified_users
  for notified_user in notified_users:
    await notified_user.send(f"{users.name} wypełnił dyspo.")


# A class for a select menu for a day of the week
class WorkHoursSelect(discord.ui.Select):
  def __init__(self, placeholder):
    global num_of_dyspo_from_users
    options =  [discord.SelectOption(label="dowolnie")] + [discord.SelectOption(label="out")] + [
      discord.SelectOption(label=f"{hour}:{"00"}")
      for hour in range(6, 20)
    ] + [discord.SelectOption(label="OD + 30min")] + [discord.SelectOption(label="DO + 30min")]
    super().__init__(placeholder=placeholder, min_values=1, max_values=4, options=options)
    

  async def callback(self, interaction: discord.Interaction):
    await interaction.response.defer(ephemeral=True)
    # Ensure "DO + 30min" and "OD + 30min" are not selected on their own
    if ("DO + 30min" in self.values or "OD + 30min" in self.values) and len(self.values) == 1:
      await interaction.followup.send("Options 'DO + 30min' and 'OD + 30min' must be selected with at least one other option.", ephemeral=True)
      return
    
    # Ensure "DO + 30min" and "OD + 30min" are not selected with "out" or "dowolnie"
    if ("DO + 30min" in self.values or "OD + 30min" in self.values) and ("out" in self.values or "dowolnie" in self.values):
      await interaction.followup.send("Options 'DO + 30min' and 'OD + 30min' cannot be selected with 'out' or 'dowolnie'.", ephemeral=True)
      return

    add_to_spredsheet(interaction.user.name, self.placeholder, self.values)


# 2 WorkHoursView classes due to discord limitation of 4 select options per view
class WorkHoursView1(discord.ui.View):
  def __init__(self):
    super().__init__()
    self.add_item(WorkHoursSelect("Monday"))
    self.add_item(WorkHoursSelect("Tuesday"))
    self.add_item(WorkHoursSelect("Wednesday"))
    self.add_item(WorkHoursSelect("Thursday"))

class WorkHoursView2(discord.ui.View):
  def __init__(self):
    super().__init__()
    self.add_item(WorkHoursSelect("Friday"))
    self.add_item(WorkHoursSelect("Saturday"))
    self.add_item(WorkHoursSelect("Sunday"))
    self.add_item(ConfirmButton())


class ConfirmButton(discord.ui.Button):
  def __init__(self):
    super().__init__(label="Confirm", style=discord.ButtonStyle.primary)

  async def callback(self, interaction: discord.Interaction):
    user_id = interaction.user.name
    if(len(user_dyspo[user_id]) < 7):
      await interaction.response.send_message("Wypełnij wszystkie dni tygodnia", ephemeral=True)
      return
    update_spredsheet(interaction.user.name, user_dyspo[interaction.user.name])
    confirmation_message = "Dyspo przesłane!\n\nWybrane godziny:\n"
    for day, hours in user_dyspo[user_id].items():
      confirmation_message += f"{day}: {'-'.join(hours)}\n"
    print(interaction.user.id)
    await interaction.response.send_message(confirmation_message, ephemeral=True)

    await notify_users(interaction.user)

    await interaction.message.delete()

    if user_id in user_messages:
      for message in user_messages[user_id]:
          try:
            await message.delete()
          except discord.errors.NotFound:
            print(f"Message {message.id} not found, skipping deletion.")
      del user_messages[user_id]
      

async def send_select_menus(user):
  start_date, end_date = get_next_week_dates()
  view1 = WorkHoursView1()
  view2 = WorkHoursView2()
  date = "[Dyspo " + start_date.strftime('%d') + "-" + end_date.strftime('%d.%m') + "]\n"
  instructions = "W celu wybrania połowy godziny dodaj do wybranej godziny 'DO + 30min' lub 'OD + 30min'.\n"
  message1 = await user.send(date + instructions, view=view1)
  message2 = await user.send(view=view2)
  
  # Store the messages to delete later
  user_messages[user.name] = [message1, message2]





bot = initialize_bot()
make_spredsheet()

@bot.event
async def on_ready():
  print(f"Logged in as {bot.user}")
  send_dyspo.start()


@bot.command()
async def dyspo(ctx):
  await send_select_menus(ctx.author)


@bot.command()
async def excel(ctx):
  if os.path.exists(current_workbook_name()):
    await ctx.send(file=discord.File(current_workbook_name()))
  else:
    await ctx.send("No dyspo sheet found")


@bot.command()
async def notifon(ctx):
  global notified_users

  if ctx.author in notified_users:
    await ctx.send("Powiadomienia już są włączone!")
    return
  else:
    notified_users.append(ctx.author)
    await ctx.send(":green_circle: Powiadomienia włączone!")


@bot.command()
async def notifoff(ctx):
  global notified_users

  if ctx.author not in notified_users:
    await ctx.send("Powiadomienia już są wyłączone!")
  else:
    notified_users.remove(ctx.author)
    await ctx.send(":red_circle: Wyłączono powiadomienia!")


@tasks.loop(hours=1) # The message will be delivered every wednesday form 16 to 16:59 depending when was the program started
async def send_dyspo():
  now = datetime.now()
  if (now.weekday() == 2 and now.hour == 16):  # Check if it's Wednesday at 4 PM
    for guild in bot.guilds:
      for member in guild.members:
        if not member.bot and discord.utils.get(member.roles, name="beboki"):  # Skip bot accounts and check for role "beboki"
          await send_select_menus(member)


@tasks.loop(hours=1)
async def create_spredsheet():
  global current_worksheet, current_workbook, num_of_dyspo_from_users
  now = datetime.now()
  if (now.weekday() == 0 and now.hour == 0):
    current_workbook, current_worksheet = make_spredsheet()
    num_of_dyspo_from_users = 0
    print("New dyspo sheet created")


bot.run(load_token())