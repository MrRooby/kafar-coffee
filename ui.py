import discord
from excel import *
from database import *
from dyspobot import DyspoBot
from util import get_next_week_dates


# A class for a select menu for a day of the week
class WorkHoursSelect(discord.ui.Select):
  def __init__(self, placeholder, spredsheet: Spredsheet):
    global num_of_dyspo_from_users
    options =  [discord.SelectOption(label="dowolnie")] + [
                discord.SelectOption(label="out")]      + [
                discord.SelectOption(label=f"{hour}:00") for hour in range(6, 20)]
    
    super().__init__(placeholder=placeholder, min_values=1, max_values=2, options=options)
    
    self.spredsheet = spredsheet

  
  async def callback(self, interaction: discord.Interaction):
    await interaction.response.defer(ephemeral=True)
    
    # Ensure that starting and ending time is selected
    #TODO

    self.spredsheet.add_dyspo(interaction.user.name, self.placeholder, self.values)


# 2 WorkHoursView classes due to discord limitation of 4 select options per view
class WorkHoursView1(discord.ui.View):
  def __init__(self, spredsheet: Spredsheet):
    super().__init__()
    self.add_item(WorkHoursSelect("Monday"),    spredsheet)
    self.add_item(WorkHoursSelect("Tuesday"),   spredsheet)
    self.add_item(WorkHoursSelect("Wednesday"), spredsheet)
    self.add_item(WorkHoursSelect("Thursday"),  spredsheet)
    self.timeout = 1800  # 30 minutes

class WorkHoursView2(discord.ui.View):
  def __init__(self, spredsheet: Spredsheet, bot: DyspoBot, user_messages):
    super().__init__()
    self.add_item(WorkHoursSelect("Friday"),    spredsheet)
    self.add_item(WorkHoursSelect("Saturday"),  spredsheet)
    self.add_item(WorkHoursSelect("Sunday"),    spredsheet)
    self.add_item(ConfirmButton(spredsheet))
    self.timeout = 1800  # 30 minutes
    self.bot - bot
    self.user_messages = user_messages

class ConfirmButton(discord.ui.Button):
  def __init__(self, spredsheet: Spredsheet, database: Database, user_dyspo: dict, bot: DyspoBot):
    super().__init__(label="Confirm", style=discord.ButtonStyle.primary)
    self.timeout = 1800  # 30 minutes
    self.spredsheet = spredsheet
    self.database = database
    self.user_dyspo = user_dyspo
    self.bot = bot

  async def callback(self, interaction: discord.Interaction):
    user_id = interaction.user.name

    if(len(self.user_dyspo[user_id]) < 7):
      await interaction.response.send_message("Wypełnij wszystkie dni tygodnia", ephemeral=True)
      return

    start_of_next_week, end_of_next_week = get_next_week_dates()

    self.spredsheet.update_spredsheet(interaction.user.name, 
                                      self.user_dyspo[interaction.user.name])
    self.database.add_dyspo_to_database(interaction.user.id, 
                                        self.user_dyspo[interaction.user.name], 
                                        start_of_next_week, end_of_next_week)

    days_order = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
    confirmation_message = "Dyspo przesłane!\n\nWybrane godziny:\n"

    for day in days_order:
      if day in self.user_dyspo[user_id]:
        hours = self.user_dyspo[user_id][day]
        confirmation_message += f"{day}: {'-'.join(hours)}\n"
    print(interaction.user.id)

    await interaction.response.send_message(confirmation_message, ephemeral=True)

    await self.bot.notify_users(interaction.user)

    await interaction.message.delete()

    if user_id in self.user_messages:
      for message in self.user_messages[user_id]:
          try:
            await message.delete()
          except discord.errors.NotFound:
            print(f"Message {message.id} not found, skipping deletion.")
      del self.user_messages[user_id]