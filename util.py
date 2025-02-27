import os
from datetime import datetime, timedelta
from dotenv import load_dotenv

def load_token():
  """Loads the discord token from the environment variables.

  Returns:
    str: The token value, or None if the token is not found.
  """
  load_dotenv()
  return os.getenv('TOKEN')
  

def get_next_week_dates():
  """Generates the starting and ending dates for the current week.

  Returns:
    tuple: A tuple containing:
      - datetime: The date of the first day of the week.
      - datetime: The date of the last day of the week.
  """
  today = datetime.today()
  start_of_next_week = today + timedelta(days=(7 - today.weekday()))
  end_of_next_week = start_of_next_week + timedelta(days=6)
  return start_of_next_week, end_of_next_week