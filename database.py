import os
from dotenv import load_dotenv
import mysql.connector
from datetime import datetime, timedelta

def load_db_password():
  load_dotenv()
  return os.getenv('DB_PASSWORD')

KAFAR_DB = mysql.connector.connect(
    host="uk01-sql.pebblehost.com",
    user="customer_862222_kafar-coffee-dyspozycje",
    password=load_db_password(),
    database="customer_862222_kafar-coffee-dyspozycje",
    port="3306"
)

DB_CURSOR = KAFAR_DB.cursor()

def get_next_week_dates():
  today = datetime.today()
  start_of_next_week = today + timedelta(days=(7 - today.weekday()))
  end_of_next_week = start_of_next_week + timedelta(days=6)
  return start_of_next_week, end_of_next_week

def add_dyspo_record(user_id):
  start_of_next_week, end_of_next_week  = get_next_week_dates()
  start_day = start_of_next_week.strftime('%Y.%m.%d')
  end_day = end_of_next_week.strftime('%Y.%m.%d')
  DB_CURSOR.execute('INSERT INTO DYSPO (USER_ID, START_DAY, END_DAY) VALUES (%s, %s, %s)', (user_id, start_day, end_day))
  KAFAR_DB.commit()


def is_dyspo_record_exists(user_id):
  start_of_next_week, end_of_next_week  = get_next_week_dates()
  start_day = start_of_next_week.strftime('%Y.%m.%d')
  end_day = end_of_next_week.strftime('%Y.%m.%d')
  
  DB_CURSOR.execute('SELECT * FROM DYSPO WHERE USER_ID = %s AND START_DAY = %s AND END_DAY = %s', (user_id, start_day, end_day))
  result = DB_CURSOR.fetchone()
  
  return result is not None


def add_form_sent_record(user_id):
  start_of_next_week, end_of_next_week = get_next_week_dates()
  start_day = start_of_next_week.strftime('%Y.%m.%d')
  end_day = end_of_next_week.strftime('%Y.%m.%d')
  DB_CURSOR.execute('INSERT INTO FORM_SENT (USER_ID, START_DAY, END_DAY) VALUES (%s, %s, %s)', (user_id, start_day, end_day))
  KAFAR_DB.commit()


def is_form_sent_record_exists(user_id):
  start_of_next_week, end_of_next_week = get_next_week_dates()
  start_day = start_of_next_week.strftime('%Y.%m.%d')
  end_day = end_of_next_week.strftime('%Y.%m.%d')
  
  DB_CURSOR.execute('SELECT * FROM FORM_SENT WHERE USER_ID = %s AND START_DAY = %s AND END_DAY = %s', (user_id, start_day, end_day))
  result = DB_CURSOR.fetchone()
  
  return result is not None