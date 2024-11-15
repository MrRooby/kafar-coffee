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



def add_dyspo_record(user_id, start_of_next_week, end_of_next_week):
  start_day = start_of_next_week.strftime('%Y.%m.%d')
  end_day = end_of_next_week.strftime('%Y.%m.%d')
  DB_CURSOR.execute('SELECT ID FROM WEEK WHERE START_DAY = %s AND END_DAY = %s', (start_day, end_day))
  week_id = DB_CURSOR.fetchone()
  
  if week_id is None:
    DB_CURSOR.execute('INSERT INTO WEEK (START_DAY, END_DAY) VALUES (%s, %s)', (start_day, end_day))
    KAFAR_DB.commit()
    DB_CURSOR.execute('SELECT ID FROM WEEK WHERE START_DAY = %s AND END_DAY = %s', (start_day, end_day))
    week_id = DB_CURSOR.fetchone()
  
  week_id = week_id[0]
  KAFAR_DB.commit()


def dyspo_record_exists(user_id, start_of_next_week, end_of_next_week):
  start_day = start_of_next_week.strftime('%Y.%m.%d')
  end_day = end_of_next_week.strftime('%Y.%m.%d')
  
  DB_CURSOR.execute('''
    SELECT * FROM AVAILABILITY 
    WHERE USER_ID = %s 
    AND WEEK_ID = (SELECT ID FROM WEEK WHERE START_DAY = %s AND END_DAY = %s)
  ''', (user_id, start_day, end_day))
  result = DB_CURSOR.fetchall()
  
  return len(result) > 0


def add_form_sent_record(user_id, start_of_next_week, end_of_next_week):
  start_day = start_of_next_week.strftime('%Y.%m.%d')
  end_day = end_of_next_week.strftime('%Y.%m.%d')
  DB_CURSOR.execute('INSERT INTO FORM_SENT (USER_ID, START_DAY, END_DAY) VALUES (%s, %s, %s)', (user_id, start_day, end_day))
  KAFAR_DB.commit()


def is_form_sent_record_exists(user_id, start_of_next_week, end_of_next_week):
  start_day = start_of_next_week.strftime('%Y.%m.%d')
  end_day = end_of_next_week.strftime('%Y.%m.%d')
  
  DB_CURSOR.execute('SELECT * FROM FORM_SENT WHERE USER_ID = %s AND START_DAY = %s AND END_DAY = %s', (user_id, start_day, end_day))
  result = DB_CURSOR.fetchone()
  
  return result is not None


def check_is_user_in_database(user_id):
  DB_CURSOR.execute('SELECT * FROM USERS WHERE ID = %s', (user_id,))
  user_exists = DB_CURSOR.fetchone()
  
  # If user does not exist, add them to the USERS table
  if not user_exists:
    DB_CURSOR.execute('INSERT INTO USERS (ID) VALUES (%s)', (user_id,))
    KAFAR_DB.commit()


def add_dyspo_to_database(user_id, dyspo, start_of_next_week, end_of_next_week):
  check_is_user_in_database(user_id)
  
  # If dyspo record exists, delete the existing availability records for the user and week
  if dyspo_record_exists(user_id, start_of_next_week, end_of_next_week):
    DB_CURSOR.execute('''
      DELETE FROM AVAILABILITY 
      WHERE USER_ID = %s 
      AND WEEK_ID = (SELECT ID FROM WEEK WHERE START_DAY = %s AND END_DAY = %s)
    ''', (user_id, start_of_next_week.strftime('%Y.%m.%d'), end_of_next_week.strftime('%Y.%m.%d')))
    KAFAR_DB.commit()
  
  # Insert the data into the DYSPO table
  add_dyspo_record(user_id, start_of_next_week, end_of_next_week)
  
  # Prepare the data for insertion into AVAILABILITY table
  for day, hours in dyspo.items():
    day_of_week = day[:3].upper()
    if hours == ["dowolnie"]:
      start_time = "06:00:00"
      end_time = "19:00:00"
    elif hours == ["out"]:
      start_time = None
      end_time = None
    else:
      start_time = f"{hours[0]}:00"
      end_time = f"{hours[1]}:00" if len(hours) > 1 else f"{hours[0]}:00"
    
    DB_CURSOR.execute('''
      INSERT INTO AVAILABILITY (USER_ID, WEEK_ID, DAY_OF_WEEK, START_TIME, END_TIME)
      VALUES (%s, (SELECT ID FROM WEEK WHERE START_DAY = %s AND END_DAY = %s), %s, %s, %s)
    ''', (user_id, start_of_next_week.strftime('%Y.%m.%d'), end_of_next_week.strftime('%Y.%m.%d'), day_of_week, start_time, end_time))
  
  KAFAR_DB.commit()
