import os
from dotenv import load_dotenv
import mysql.connector
import time


class Database:
    def __init__(self):
        load_dotenv()
        self.connect_to_database()


    def connect_to_database(self):
        self.connection = mysql.connector.connect(
            host     = "uk01-sql.pebblehost.com",
            user     = "customer_862222_kafar-coffee-dyspozycje",
            password = os.getenv('DB_PASSWORD'),
            database = "customer_862222_kafar-coffee-dyspozycje",
            port     = "3306")
        
        self.cursor = self.connection.cursor()
        

    def close_connection(self):
        self.cursor.close()
        self.connection.close()


    def execute_query(self, query, params=None):
        if not self.connection.is_connected():
            print("Connection to MySQL server lost. Reconnecting...")
            self.connect_to_database()
            if self.connection.is_connected():
                print("Reconnected to MySQL server.")
        
        self.cursor.execute(query, params)
        return self.cursor


    def add_dyspo_record(self, start_of_next_week, end_of_next_week):
        start_day = start_of_next_week.strftime('%Y.%m.%d')
        end_day   = end_of_next_week.strftime('%Y.%m.%d')
       
        cursor = self.execute_query('''
            SELECT ID FROM WEEK 
            WHERE START_DAY = %s 
            AND END_DAY = %s
            ''', (start_day, end_day))
       
        week_id = cursor.fetchone()

        if week_id is None:
            self.execute_query('''
                INSERT INTO WEEK (START_DAY, END_DAY)
                VALUES (%s, %s)
                ''', (start_day, end_day))

            self.connection.commit()
            
            cursor = self.execute_query('''
                SELECT ID FROM WEEK
                WHERE START_DAY = %s 
                AND END_DAY = %s
                ''', (start_day, end_day))

            week_id = cursor.fetchone()

        week_id = week_id[0]
        self.connection.commit()


    def dyspo_record_exists(self, user_id, start_of_next_week, end_of_next_week):
        start_day = start_of_next_week.strftime('%Y.%m.%d')
        end_day = end_of_next_week.strftime('%Y.%m.%d')

        cursor = self.execute_query('''
            SELECT * FROM AVAILABILITY 
            WHERE USER_ID = %s 
            AND WEEK_ID = (
                SELECT ID FROM WEEK 
                WHERE START_DAY = %s 
                AND END_DAY = %s)
        ''', (user_id, start_day, end_day))

        result = cursor.fetchall()

        return len(result) > 0


    def add_form_sent_record(self, user_id, start_of_next_week, end_of_next_week):
        start_day = start_of_next_week.strftime('%Y.%m.%d')
        end_day = end_of_next_week.strftime('%Y.%m.%d')

        self.execute_query('''
            INSERT INTO FORM_SENT (USER_ID, START_DAY, END_DAY) 
            VALUES (%s, %s, %s)
            ''', (user_id, start_day, end_day))

        self.connection.commit()


    def is_form_sent_record_exists(self, user_id, start_of_next_week, end_of_next_week):
        start_day = start_of_next_week.strftime('%Y.%m.%d')
        end_day = end_of_next_week.strftime('%Y.%m.%d')
        
        cursor = self.execute_query('''
            SELECT * FROM FORM_SENT 
            WHERE USER_ID = %s 
            AND START_DAY = %s 
            AND END_DAY = %s
            ''', (user_id, start_day, end_day))

        result = cursor.fetchone()
        
        return result is not None


    def check_is_user_in_database(self, user_id, retries=3):
        cursor = self.execute_query('''
            SELECT * FROM USERS 
            WHERE ID = %s
            ''', (user_id,))
        
        user_exists = cursor.fetchone()
                
        if not user_exists:
            self.execute_query('''
                INSERT INTO USERS (ID) 
                VALUES (%s)
                ''', (user_id,))
       
            self.connection.commit()
       
        return

    def add_dyspo_to_database(self, user_id, dyspo, start_of_next_week, end_of_next_week):
        self.check_is_user_in_database(user_id)
        
        # If dyspo record exists, delete the existing availability records for the user and week
        if self.dyspo_record_exists(user_id, start_of_next_week, end_of_next_week):
            self.execute_query('''
                DELETE FROM AVAILABILITY 
                WHERE USER_ID = %s 
                AND WEEK_ID = (
                    SELECT ID FROM WEEK 
                    WHERE START_DAY = %s 
                    AND END_DAY = %s)
            ''', (user_id, start_of_next_week.strftime('%Y.%m.%d'), end_of_next_week.strftime('%Y.%m.%d')))
            self.connection.commit()
        
        # Insert the data into the DYSPO table
        self.add_dyspo_record(start_of_next_week, end_of_next_week)
        
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
            
            self.execute_query('''
                INSERT INTO AVAILABILITY (USER_ID, WEEK_ID, DAY_OF_WEEK, START_TIME, END_TIME)
                VALUES (
                    %s, 
                    (SELECT ID FROM WEEK 
                    WHERE START_DAY = %s 
                    AND END_DAY = %s),
                    %s, 
                    %s, 
                    %s)
            ''', (user_id, start_of_next_week.strftime('%Y.%m.%d'), end_of_next_week.strftime('%Y.%m.%d'), day_of_week, start_time, end_time))
        self.connection.commit()


    def notified_users(self):
        self.execute_query('''
            SELECT USER_ID 
            FROM NOTIFIED_USERS
            ''')
        return self.cursor.fetchall()


    def is_notified(self, user_id):
        self.execute_query('''
            SELECT * FROM NOTIFIED_USERS 
            WHERE USER_ID = %s
            ''', (user_id,))
        result = self.cursor.fetchone()
        
        if result is None:
            return False
        else:
            return True


    def add_to_notified_users(self, user_id):
        self.execute_query('''
            INSERT INTO NOTIFIED_USERS (USER_ID) 
            VALUES (%s)
            ''', (user_id,))
        self.connection.commit()


    def delete_from_notified_users(self, user_id):
        self.execute_query('''
            DELETE FROM NOTIFIED_USERS 
            WHERE USER_ID = %s
            ''', (user_id,))
        self.connection.commit()