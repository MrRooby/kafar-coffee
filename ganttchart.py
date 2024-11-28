import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates


def get_availability(start_of_next_week, end_of_next_week, day_of_week, database):
    start_day = start_of_next_week.strftime('%Y.%m.%d')
    end_day = end_of_next_week.strftime('%Y.%m.%d')

    # Query to get availability data for Monday of the specified week
    query = '''
        SELECT u.NAME, a.START_TIME, a.END_TIME
        FROM AVAILABILITY a
        JOIN WEEK w ON a.WEEK_ID = w.id
        JOIN USERS u ON a.USER_ID = u.id
        WHERE w.START_DAY = %s AND w.END_DAY = %s AND a.DAY_OF_WEEK = %s
    '''
    database.execute_query(query, (start_day, end_day, day_of_week))
    results = database.cursor.fetchall()

    # Create a DataFrame with the retrieved data
    availability_df = pd.DataFrame(results, columns=['name', 'start_time', 'end_time'])

    # Convert timedelta to hours
    availability_df['start_time'] = availability_df['start_time'].apply(lambda x: x.total_seconds() / 3600 if pd.notnull(x) else None)
    availability_df['end_time'] = availability_df['end_time'].apply(lambda x: x.total_seconds() / 3600 if pd.notnull(x) else None)

    # Filter out rows with None values in start_time or end_time
    availability_df = availability_df.dropna(subset=['start_time', 'end_time'])

    # Calculate the duration
    availability_df['duration'] = availability_df['end_time'] - availability_df['start_time']

    return availability_df


def create_gantt_chart(start_of_next_week, end_of_next_week, day_of_week, database):
    switcher = {
        1: 'MON',
        2: 'TUE',
        3: 'WED',
        4: 'THU',
        5: 'FRI',
        6: 'SAT',
        7: 'SUN'
    }

    availability_df = get_availability(start_of_next_week, end_of_next_week, switcher[day_of_week], database)
    fig, ax = plt.subplots()

    # Plot the bars
    ax.barh(y=availability_df['name'], width=availability_df['duration'], left=availability_df['start_time'], color='skyblue')

    ax.set_xlim(left=6.0, right=19.0)

    # Format the x-axis to show only hours
    ax.xaxis.set_major_locator(mdates.HourLocator(interval=24))

    # Rotate the x-axis labels for better readability
    plt.xticks(rotation=45)

    plt.title(f'Availability on {switcher[day_of_week]} {start_of_next_week.strftime("%Y-%m-%d")} to {end_of_next_week.strftime("%Y-%m-%d")}')

    plt.tight_layout()

    plt.grid()

    # Save the figure as a PNG file
    plt.savefig(f'charts/{switcher[day_of_week]}.png')