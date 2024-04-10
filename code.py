import pandas as pd
### checking the working days
def get_previous_working_day(date):
    # Get the previous day
    previous_day = date - pd.Timedelta(days=1)
    # Check if the previous day is a weekday (Monday to Friday)
    while previous_day.weekday() >= 5:  # Monday: 0, Tuesday: 1, ..., Friday: 4
        previous_day -= pd.Timedelta(days=1)
    return previous_day

def check_rank(df):
    try:
        # Prompt user for input
        day = input("Enter the day (YYYY-MM-DD): ")
        time = input("Enter the time (HH:MM:SS): ")

        # Convert user input to datetime format
        check_datetime = pd.to_datetime(day + ' ' + time)

        # Get the current and previous working days
        current_day = check_datetime.date()
        previous_working_day = get_previous_working_day(current_day)

        # Filter DataFrame for the current day and specified time
        current_df = df[(df['Datetime'].dt.date == current_day) & (df['Datetime'].dt.time == check_datetime.time())]

        if len(current_df) == 0:
            print("No data available for the specified day and time.")
        else:
            print("Rank of volume at {} on {} is: {}".format(time, current_day.strftime('%Y-%m-%d'), current_df['Rank'].iloc[0]))
    except ValueError as e:
        print("Error:", e)

# Read the Excel file into a DataFrame
df = pd.read_excel("C:\\Users\\sahub\\Downloads\\excel.xlsx")

# Convert the 'Date' and 'Time' columns to datetime format
df['Date'] = pd.to_datetime(df['Date'])
df['Time'] = df['Time'].astype(str)

# Combine the 'Date' and 'Time' columns into a single datetime column
df['Datetime'] = df['Date'] + pd.to_timedelta(df['Time'])

# Sort the DataFrame by 'Datetime' column in ascending order
df = df.sort_values(by='Datetime')

# Calculate the rank of volume for each time across the last 5 days
df['Rank'] = df.groupby(df['Datetime'].dt.time)['Volume'].transform(lambda x: x.rolling(window=5).apply(lambda x: pd.Series(x).rank(ascending=False).iloc[0]))

# Display the DataFrame
print(df)

# After displaying the DataFrame, asking the user to specify any specific date and time user want to check the rank
check_rank(df)
