import pandas as pd
import numpy as np
from tqdm import tqdm
import yfinance as yf

# Read the input data from list.xlsx
input_data = pd.read_excel("list.xlsx")

# Create an empty DataFrame to store the calculated data
output_data = pd.DataFrame(columns=['Ticker', 'Link', 'Name', 'Industry', 'Market Cap',
                                    'Standard Deviation', 'Measured Move',
                                    'Date A', 'Price A', 'Date B', 'Price B',
                                    'Date C', 'Price C', 'Current Price'])

# Iterate over each ticker
for _, row in tqdm(input_data.iterrows(), total=len(input_data), desc="Processing Tickers"):
    ticker = row['Ticker']
    link = row['Link']
    name = row['Name']
    industry = row['Industry']
    market_cap = row['Market Cap']

    # Fetch the historical data for the ticker
    ticker_data = yf.download(ticker, period='90d')

    # Calculate the linear regression line using numpy's polyfit
    x = np.arange(len(ticker_data))
    y = ticker_data['Close'].values
    slope, intercept = np.polyfit(x, y, 1)
    regression_line = slope * x + intercept

    # Calculate the standard deviation
    std_dev = np.std(y)

    # Calculate the number of standard deviations away from the regression line
    deviations_away = (y[-1] - regression_line[-1]) / std_dev

    # Find the true dates for A, B, C, and D
    current_price = y[-1]
    date_d = ticker_data.index[-1]
    date_c = None
    date_b = None
    date_a = None

    # Find Date_C
    for i in range(len(ticker_data) - 2, -1, -1):
        if (y[i] < y[i + 1] and current_price > y[i]) or (y[i] > y[i + 1] and current_price < y[i]):
            date_c = ticker_data.index[i]
            break

    if date_c is not None:
        # Find Date_B
        for i in range(ticker_data.index.get_loc(date_c) - 1, -1, -1):
            if (y[i] < y[i + 1] and current_price > y[i]) or (y[i] > y[i + 1] and current_price < y[i]):
                date_b = ticker_data.index[i]
                break

    if date_b is not None:
        # Find Date_A
        for i in range(ticker_data.index.get_loc(date_b) - 1, -1, -1):
            if (y[i] < y[i + 1] and current_price > y[i]) or (y[i] > y[i + 1] and current_price < y[i]):
                date_a = ticker_data.index[i]
                break

    # Get the relevant price points for measured move calculation
    d = y[-1]
    c = y[ticker_data.index.get_loc(date_c)]
    b = y[ticker_data.index.get_loc(date_b)]
    a = y[ticker_data.index.get_loc(date_a)]

    # Check the measured move criteria
    measured_move = "NO"
    if c < d and c < b and c > a and d < b:
        measured_move = "UP"
    elif c > d and c > b and c < a and d > b:
        measured_move = "DOWN"

    # Create a DataFrame with the calculated data for the current ticker
    ticker_output_data = pd.DataFrame({
        'Ticker': [ticker],
        'Link': [link],
        'Name': [name],
        'Industry': [industry],
        'Market Cap': [market_cap],
        'Standard Deviation': [deviations_away],
        'Measured Move': [measured_move],
        'Date A': [date_a.strftime("%m/%d/%Y") if date_a is not None else None],
        'Price A': [a if date_a is not None else None],
        'Date B': [date_b.strftime("%m/%d/%Y") if date_b is not None else None],
        'Price B': [b if date_b is not None else None],
        'Date C': [date_c.strftime("%m/%d/%Y") if date_c is not None else None],
        'Price C': [c if date_c is not None else None],
        'Current Price': [current_price]
    })

    # Concatenate the current ticker's data with the output DataFrame
    output_data = pd.concat([output_data, ticker_output_data], ignore_index=True)

# Save the output data to list_append.xlsx
output_data.to_excel("list_append.xlsx", index=False)
