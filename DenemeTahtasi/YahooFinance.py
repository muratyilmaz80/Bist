import yfinance as yf

ticker = yf.Ticker('EREGL').info

market_price = ticker['regularMarketPrice']

previous_close_price = ticker['regularMarketPreviousClose']

print('Ticker: EREGL')

print('Market Price:', market_price)

print('Previous Close Price:', previous_close_price)