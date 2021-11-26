# Due to a change in UI, this program is now deprecated

# Investopedia-Stock-Simulator-Diagnostic
Python program that scrapes data from Investopedia's Stock Simulator and writes to Excel

## Data Scraped from Investopedia
* Ticker
* Quantity Bought
* Purchase Price
* Today's Change (%)
* Overall Change (%)

## Features
* Automatically logs in to Investopedia and navigates to Portfolio
* Scrapes Portfolio, regardless of amount of stocks held
* Stores data and writes it to an excel file (default = stocks.xlsx) with column and row headers
* Writes formulas across a row for adjusted purchase price (QTY * Purchase Price + Commission (Default = 29.95)) and profit (Total Value - Total Purchase Price)
* Conditionally formats profit and percent rows to be highlighted green or red depending on value of cell (can be changed)

## Disclaimers
* Options are not supported, but implementation of options scraping would likely be easy to integrate
* User may also have to tweak code to indicate what the Excel file should be called and where it should be stored
* User must fill out username and password for their own account (Lines 27-29)
```# Log in
driver.find_element_by_xpath(r"//*[@id='username']").send_keys(username)
driver.find_element_by_xpath(r"//*[@id='password']").send_keys(password)
```
![Screenshot of Excel Output](https://github.com/aydenscott/Investopedia-Stock-Simulator-Diagnostic/blob/main/Screenshot%20(1).png)
