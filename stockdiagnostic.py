import xlsxwriter, re, os, sys
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from xlsxwriter.utility import xl_rowcol_to_cell

# Delete file path to ensure xlsxwriter doesn't raise exception
filename = r'stocks.xlsx'
try:
    os.remove(filename)
except FileNotFoundError: # This just means the file was not found, which is fine for xlsx (since it can't overwrite)
    pass
except PermissionError: # Unable to find a way to actively close an open excel file so requires user to
    print("Please close out of excel before running the program and try again")
    sys.exit()
# Set up browser
source = r"https://www.investopedia.com/simulator/portfolio/"
path_to_extension = r'C:\Users\tscot\OneDrive\Documents\OneDrive\Desktop\1.35.2_0' # Loads adblocker for selenium
chrome_options = Options()
chrome_options.add_argument('load-extension=' + path_to_extension)
driver = webdriver.Chrome(options=chrome_options)
driver.create_options()
driver.get(source)
driver.maximize_window()
driver.implicitly_wait(5)

# Log in
username = os.environ.get('invest_username')
password = os.environ.get('invest_password')
driver.find_element_by_xpath(r"//*[@id='username']").send_keys(username)
driver.find_element_by_xpath(r"//*[@id='password']").send_keys(password)
driver.find_element_by_xpath(r"//*[@id='login']").click()

# Navigate to portfolio
driver.find_element_by_xpath(r"//*[@id='Content']/div[1]/ul/li[2]/a").click()

# Set up data structures
Tickers = []
Quantity = []
Purchase = []
T_Value = []
DayChangeList = []
OverallChangeList = []
parsed_day = []
parsed_overall = []
TempPurchase = []
TempT_Value = []
num_stocks = len(driver.find_elements_by_class_name("expandable"))
tags = driver.find_elements_by_tag_name("td")


# Retrieve all tickers
def get_tickers(num_stocks):
    for ticker_start in range(2, len(tags), 11):
        if num_stocks == 0:
            break
        ticker = tags[ticker_start]
        Tickers.append(ticker.text)
        num_stocks -= 1


# Retrieve all QTY
def get_qty(num_stocks):
    for qty_start in range(4, len(tags), 11):
        if num_stocks == 0:
            break
        qty = tags[qty_start]
        Quantity.append(qty.text)
        num_stocks -= 1


# Retrieve all purchase prices
def get_purchase(num_stocks):
    for pur_start in range(5, len(tags), 11):
        if num_stocks == 0:
            break
        pur = tags[pur_start]
        TempPurchase.append(pur.text)
        num_stocks -= 1
    for term in TempPurchase:
        newterm = term.replace('$', '').strip()
        Purchase.append(newterm)


# Retrieve all total values
def get_totals(num_stocks):
    for tvalue in range(7, len(tags), 11):
        if num_stocks == 0:
            break
        value = tags[tvalue]
        TempT_Value.append(value.text)
        num_stocks -= 1
    for term in TempT_Value:
        newterm = term.replace('$', '').strip()
        T_Value.append(newterm)


# Retrieve all percent day changes
def get_daychanges(num_stocks):
    for daychange in range(8, len(tags), 11):
        if num_stocks == 0:
            break
        currentdaychange = tags[daychange]
        parsed_day.append(currentdaychange.text)
        num_stocks -= 1


# Retrieve all percent overall changes
def get_overallchanges(num_stocks):
    for overallchange in range(9, len(tags), 11):
        if num_stocks == 0:
            break
        currentoverall = tags[overallchange]
        parsed_overall.append(currentoverall.text)
        num_stocks -= 1


def run():
    get_tickers(num_stocks)
    get_qty(num_stocks)
    get_purchase(num_stocks)
    get_totals(num_stocks)
    get_daychanges(num_stocks)
    get_overallchanges(num_stocks)


run()

# Use regular expressions to find percents
percents = re.compile(r"(-*\d?\d?\d.\d\d %)")
for i in parsed_day: # parsed lists are found in get_daychanges and get_overallchanges functions
    parsed_daycount = percents.findall(i)
    for x in parsed_daycount:
        new = x.replace('%', '').strip()
        DayChangeList.append(new)

for i in parsed_overall:
    parsed_overallcount = percents.findall(i)
    for x in parsed_overallcount:
        new = x.replace('%', '').strip()
        OverallChangeList.append(new)

# Write to xlsx
workbook = xlsxwriter.Workbook(filename, {'strings_to_numbers': True})
worksheet = workbook.add_worksheet()
green_format = workbook.add_format({'bg_color': '#C6EFCE',
                                    'font_color': '#006100'})
red_format = workbook.add_format({'bg_color': '#FFC7CE',
                                  'font_color': '#9C0006'})
col = 1
row = 0

# Write data to respective rows
worksheet.write_row(row, col, Tickers)
worksheet.write_row(row + 1, col, Quantity)
worksheet.write_row(row + 2, col, Purchase)
worksheet.write_row(row + 3, col, T_Value)
worksheet.write_row(row + 6, col, DayChangeList)
worksheet.write_row(row + 7, col, OverallChangeList)

# Write formulas to entire row
for col_reference in range(1, num_stocks + 1):
    # Writes adjusted price formula
    qty_cell = xl_rowcol_to_cell(1, col_reference)
    purchaseprice_cell = xl_rowcol_to_cell(2, col_reference)
    commission = '29.95'
    worksheet.write_formula(4, col_reference, '=%s * %s + %s' % (qty_cell, purchaseprice_cell, commission))
    # Writes profit formula
    totalvalue_cell = xl_rowcol_to_cell(3, col_reference)
    adjprice_cell = xl_rowcol_to_cell(4, col_reference)
    worksheet.write_formula(5, col_reference, '=%s - %s' % (totalvalue_cell, adjprice_cell))
    col_reference += 1

worksheet.write(1, 0, "QTY")
worksheet.write(2, 0, "Purchase Price")
worksheet.write(3, 0, "Total Value")
worksheet.write(4, 0, "Adjusted Price")
worksheet.write(5, 0, "Profit")
worksheet.write(6, 0, "Today's Change")
worksheet.write(7, 0, "Overall Change")
worksheet.set_column(0, 0, 14)

# Color profit cells
worksheet.conditional_format(5, 1, 5, num_stocks, {
    'type': 'cell',
    'criteria': 'less than',
    'value': 15,
    'format': red_format
})
worksheet.conditional_format(5, 1, 5, num_stocks, {
    'type': 'cell',
    'criteria': 'greater than',
    'value': 15,
    'format': green_format
})

# Color percent changes
worksheet.conditional_format(6, 1, 7, num_stocks, {
    'type': 'cell',
    'criteria': 'greater than',
    'value': 1.5,
    'format': green_format
})
worksheet.conditional_format(6, 1, 7, num_stocks, {
    'type': 'cell',
    'criteria': 'less than',
    'value': -1.5,
    'format': red_format
})

driver.quit()
workbook.close()
os.startfile(filename)
