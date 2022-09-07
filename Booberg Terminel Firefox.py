from selenium import webdriver
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import *
import tkinter
from tkinter import *
from bs4 import BeautifulSoup
import pandas as pd
import time


# -------| WEBDRIVER

my_service = Service("webdriver/geckodriver.exe")
my_options = Options()
my_options.add_argument("--headless")
my_options.set_preference('profile', "webdriver/profiles/cw5mp8a1.BoobergTerminel")
driver = webdriver.Firefox(service = my_service, options = my_options)


# -------| WINDOW

win = tkinter.Tk('Booberg Terminel') 
win.title('Booberg Terminel')
win.geometry('440x330+650+250')
win.iconbitmap('assets/monke.ico')
win.configure(bg = 'black')

# run button function
def printValue():
    global ticker, ss_path, sheet_name, export
    ticker = ent_ticker.get()
    if not ticker:
        return
    export = cbvar_export.get()
    if export == 1:
        ss_path = ent_ss_path.get()
        if not ss_path:
            return 
        sheet_name = ent_sheet_name.get()
    win.destroy()

# stock ticker
ent_ticker_txt = Label(win, bg = 'black', fg = 'orange', text = 'Ticker: ').grid(row = 1, sticky = W)

ent_ticker = Entry(win, background = 'orange')
ent_ticker.grid(row = 1, column = 1) 

# excel
cbvar_export = IntVar()

cb_export = Checkbutton(win, bg = 'black', fg = 'orange', text = 'Export to Excel', variable = cbvar_export, onvalue = 1, offvalue = 0).grid(row = 6, sticky = W) 

ent_ss_path_text = Label(win, bg = 'black', fg = 'orange', text = 'Spreadsheet path: ').grid(row = 7, sticky = W) 
ent_ss_path = Entry(win, background = 'orange')
ent_ss_path.grid(row = 7, column = 1) 

ent_sheet_name_text = Label(win, bg = 'black', fg = 'orange', text = 'Sheet name (optional): ').grid(row = 8, sticky = W) 
ent_sheet_name = Entry(win, background = 'orange')
ent_sheet_name.grid(row = 8, column = 1) 

# run button
btn_run = Button(win, text = "RUN!!", command = printValue).grid(row = 9, sticky = W)

win.mainloop()  

# -------| GET AND FORMAT DATA

# formatting
url_financials = "https://finance.yahoo.com/quote/{}/financials?p={}"
driver.get(url_financials.format(ticker, ticker))

# fluent wait
wait = WebDriverWait(driver, timeout = 4, poll_frequency = 0.5, ignored_exceptions = [ElementNotVisibleException, ElementNotSelectableException])

# accepting cookies
wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div/div/div/div/form/div[2]/div[2]/button')))
time.sleep(1)
driver.find_element(By.XPATH, '/html/body/div/div/div/div/form/div[2]/div[2]/button').click()

# function to extract the data from the html code and format it in to a table
def get_table():
    html = driver.execute_script('return document.body.innerHTML;')
    soup = BeautifulSoup(html,'lxml')
    
    features = soup.find_all('div', class_ = 'D(tbr)')

    headers = []
    temp_list = []
    final = []
    index = 0

    # create headers
    for item in features[0].find_all('div', class_ = 'D(ib)'):
        headers.append(item.text)

    # statement contents
    while index <= len(features)-1:

        # filter for each line of the statement
        temp = features[index].find_all('div', class_ = 'D(tbc)')
        for line in temp:

            # each item adding to a temporary list
            temp_list.append(line.text)

        # temp_list added to final list
        final.append(temp_list)

        # clear temp_list
        temp_list = []
        index += 1
        
    cool_table = pd.DataFrame(final[1:])
    cool_table.columns = headers
    return cool_table

# function to expand the table if needed
def expand_table():
    try:
        wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text() = 'Expand All']")))
        time.sleep(0.7)
        driver.find_element(By.XPATH, "//span[text() = 'Expand All']").click()
    except:
        pass
    time.sleep(0.7)


# -------| INCOME STATEMENT

expand_table()

is_fy = get_table()

# click quarterly
try:
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="Col1-1-Financials-Proxy"]/section/div[1]/div[2]/button/div/span')))
    time.sleep(0.7)
    driver.find_element(By.XPATH, '//*[@id="Col1-1-Financials-Proxy"]/section/div[1]/div[2]/button/div/span').click()
except:
    print("\n\nAn error occurred!\n")
time.sleep(2)

is_fq = get_table()


# -------| BALANCE SHEET

# click balance sheet
try:
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="Col1-1-Financials-Proxy"]/section/div[1]/div[1]/div/div[2]/a/div/span')))
    time.sleep(0.7)
    driver.find_element(By.XPATH, '//*[@id="Col1-1-Financials-Proxy"]/section/div[1]/div[1]/div/div[2]/a/div/span').click()
except:
    print("\n\nAn error occurred!\n")
time.sleep(2)

expand_table()

bs_fy = get_table()

# click quarterly
try:
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="Col1-1-Financials-Proxy"]/section/div[1]/div[2]/button/div/span')))
    time.sleep(0.7)
    driver.find_element(By.XPATH, '//*[@id="Col1-1-Financials-Proxy"]/section/div[1]/div[2]/button/div/span').click()
except:
    print("\n\nAn error occurred!\n")
time.sleep(2)

bs_fq = get_table()


# -------| CASHFLOW STATEMENT

# click cash flow
try:
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="Col1-1-Financials-Proxy"]/section/div[1]/div[1]/div/div[3]/a/div/span')))
    time.sleep(0.7)
    driver.find_element(By.XPATH, '//*[@id="Col1-1-Financials-Proxy"]/section/div[1]/div[1]/div/div[3]/a/div/span').click()
except:
    print("\n\nAn error occurred!\n")
time.sleep(2)

expand_table()

cf_fy = get_table()

# click quarterly
try:
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="Col1-1-Financials-Proxy"]/section/div[1]/div[2]/button/div/span')))
    time.sleep(0.7)
    driver.find_element(By.XPATH, '//*[@id="Col1-1-Financials-Proxy"]/section/div[1]/div[2]/button/div/span').click()
except:
    print("\n\nAn error occurred!\n")
time.sleep(2)

cf_fq = get_table()


# -------| OUTPUT

print("\n\n\nINCOME STATEMENT (ANNUAL)\n")
print(is_fy)

print("\n\n\nINCOME STATEMENT (QUARTERLY)\n")
print(is_fq)

print("\n\n\nBALANCE SHEET (ANNUAL)\n")
print(bs_fy)

print("\n\n\nBALANCE SHEET (QUARTERLY)\n")
print(bs_fq)

print("\n\n\nCASH FLOW STATEMENT (ANNUAL)\n")
print(cf_fy)

print("\n\n\nCASH FLOW STATEMENT (QUARTERLY)\n")
print(cf_fq)




if export == 1:

    # Creating Excel Writer Object from Pandas  
    writer = pd.ExcelWriter(ss_path, engine = 'xlsxwriter')

    is_fy.to_excel(writer, startrow = 1, startcol = 1)
    is_fq.to_excel(writer, startrow = 1, startcol = 1 + len(is_fy.columns) + 2)
    bs_fy.to_excel(writer, startrow = 1 + len(is_fy) + 4, startcol = 1)
    bs_fq.to_excel(writer, startrow = 1 + len(is_fy) + 4, startcol = 1 + len(bs_fy.columns) + 2)
    cf_fy.to_excel(writer, startrow = 1 + len(is_fy) + len(bs_fy) + 8, startcol = 1)
    cf_fq.to_excel(writer, startrow = 1 + len(is_fy) + len(bs_fy) + 8, startcol = 1 + len(cf_fy.columns) + 2)
    
    writer.save()
    
# -------| DRIVER QUIT

# deleting cookies and terminating the webdriver
driver.delete_all_cookies()
driver.quit()
