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

# run button function
def printValue():
    global ticker, ss_path, sheet_name, export, get_is, get_bs, get_cf, get_fy, get_fq
    ticker = ent_ticker.get()
    if not ticker:
        return
    get_is = cbvar_is.get()
    get_bs = cbvar_bs.get()
    get_cf = cbvar_cf.get()
    get_fy = cbvar_annual.get()
    get_fq = cbvar_quarterly.get()
    sum = get_is + get_bs + get_cf + get_fy + get_fq 
    if sum == 0:
        return
    export = cbvar_export.get()
    if export == 1:
        ss_path = ent_ss_path.get()
        if not ss_path:
            return 
        sheet_name = ent_sheet_name.get()
    win.destroy()

# title
boober_terminel_title = Label(win, text = 'BOOBERG TERMINEL', font = 'Romantic').grid(row = 0, columnspan = 3)

# stock ticker
ent_ticker_txt= Label(win, text = 'Ticker: ').grid(row = 1, sticky = W) 
ent_ticker = Entry(win)
ent_ticker.grid(row = 1, column = 1) 

# income statement, balance sheet, cashflow statement checkbox
cbvar_is = IntVar() 
cbvar_bs = IntVar() 
cbvar_cf = IntVar() 

cb_is = Checkbutton(win, text = 'Income Statement', variable = cbvar_is, onvalue = 1, offvalue = 0).grid(row = 2, sticky = W) 
cb_bs = Checkbutton(win, text = 'Balance Sheet', variable = cbvar_bs, onvalue = 1, offvalue = 0).grid(row = 2, column = 1, sticky = W) 
cb_cf = Checkbutton(win, text = 'Cashflow Statement', variable = cbvar_cf, onvalue = 1, offvalue = 0).grid(row = 2, column = 2, sticky = W) 

# annual, quarterly checkbox
cbvar_annual = IntVar()
cbvar_quarterly = IntVar() 

cb_annual = Checkbutton(win, text = 'Annual', variable = cbvar_annual, onvalue = 1, offvalue = 0).grid(row = 3, sticky = W) 
cb_quarterly = Checkbutton(win, text = 'Quarterly', variable = cbvar_quarterly, onvalue = 1, offvalue = 0).grid(row = 3, column = 1, sticky = W) 

# excel
cbvar_export = IntVar()

cb_export = Checkbutton(win, text = 'Export to Excel', variable = cbvar_export, onvalue = 1, offvalue = 0).grid(row = 6, sticky = W) 

ent_ss_path_text = Label(win, text = 'Spreadsheet path: ').grid(row = 7, sticky = W) 
ent_ss_path = Entry(win)
ent_ss_path.grid(row = 7, column = 1) 

ent_sheet_name_text = Label(win, text = 'Sheet name (optional): ').grid(row = 8, sticky = W) 
ent_sheet_name = Entry(win)
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

if get_is == 1:

    expand_table()

    # annual income statement
    if get_fy == 1:
        is_fy = get_table()


    # quarterly income statement
    if get_fq == 1:
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

if get_bs == 1:

    # click balance sheet
    try:
        wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="Col1-1-Financials-Proxy"]/section/div[1]/div[1]/div/div[2]/a/div/span')))
        time.sleep(0.7)
        driver.find_element(By.XPATH, '//*[@id="Col1-1-Financials-Proxy"]/section/div[1]/div[1]/div/div[2]/a/div/span').click()
    except:
        print("\n\nAn error occurred!\n")
    time.sleep(2)

    expand_table()

    # annual balance sheet
    if get_fy == 1:
        bs_fy = get_table()


    # quarterly balance sheet
    if get_fq == 1:
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

if get_cf == 1:

    # click cash flow
    try:
        wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="Col1-1-Financials-Proxy"]/section/div[1]/div[1]/div/div[3]/a/div/span')))
        time.sleep(0.7)
        driver.find_element(By.XPATH, '//*[@id="Col1-1-Financials-Proxy"]/section/div[1]/div[1]/div/div[3]/a/div/span').click()
    except:
        print("\n\nAn error occurred!\n")
    time.sleep(2)

    expand_table()

    # annual cash flow
    if get_fy == 1:
        cf_fy = get_table()

    # quarterly cash flow
    if get_fq == 1:
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
if get_is:
    if get_fy:
        print("\n\n\nINCOME STATEMENT (ANNUAL)\n")
        print(is_fy)
    if get_fq:
        print("\n\n\nINCOME STATEMENT (QUARTERLY)\n")
        print(is_fq)

if get_bs:
    if get_fy:
        print("\n\n\nBALANCE SHEET (ANNUAL)\n")
        print(bs_fy)
    if get_fq:
        print("\n\n\nBALANCE SHEET (QUARTERLY)\n")
        print(bs_fq)

if get_cf:
    if get_fy:
        print("\n\n\nCASH FLOW STATEMENT (ANNUAL)\n")
        print(cf_fy)
    if get_fq: 
        print("\n\n\nCASH FLOW STATEMENT (QUARTERLY)\n")
        print(cf_fq)




if export == 1:

    # Creating Excel Writer Object from Pandas  
    writer = pd.ExcelWriter(ss_path, engine = 'xlsxwriter')

    if get_is:
        if get_fy:
            is_fy.to_excel(writer, startrow = 1, startcol = 1)
        if get_fq:
            is_fq.to_excel(writer, startrow = 1 + len(is_fy) + 4, startcol = 1)

    if get_bs:
        if get_fy:
            bs_fy.to_excel(writer, startrow = 1 + len(is_fy) + len(is_fq) + 8, startcol = 1)
        if get_fq:
            bs_fq.to_excel(writer, startrow = 1 + len(is_fy) + len(is_fq) + len(bs_fy) + 12, startcol = 1)

    if get_cf:
        if get_fy:
            cf_fy.to_excel(writer, startrow = 1 + len(is_fy) + len(is_fq) + len(bs_fy) + len(bs_fq) + 16, startcol = 1)
        if get_fq: 
            cf_fq.to_excel(writer, startrow = 1 + len(is_fy) + len(is_fq) + len(bs_fy) + len(bs_fq) + len(cf_fy) + 20, startcol = 1)

    writer.save()

# -------| DRIVER QUIT

# deleting cookies and terminating the webdriver
driver.delete_all_cookies()
driver.quit()

# -------| 

input()
