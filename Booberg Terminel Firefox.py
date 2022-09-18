from selenium import webdriver
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import *
import tkinter
from tkinter import *
from pandastable import Table, TableModel
from bs4 import BeautifulSoup
import pandas as pd
import time
import configparser


# -------| SCRAPING
config = configparser.ConfigParser()
config.read('config.ini')
profile_name = config['DEFAULT']['FirefoxProfile']

driver_service = Service('geckodriver.exe')
driver_options = Options()
driver_options.add_argument("--headless")
driver_options.set_preference('profile', profile_name)

driver = webdriver.Firefox(service = driver_service, options = driver_options)

def get_data():
    global is_fy, is_fq, bs_fy, bs_fq, cf_fy, cf_fq

    # formatting
    url_financials = 'https://finance.yahoo.com/quote/{}/financials?p={}'
    driver.get(url_financials.format(ticker, ticker))

    # fluent wait
    wait = WebDriverWait(driver, timeout = 4, poll_frequency = 0.5, ignored_exceptions = [ElementNotVisibleException, ElementNotSelectableException])

    # accepting cookies
    wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div/div/div/div/form/div[2]/div[2]/button')))
    time.sleep(1)
    driver.find_element(By.XPATH, '/html/body/div/div/div/div/form/div[2]/div[2]/button').click()

    # function to expand the table if needed
    def expand_table():
        try:
            wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text() = 'Expand All'")))
            time.sleep(0.7)
            driver.find_element(By.XPATH, "//span[text() = 'Expand All']").click()
        except:
            pass
        time.sleep(0.7)

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


    # deleting cookies and terminating the webdriver
    driver.delete_all_cookies()
    driver.quit()

def export_to_excel():

    # Creating Excel Writer Object from Pandas  
    writer = pd.ExcelWriter(file_path, engine = 'xlsxwriter')

    is_fy.to_excel(writer, index = False, startrow = 1, startcol = 1)
    is_fq.to_excel(writer, index = False, startrow = 1, startcol = 1 + len(is_fy.columns) + 2)
    bs_fy.to_excel(writer, index = False, startrow = 1 + len(is_fy) + 4, startcol = 1)
    bs_fq.to_excel(writer, index = False, startrow = 1 + len(is_fy) + 4, startcol = 1 + len(bs_fy.columns) + 2)
    cf_fy.to_excel(writer, index = False, startrow = 1 + len(is_fy) + len(bs_fy) + 8, startcol = 1)
    cf_fq.to_excel(writer, index = False, startrow = 1 + len(is_fy) + len(bs_fy) + 8, startcol = 1 + len(cf_fy.columns) + 2)
    
    writer.save()


# -------| INTERFACE

win = tkinter.Tk('Booberg Terminel') 
win.title('Booberg Terminel')
win.geometry('600x400+550+270')
win.resizable(width=False, height=False)
win.iconbitmap('assets/monke.ico')
win.configure(bg = 'black')

data_scraped = False

# run button function
def click_run_btn():
    global ticker, data_scraped
    ticker = ent_ticker.get()
    if not ticker:
        return
    get_data()
    data_scraped = True
    click_rt_btn()

def click_rt_btn():
    if not data_scraped:
        return
    var1 = dd1_var.get()
    var2 = dd2_var.get()
    if var1 == "Income Statement":
        if var2 == "Annual":
           pt.model.df = is_fy
        if var2 == "Quarterly":
           pt.model.df = is_fq
    if var1 == "Balance Sheet":
        if var2 == "Annual":
           pt.model.df = bs_fy
        if var2 == "Quarterly":
           pt.model.df = bs_fq
    if var1 == "Cashflow Statement":
        if var2 == "Annual":
           pt.model.df = cf_fy
        if var2 == "Quarterly":
           pt.model.df = cf_fq
    pt.redraw()

def click_export_btn():
    if not data_scraped:
        return
    global file_path, sheet_name
    file_path = ent_file_path.get()
    if not file_path:
        return 
    sheet_name = ent_sheet_name.get()
    export_to_excel()

# enter ticker 
ent_ticker_txt = Label(win, bg = 'black', fg = '#ff9e2b', text = 'Ticker: ')
ent_ticker_txt.place(x = 2, y = 2, height = 23)
ent_ticker = Entry(win, relief = 'flat', background = '#ff9e2b')
ent_ticker.place(x = 4 + ent_ticker_txt.winfo_reqwidth(), y = 2, height = 23) 

# run button
btn_run = Button(win, text = "RUN", relief = 'flat', background = '#ff9e2b', command = click_run_btn)
btn_run.place(x = 6 + ent_ticker_txt.winfo_reqwidth() + ent_ticker.winfo_reqwidth(), y = 2, height = 23)

dd1_choices = {"Income Statement", "Balance Sheet", "Cashflow Statement"}
dd1_var = StringVar()
dd1_var.set("Income Statement")
dd1 = OptionMenu(win, dd1_var, *dd1_choices)
dd1.config(bg = '#ff9e2b', relief = 'flat', borderwidth = 0, highlightthickness = 0)
dd1['menu'].config(background = '#ff9e2b', borderwidth = 0, relief = 'flat')
dd1.place(x = 16 + ent_ticker_txt.winfo_reqwidth() + ent_ticker.winfo_reqwidth() + btn_run.winfo_reqwidth(), y = 2, width = 143, height = 23)

dd2_choices = {"Annual", "Quarterly"}
dd2_var = StringVar()
dd2_var.set("Annual")
dd2 = OptionMenu(win, dd2_var, *dd2_choices)
dd2.config(bg = '#ff9e2b', relief = 'flat', borderwidth = 0, highlightthickness = 0)
dd2['menu'].config(background = '#ff9e2b', borderwidth = 0, relief = 'flat')
dd2.place(x = 18 + ent_ticker_txt.winfo_reqwidth() + ent_ticker.winfo_reqwidth() + btn_run.winfo_reqwidth() + 143, y = 2, width = 86, height = 23)

btn_rt = Button(win, text = "Refresh Table", relief = 'flat', background = '#ff9e2b', command = click_rt_btn)
btn_rt.place(x = 20 + ent_ticker_txt.winfo_reqwidth() + ent_ticker.winfo_reqwidth() + btn_run.winfo_reqwidth() + 143 + 86, y = 2, height = 23)

# enter spreadsheet path
ent_file_path_text = Label(win, bg = 'black', fg = '#ff9e2b', text = 'Spreadsheet path: ')
ent_file_path_text.place(x = 2, y = 27, height = 23) 
ent_file_path = Entry(win, relief = 'flat', background = '#ff9e2b')
ent_file_path.place(x = 4 + ent_file_path_text.winfo_reqwidth(), y = 27, height = 23) 

# enter sheet name 
ent_sheet_name_text = Label(win, bg = 'black', fg = '#ff9e2b', text = 'Sheet name (optional): ')
ent_sheet_name_text.place(x = 6 + ent_file_path_text.winfo_reqwidth() + ent_file_path.winfo_reqwidth(), y = 27, height = 23) 
ent_sheet_name = Entry(win, relief = 'flat', background = '#ff9e2b')
ent_sheet_name.place(x = 8 + ent_file_path_text.winfo_reqwidth() + ent_file_path.winfo_reqwidth() + ent_sheet_name_text.winfo_reqwidth(), y = 27, height = 23) 

# export button
btn_export = Button(win, text = "Export", relief = 'flat', background = '#ff9e2b', command = click_export_btn)
btn_export.place(x = 10 + ent_file_path_text.winfo_reqwidth() + ent_file_path.winfo_reqwidth() + ent_sheet_name_text.winfo_reqwidth() + ent_sheet_name.winfo_reqwidth(), y = 27, height = 23)

frame = tkinter.Frame(win)
frame.place(x = 2, y = 52, width = 600, height = 400 - ent_ticker_txt.winfo_reqheight() - ent_file_path_text.winfo_reqheight() - 10)
pt = Table(frame)
pt.setRowColors(clr = 'black')
pt.show()

win.mainloop()
