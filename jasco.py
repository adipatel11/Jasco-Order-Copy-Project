# Decompiled with PyLingual (https://pylingual.io)
# Internal filename: main.py
# Bytecode version: 3.11a7e (3495)
# Source timestamp: 1970-01-01 00:00:00 UTC (0)

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from time import sleep
from datetime import date
from pwinput import pwinput
import writeExcel

# to use explicit waits:

"""

wait = WebDriverWait(driver, 10)
element = wait.until(EC.presence_of_element_located((By.ID, 'element_id')))


"""

PATH = './chromedriver'

s = webdriver.ChromeService(executable_path=PATH)

options = Options()
#options.add_argument("headless=true")
user_agent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36"
options.add_argument(f'user-agent={user_agent}')

driver = webdriver.Chrome(service = s, options=options)
driver.get('https://tap.dor.ms.gov/_/')

driver.set_page_load_timeout(10)  # Set timeout for page load
driver.set_script_timeout(10)     # Set timeout for script execution

def scrollTo(element):
    driver.execute_script('arguments[0].scrollIntoView();', element)

def changePage(action):
    while True:
        try:
            if action == 1:
                next_arrow = driver.find_element(By.ID, 'Dc-v_Vpgnext')
                scrollTo(next_arrow)
                next_arrow.click()
            else:
                prev_arrow = driver.find_element(By.ID, 'Dc-v_Vpgprev')
                scrollTo(prev_arrow)
                prev_arrow.click()
            return None
        except:
            sleep(1)

def screenZoom(percent):
    driver.execute_script(f"document.body.style.zoom='{percent}%'")

def findRowElements():
    ORDER_ID_STRING = 'td.TDS CellText TDCTxt TC-Dc-n FieldDisabled Field TAAuto TAAutoLeft'.replace(' ', '.')
    elems = driver.find_elements(By.CSS_SELECTOR, 'tr.TDR.TDRE') + driver.find_elements(By.CSS_SELECTOR, 'tr.TDR.TDRO')
    startOver = True
    while startOver:
        startOver = False
        for elem in range(len(elems)):
            try:
                elems[elem].find_element(By.CSS_SELECTOR, ORDER_ID_STRING)
            except:
                del elems[elem]
                startOver = True
                break
    return elems

def login():
       
    driver.implicitly_wait(5)
    username_box = driver.find_element(By.ID, 'Dd-5')
    username_box.send_keys(input('Enter your username: '))
    password_box = driver.find_element(By.ID, 'Dd-6')
    password_box.send_keys(pwinput())
    password_box.send_keys(Keys.RETURN)

def passSecurity():
    code = input('Please enter the security code: ')
    success = 0
    while not success:
        try:
            int(code)
            if len(code) == 6:
                success = 1
            else:
                code = input('\nPlease enter a valid 6 digit code: ')
        except:
            code = input('\nPlease enter a valid 6 digit code: ')
    code_box = driver.find_element(By.ID, 'Dc-b')
    code_box.send_keys(code)
    driver.implicitly_wait(5)
    trust_button = driver.find_element(By.ID, 'Dc-d')
    trust_button.click()
    driver.implicitly_wait(5)
    confirm_button = driver.find_element(By.ID, 'action_2')
    confirm_button.click()
    print()

def goToOrders():
    driver.implicitly_wait(5)
    screenZoom(100)
    driver.execute_script('scroll(0, 250)')
    sleep(1)
    retail_order_link = driver.find_element(By.ID, 'cl_Dl-n1-7')
    retail_order_link.click()
    driver.implicitly_wait(5)

def readInv(month, day, year):
    try:
        pages = driver.find_element(By.ID, 'Dc-i_Vpgcurrent').text
        ind = pages.index('of ')
        pages = int(pages[ind + 3:])
    except:
        pages = 1
    currentItem = 1
    for page in range(1, pages + 1):
        if page > 1:
            while True:
                try:
                    next_row = driver.find_element(By.ID, 'Dc-i_Vpgnext')
                    next_row.click()
                except:
                    sleep(1)
        sleep(1)
        while True:
            row = []
            try:
                itemNum = int(driver.find_element(By.ID, f'Dc-8-{currentItem}').text)
                row.append(itemNum)
            except:
                pass
                itemName = driver.find_element(By.ID, f'Dc-9-{currentItem}').text
                row.append(itemName)
                row.append('')
                itemQty = int(driver.find_element(By.ID, f'c_Dc-f-{currentItem}').text)
                row.append(itemQty)
                orderNum = driver.find_element(By.ID, 'caption2_Dc-6').text
                row.append(orderNum[30:])
                months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'AugSep', 'Oct', 'Nov', 'Dec']
                row.append(date(year, months.index(month) + 1, int(day)))
                print(row)
                print()
                writeExcel.writeToFile(row)
                currentItem += 1
    backButton = driver.find_element(By.ID, 'ManagerBackNavigation')
    backButton.click()

def gatherInfo(m, d):
    driver.maximize_window()
    filterBox = driver.find_element(By.ID, 'Dc-z')
    filterBox.send_keys(f"submitted='{d}-{m}' and status='Reserve Inventory'")
    filterBox.send_keys(Keys.RETURN)
    sleep(2)
    pageText = driver.find_element(By.ID, 'Dc-v_Vpgcurrent').text
    ind = pageText.index('of ')
    maxPage = int(pageText[ind + 3:])
    scrollTo(driver.find_element(By.ID, 'Dc-v_Vpgcurrent'))
    changePage(-1)
    for page in range(1, maxPage + 1):
        sleep(2)
        ORDER_ID_STRING = 'td.TDS CellText TDCTxt TC-Dc-n FieldDisabled Field TAAuto TAAutoLeft'.replace(' ', '.')
        DATE_STRING = 'td.TDS CellDate TDCTxt TC-Dc-o FieldDisabled Field TAAuto TAAutoLeft'.replace(' ', '.')
        rowElements = findRowElements()
        orderIds = []
        for row in rowElements:
            try:
                if len(row.find_elements(By.LINK_TEXT, 'View')) > 0:
                    orderIds.append(row.find_element(By.CSS_SELECTOR, ORDER_ID_STRING).text)
            except:
                pass
        for id in orderIds:
            sleep(1)
            rowElements = findRowElements()
            scrollTo(driver.find_element(By.ID, 'Dc-v_Vpgcurrent'))
            for row in rowElements:
                while True:
                    try:
                        currentId = row.find_element(By.CSS_SELECTOR, ORDER_ID_STRING).text
                        y = int(row.find_element(By.CSS_SELECTOR, DATE_STRING).text.split('-')[-1])
                    except Exception as e:
                        sleep(1)
                if currentId == id:
                    button = row.find_element(By.LINK_TEXT, 'View')
                    while True:
                        try:
                            button.click()
                        except Exception as e:
                            sleep(1)
                    readInv(m, d, y)
                    break
        changePage(1)
        sleep(3)
    for i in range(maxPage):
        changePage(-1)

def main():
    login()
    passSecurity()
    goToOrders()
    keepGoing = '1'
    months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
    while keepGoing == '1':
        success = 0
        while not success:
            month = input('Please enter the month for the orders: ').lower().capitalize()
            if month not in months:
                print('Please enter a valid month, for example jan, feb, mar, etc...\n')
            else:
                success = 1
        success = 0
        while not success:
            day = input('Please enter the day for the orders: ')
            try:
                day = int(day)
                if day not in range(1, 32):
                    print('Please enter a number from 1-31\n')
                    break
                day = str(day)
                if len(day) == 1:
                    day = '0' + day
                success = 1
            except:
                print('Please enter a valid day number from 1-31\n')
        gatherInfo(month, day)
        writeExcel.wb.save(writeExcel.order_file_path)
        keepGoing = input('Would you like to enter another date? Enter 1 for yes: ')
    print('Saving file, please do not exit...')
    writeExcel.deleteDuplicates()
    writeExcel.cleanFile()

if __name__ == "__main__":
    main()
    print('File Saved, Safe to Exit Program')
    driver.quit
    print('Driver Quit')