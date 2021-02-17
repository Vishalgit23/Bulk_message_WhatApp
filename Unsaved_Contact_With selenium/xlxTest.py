# TO SEND SIMPLE MESSAGE FROM WHATSAPP

from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from time import sleep
import urllib.parse
import xlrd 
driver = None
Link = "https://web.whatsapp.com/"
wait = None

def whatsapp_login():
    global wait, driver, Link
    chrome_options = Options()
    chrome_options.add_argument('--user-data-dir=./User_Data')
    driver = webdriver.Chrome(options=chrome_options)
    wait = WebDriverWait(driver, 20)
    print("SCAN YOUR QR CODE FOR WHATSAPP WEB IF DISPLAYED")
    driver.get(Link)
    driver.maximize_window()
    print("QR CODE SCANNED")




def send_message_to_unsavaed_contact(number,msg,count):
    # Reference : https://faq.whatsapp.com/en/android/26000030/
    print("In send_message_to_unsavaed_contact method")
    params = {'phone': str(number), 'text': str(msg)}
    end = urllib.parse.urlencode(params)
    final_url = Link + 'send?' + end
    print(final_url)
    driver.get(final_url)
    WebDriverWait(driver, 300).until(EC.presence_of_element_located((By.XPATH, '//div[@title = "Menu"]')))
    print("Page loaded successfully.")
    for retry in range(3):
        try:
            sleep(10)
            wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="main"]/footer/div[1]/div[3]/button'))).click()
            break
        except Exception as e:
            print("Fail during click on send button.")
            if retry==2:return
    msg_box = driver.find_element_by_xpath('//*[@id="main"]/footer/div[1]/div[2]/div/div[2]')
    for index in range(count-1):
        msg_box.send_keys(msg)
        driver.find_element_by_xpath('//*[@id="main"]/footer/div[1]/div[3]/button').click()
    print("Message sent successfully.")


if __name__ == "__main__":

    print("Web Page Open")
    # Let us login and Scan
    whatsapp_login()

    loc = r"C:\Users\hp\Desktop\gitPY\code\Cont.xlsx" # Location of excel file 

    # opening workbook 
    wb = xlrd.open_workbook(loc) 
    # selecting the first worksheet
    sheet = wb.sheet_by_index(0) 
    lst = []

    # not including the first row
    for i in range(1,sheet.nrows):
        dummy = []
        for j in range(sheet.ncols):
            # appending the columns with the same row to dummy list
            
            dummy.append(sheet.cell_value(i, j))
            
            
        # appending the dummy list to the main list
        lst.append(dummy)
    for row in lst:
        num=row[0]
        msg=row[1]
        # print(row[0])
        # print(row[1])
        send_message_to_unsavaed_contact(num, msg, 1) # calling Function by passing number ,message,count 
        sleep(5)
    
    sleep(10)
    driver.close() # Close the Open tab
    print("THE END")
    driver.quit()

