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

def send_attachment_to_unsavaed_contact(number, file_name):
    print("In send_attachment_to_unsavaed_contact method")
    params = {'phone': str(number)}
    end = urllib.parse.urlencode(params)
    final_url = Link + 'send?' + end
    print(final_url)
    driver.get(final_url)
    WebDriverWait(driver, 300).until(EC.presence_of_element_located((By.XPATH, '//div[@title = "Menu"]')))
    print("Page loaded successfully.")
    for retry in range(3):
        try:
            sleep(10)
            wait.until(EC.presence_of_element_located((By.XPATH, '//div[@title = "Attach"]'))).click()
            break
        except Exception as e:
            print("Fail during click on Attachment button.")
            if retry==2:return
    attachment = driver.find_element_by_xpath('//input[@accept="image/*,video/mp4,video/3gpp,video/quicktime"]')
    print("check")
    attachment.send_keys(file_name)
    print("check1")
    sleep(15)
    print("check")
    send = wait.until(EC.presence_of_element_located((By.XPATH, '//span[@data-icon="send"]')))
    send.click()
    print("File send successfully.")




if __name__ == "__main__":

    print("Web Page Open")
    # Let us login and Scan
    whatsapp_login()

    loc = r"C:\Users\hp\Desktop\gitPY\code\Cont.xlsx"  # Location of  Excel file
 
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
        # msg=row[1]
        # print(row[0])
        # print(row[1])

        send_attachment_to_unsavaed_contact(num , r"C:\Users\hp\Desktop\gitPY\code\img1.mp4") # call function byr passing number ans Path of file 
        sleep(15)
   
    sleep(10)
    driver.close() # Close the Open tab
    print("THE END")
    driver.quit()

