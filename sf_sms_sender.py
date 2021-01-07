from selenium import webdriver
import time
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from selenium.webdriver.chrome.options import Options


def display_check(display):
    if display[0] == 'y':
        print(" Launching chrome")
        return True
    elif display[0] == "n":
        print("All operations Will be done in background")
        return True
    else:
        return False

while True:
    try:
        display = input('Do You want to see the display of sending messages type Yes or No ').lower()
        if display_check(display): break
    except ValueError:
        print("Sorry, Wrong Input Try Again")


def sms_check_value(sms_check):
    if sms_check == 1:
        print('Message from the Message Box will be Sent')
        return True
    elif sms_check == 2:
        print('Custom message from Message Column will be sent')
        return True
    else:
        return False


while True:
    try:
        sms_check = int(input('Send messages from (Message box Press 1) or (Press 2 for Sending Custom message) from C column '))
        if sms_check_value(sms_check): break
    except ValueError:
        print("Sorry, Wrong Input Try Again")

#To Run the Program headless without opening any physical window
if display[0] == 'n':
    chrome_options = Options()
    chrome_options.add_argument("--headless")
    chrome_options.add_argument("--window-size=1920x1080")
    chrome_options.binary_location = 'C:\link\Link_gen_lav\Canary\Chrome SxS\Application\chrome.exe'
    driver = webdriver.Chrome(options=chrome_options, executable_path='C:\Codes_projects_python\SMS Sender\chromedrivercanary.exe')
elif display[0] =='y':
    driver = webdriver.Chrome('C:\Codes_projects_python\SMS Sender\chromedriver')
#-----------------------------#----------------------------------

#authentication to Login and access google sheet and open and read the Sheet

scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('client_secret.json', scope)
client = gspread.authorize(creds)
sheet = client.open('SMS').sheet1


#Open The chrome and login using the credentials and opening the website to generate link



driver.get("https://zm.force.com/ZMCollectionteam/login")
driver.implicitly_wait(100)
driver.find_element_by_id("username").send_keys("simran@zestmoney.in")
driver.find_element_by_id("password").send_keys("welcome@1")
driver.find_element_by_id("Login").click()
time.sleep(5)
driver.get("https://zm.force.com/ZMCollectionteam/apex/SMSUsingGupshup_1?id=a000K0000223jLcQAI&ObjectType=Customer")

#----------------------------------------#--------------------------------
#To check the count of the Names to control the Loop
z = int(sheet.cell(1,5).value)

#Initialize
a = 2

#for loop to check and tell the program to stop
while a <= z:
    # if loop to check the id are already generated?

    if sheet.cell(a, 4).value != 'Sent':

        # Name Area clearing and entering data

        #Name = sheet.cell(a,1).value
        driver.find_element_by_xpath("/html/body/form/div[1]/div/div/div/div[2]/div/div[1]/input").clear()
        #search_input1 = driver.find_element_by_xpath("/html/body/form/div[1]/div/div/div/div[2]/div/div[1]/input")
        #search_input1.send_keys(Name)

        # Mobile Number Area clearing and entering data
        Mobile = sheet.cell(a,2).value
        driver.find_element_by_xpath("/html/body/form/div[1]/div/div/div/div[2]/div/div[2]/input").clear()
        search_input2 = driver.find_element_by_xpath("/html/body/form/div[1]/div/div/div/div[2]/div/div[2]/input")
        search_input2.send_keys(Mobile)

        # to check and send SMS # Message Area clearing and entering data

        if sms_check == 1:
            Message = sheet.cell(2, 10).value
            driver.find_element_by_name("j_id0:frmid:sms1").clear()
            search_input3 = driver.find_element_by_name("j_id0:frmid:sms1")
            search_input3.send_keys(Message)

        elif sms_check == 2:
            Message = sheet.cell(a, 3).value
            driver.find_element_by_name("j_id0:frmid:sms1").clear()
            search_input3 = driver.find_element_by_name("j_id0:frmid:sms1")
            search_input3.send_keys(Message)




        #Click Send button
        driver.find_element_by_xpath("/html/body/form/div[1]/div/div/div/div[2]/div/div[5]/input").click()

        #update in sheet as message Status sent
        sheet.update_cell(a, 4, 'Sent')
        time.sleep(1)
        a+=1

    else:
        a += 1

driver.quit()