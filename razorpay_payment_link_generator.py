from selenium import webdriver
import time
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from selenium.webdriver.chrome.options import Options

#To Run the Program headless without opening any physical window

chrome_options = Options()
chrome_options.add_argument("--headless")
chrome_options.add_argument("--window-size=1920x1080")
chrome_options.binary_location = 'C:\link\Canary\Chrome SxS\Application\chrome.exe'

#-----------------------------#----------------------------------

#Open The chrome and login using the credentials and opening the website to generate link

driver = webdriver.Chrome(options=chrome_options, executable_path='C:\link\chromedriver')
#driver = webdriver.Chrome('C:\Programming\chromedriver')
driver.implicitly_wait(100)
driver.get("https://prodsupport.zestmoney.in")
driver.find_element_by_id("email-id").send_keys("##########") # user name area, enter the username to login
driver.find_element_by_id("pwd").send_keys("########") # password area, enter the password to login
driver.find_element_by_class_name("jss175").click()
time.sleep(5)
driver.get("https://prodsupport.zestmoney.in/dashboard/PaymentLink")

#-----------------------------------#------------------------------------

#authentication to Login and access google sheet and open and read the Sheet

scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('client_secret.json', scope) # json file with credentials to be accessed hence place in same folder.
client = gspread.authorize(creds)

sheet = client.open("Link Generating sheet").sheet1 # sheet name which is created in google sheet which has data
#----------------------------------------#--------------------------------
#To check the count of the Mambu id(Loan Id) to control the Loop

z = int(sheet.cell(1,5).value) #at this place in the sheet there is COUNTA function to count.

#Initialize
a = 2

#for loop to check and tell the program to stop

while a <= z:
    #if loop to check the id are already generated?

    if sheet.cell(a,4).value != 'Done':
        #Getting Mambu id and putting in the website.
        MAMBU_ID = sheet.cell(a, 1).value
        driver.find_element_by_xpath("/html/body/div[1]/div/div/div[2]/div/div/div[1]/div/div/div/div/div/div/div[1]/div/input").clear()
        search_input1 = driver.find_element_by_xpath("/html/body/div[1]/div/div/div[2]/div/div/div[1]/div/div/div/div/div/div/div[1]/div/input")
        search_input1.send_keys(MAMBU_ID)

        # Getting Amount and putting in the website.
        AMOUNT = sheet.cell(a, 2).value
        driver.find_element_by_xpath("/html/body/div[1]/div/div/div[2]/div/div/div[1]/div/div/div/div/div/div/div[2]/div/input").clear()
        search_input2 = driver.find_element_by_xpath("/html/body/div[1]/div/div/div[2]/div/div/div[1]/div/div/div/div/div/div/div[2]/div/input")
        search_input2.send_keys(AMOUNT)

        # click to generate Link
        driver.find_element_by_xpath("/html/body/div[1]/div/div/div[2]/div/div/div[1]/div/div/div/div/div/div/div[3]/div/div/select/option[3]").click()
        driver.find_element_by_xpath("/html/body/div[1]/div/div/div[2]/div/div/div[1]/div/div/div/div/div/div/div[4]/button/span[1]").click()

        #wait for the Link to generate just in case poor connection
        time.sleep(5)

        #check the Link is generated or its a error
        suxes = driver.find_element_by_xpath("/html/body/div[4]/div[2]/div/div[1]/h6").text
        err = driver.find_element_by_xpath("/html/body/div[4]/div[2]/div/div[1]/h6").text

        if suxes == 'SUCCESS':
            link = driver.find_element_by_xpath("/html/body/div[4]/div[2]/div/div[2]/div[2]").text

        elif err == 'Error':
            link = driver.find_element_by_xpath("/html/body/div[4]/div[2]/div/div[2]/div").text

        #---------------------------------#----------------------------------------
        #capture ss to know what happened in the end
        #driver.get_screenshot_as_file("capture.png") # if required use to see to know the status as the process runs in background.
        #Print to check the status activate if required
        #print(link)

        #update the sheet with the link
        sheet.update_cell(a, 3, link)

        #update the sheet with completion status
        sheet.update_cell(a,4,'Done')

        #click cancel button
        driver.find_element_by_xpath("/html/body/div[4]/div[2]/div/div[3]/button").click()

        #iterate
        a += 1
    else:
        a += 1


driver.quit()
