############### Imports ##################

from selenium import webdriver
import requests
import time


############### Initiate Chrome drivers #################

options = webdriver.ChromeOptions()
options.add_argument('--ignore-certificate-errors')
options.add_argument("--test-type")
options.binary_location = "/usr/bin/google-chrome"

driver = webdriver.Chrome(chrome_options=options)


############### Get MS 365 login page ###################

driver.get('https://www.office.com/?auth=2')


############### Get organisation login ##################

uid_1 = driver.find_element_by_xpath('//*[@id="i0116"]')
uid_1.send_keys("Enter login ID")

sign_in_1 = driver.find_elements_by_xpath('//*[@id="idSIButton9"]')[0]
sign_in_1.click()

time.sleep(20)


################ Enter login credentials in organisation login page ###################

uid_2 = driver.find_element_by_xpath('//*[@id="username"]')
uid_2.send_keys("Enter login ID")

pswd = driver.find_element_by_xpath('//*[@id="password"]')
pswd.send_keys("Enter Password")

sign_in_2 = driver.find_elements_by_xpath('//*[@id="fm1"]/input[4]')[0]
sign_in_2.click()


################# Disable permanant login ###################

set_no = driver.find_elements_by_xpath('//*[@id="idBtn_Back"]')[0]
set_no.click()


################ Get Form #####################

driver.get('Enter form link')

time.sleep(10)


############## Checking number of responses and the name of the form ################

form_name = driver.find_element_by_xpath('//*[@id="analyzeViewPrintChildContainer"]/div[1]').text

response_num = driver.find_element_by_xpath('//*[@id="analyzeViewPrintChildContainer"]/div[2]/div[1]/div[1]/div[1]/div[1]').text

print("Form name: ", form_name)
print("Responses so far: ", response_num)


################# Download the responses as Excel file #######################

download_button = driver.find_elements_by_xpath('//*[@id="analyzeViewPrintChildContainer"]/div[3]/div[2]/div/button')[0]
download_button.click()

time.sleep(10)

driver.close()
