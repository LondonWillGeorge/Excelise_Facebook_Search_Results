
from datetime import datetime

from selenium import webdriver

from selenium.webdriver.common.by import By

from selenium.webdriver.support.ui import WebDriverWait

from selenium.webdriver.support import expected_conditions as EC

from selenium.webdriver.common.keys import Keys

import time

import random

FBURL = "https://www.facebook.com"

PASSWD = "supersecretpassword1234"
# maybe easier to use global password all accounts, 
# NB 3 real words often used as more secure mix with digits and capitals

def scrollDown(driver, value):
    driver.execute_script("window.scrollBy(0,"+str(value)+")")

def scrollDownAllTheWay(driver):
    old_page = driver.page_source
    while True:
        print("Scrolling loop")
        for i in range(2):
            scrollDown(driver, 500)
            time.sleep(2) # with original 2 did not get whole page, 10 still doesn't...
        new_page = driver.page_source
        if new_page != old_page:
            old_page = new_page
        else:
            break
    return True

def WaitForLogin(driver):
    try:
        WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.ID,"loginbutton"))) # was u_0_2 inner element works most not all of time
    except KeyboardInterrupt as ki:
        print("keyboard interrupt error, error info here: " + datetime.now() + ki)
    except Exception as ex:
        print("Exception happened at time: " + datetime.now() + " Exception info is: " + ex) # TODO: log ex and the current time to separate Excel errors file
        driver.quit()
        time.sleep(10)
        driver.get(FBURL)
        WaitForLogin(driver)

def setDriver():
    options = webdriver.ChromeOptions()
    options.add_argument("--disable-notifications")
    driver = webdriver.Chrome("C:/Program Files (x86)/Google/Chrome/ChromeDriverForSeleniumPython/ChromeDriver.exe", options = options)
    # driver.minimize_window()
    return driver

def loginFB(driver, recruiter):
    email = driver.find_element_by_id("email") #username form field
    password = driver.find_element_by_id("pass") #password form field
    time.sleep(1)
    email.send_keys(recruiter)
    time.sleep(2) # +2 seconds after typing email address
    password.send_keys(PASSWD) # put password in!
    
    loginButton = driver.find_element_by_xpath(".//label[@id =  'loginbutton']//input")
    time.sleep(4) # + 4 seconds after typing password, before clicking Login button
    loginButton.click() 

# TEST1: cycle 3 logins, write message to 3 logins in this file, don't send, 2 minute + random(30seconds) delay each time
# TEST2 - send the messages and friend them, 5 minute + random 30

recruiterNames = ['Joe Bloggs']

recruiterLogins = ['joebloggs@.com']

candidateLogins = ['https://www.facebook.com/testperson']

def main(recruiter, nurse):
    startTime = time.time()
    print("system time start is " + str(startTime))
    
    driver = setDriver()
    
    driver.get(FBURL)
    
    WaitForLogin(driver)
    
    loginFB(driver, recruiter)
    
    time.sleep(5) # + 5 seconds before getting next webpage
    # 21/2/19 first login in debug mode logged in OK, but fb blocked the login at this point
    # 21/2/19 2nd login logged in then immediate block, not in debug  - both using same login
    # 21/2/19 3rd login logged in then block when navigate to certain page also, NB normal user wouldn't paste in login profile then hit return, so perhaps this picked up as suspicious points by fb
    # Remember fb cannot detect selenium directly, only behaviour patterns different from normal human user

    profileURL = nurse
    
    driver.get(profileURL)
    
    # <a ... href="/messages/t/<fb profile id>/" ...>
    # OR messages/t/<fb profile digits only type id like 12345678901234>
    if 'profile.php' in profileURL:
        msgString = profileURL.split('id=')[1]
    else:
        msgString = profileURL.split('facebook.com/')[1]
    
    messageButton = driver.find_element_by_xpath(".//a[contains(@href, msgString)]")
    messageButton.click()
    
    driver.implicitly_wait(10) # wait 3 or 10? seconds before clicking Add Friend
    
    # try only because often not there
    friendButton = driver.find_element_by_xpath(".//button[contains(@aria-label, 'as a friend')]")
    print('friendButton is ' + str(friendButton))
    # if friendButton:
    # friendButtton.click()
    
    driver.implicitly_wait(4)
    
    messageInner = driver.find_element_by_xpath("//div[contains(@class,'_5rpu') and @role='combobox']")
    
    pageName = driver.find_element_by_xpath(".//span[@data-testid =  'profile_name_in_profile_page']//a").text
    
    # initialise list of 11 phrases for the message NB range needs (0,11)
    fbMsg = []
    for _ in range(0,11):
        fbMsg.append('')
    
    # if name blank or just spaces, this should return firstName as None
    if pageName.split():
        firstName = pageName.split()[0]
    else:
        firstName = None
    
    if firstName is not None:
        fbMsg[0] = "Dear " + firstName + ","
    else:
        fbMsg[0] = "Hiya,"
    # OK so this sentence might have implications, may need to adjust
    fbMsg[1] = "Sorry to message you directly but a colleague of yours gave me your contact details. I didn't have your contact number so I thought I would email you via Facebook."
    fbMsg[2] = "I represent the recruitment team at Shady Contractors Ltd. We're (reasons why better than the competition) ....."
    fbMsg[3] = "We are getting extremely busy and we are needing more people like you to join our team. We recruit people that get recommended to us."
    fbMsg[4] = "At Shady Contractors, it's free to register. It's quick and simple and our streamlined xxx process makes the whole joining experience hassle free for you. Our Pay Rates are extremely high and yyyy."
    fbMsg[5] = "To register you can go on our website and we will then call you to go through the registration process, do click on the link, https://shadycontractors.dprk. Alternatively, you can call us on 0898 xxxxxxx."
    fbMsg[6] = "We have lots of work available, so if you only want x or you would like y, we can offer this to you, so if you might be interested working agency at a high hourly rate of pay, please follow the link above and register with us or alternatively, do call us and I will happily assist you."
    fbMsg[7] = "Thank you for taking the time to read my email."
    fbMsg[8] = recruiterNames[x]
    fbMsg[9] = "Recruiter for zzz"
    fbMsg[10] = "Tel: 0898 zzzzzzz"
    
    for msgInd in range(0,8):
        messageInner.send_keys(fbMsg[msgInd])
        messageInner.send_keys(Keys.SHIFT, Keys.ENTER) # new line
        messageInner.send_keys(Keys.SHIFT, Keys.ENTER) # space line
    
    for msgInd in range(8,11):
        messageInner.send_keys(fbMsg[msgInd])
        messageInner.send_keys(Keys.SHIFT, Keys.ENTER)
    
    # Now send the message: messageInner.send_keys(Keys.ENTER)
    
    endTime = time.time()
    elapsed_time = endTime - startTime
    print("elapsed time is " + str(elapsed_time))
    
    if elapsed_time < 120: # fix to 2 minutes minimum
        switchDelay = (120 - elapsed_time) + random.uniform(0.0, 30.0)
        tenthsDelay = int(10 * switchDelay)/10
        print('Now there will be delay of ' + tenthsDelay + ' seconds waited before window closes and next login loop begins.')
        time.sleep(tenthsDelay)
    
    driver.quit()

# run the main routine cycling through recruiters and nurses
for x in range(0,3): # loops indexes 0 - 2
    main(recruiterLogins[x], candidateLogins[x])
