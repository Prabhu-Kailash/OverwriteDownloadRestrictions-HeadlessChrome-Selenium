from selenium import webdriver
import time
from selenium.webdriver.common.keys import Keys
from selenium.webdriver import ActionChains
from selenium.webdriver.chrome.options import Options
import os
import pandas
import win32com.client
from datetime import datetime, timedelta
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities


def CHGDownloader():
    option = Options()
    option.add_experimental_option("prefs", {
        "download.default_directory": "C:\\path\\to\\Downloads",
        "download.prompt_for_download": False,
        'download.directory_upgrade': True,
        'safebrowsing.enabled': False,
        'safebrowsing.disable_download_protection': True,
    })
    option.add_argument("--disable-extensions")
    option.add_argument("--proxy-server=direct://")
    option.add_argument("--proxy-bypass-list=*")
    option.add_argument("--start-maximized")
    option.add_argument("--headless")
    option.add_argument("--disable-gpu")
    option.add_argument("--allow-insecure-localhost")
    option.add_argument("--disable-dev-shm-usage")
    option.add_argument("--no-sandbox")
    option.add_argument("--ignore-certificate-errors")

    capabilities = DesiredCapabilities.CHROME.copy()
    capabilities['acceptSslCerts'] = True
    capabilities['acceptInsecureCerts'] = True
    driver = webdriver.Chrome(
        executable_path="\\path\\to\\chromedriver.exe", options=option, desired_capabilities=capabilities)
    try:
        print("Initializing")
        driver.command_executor._commands["send_command"] = ("POST", '/session/$sessionId/chromium/send_command')
        params = {'cmd': 'Page.setDownloadBehavior',
                  'params': {'behavior': 'allow', 'downloadPath': "C:\\path\\to\\Downloads"}}
        driver.execute("send_command", params)
        driver.implicitly_wait(10)
        driver.get("https://example.com") # site 
        time.sleep(10)
        driver.find_element_by_xpath("//*[@id='favorites_tab']").click()
        time.sleep(2)
        driver.find_element_by_link_text("Reports - View / Run").click()
        time.sleep(5)
        driver.switch_to.frame("gsft_main")
        driver.find_element_by_xpath("//*[@id='li_group_reports_tab']/span").click()
        time.sleep(2)
        driver.find_element_by_link_text("DailyCR_Report_DM").click()
        time.sleep(10)
        element = driver.find_element_by_xpath("/html/body/div[2]/table/tbody/tr/td/div[1]/div/span/div/div[2]/table/tbody/tr/td/div/table/thead/tr/th[3]/span/a")
        actionChains = ActionChains(driver)
        actionChains.context_click(element).perform()
        ActionChains(driver).send_keys(Keys.ARROW_DOWN).perform()
        ActionChains(driver).send_keys(Keys.ARROW_DOWN).perform()
        ActionChains(driver).send_keys(Keys.ARROW_DOWN).perform()
        ActionChains(driver).send_keys(Keys.ARROW_DOWN).perform()
        ActionChains(driver).send_keys(Keys.ARROW_DOWN).perform()
        ActionChains(driver).send_keys(Keys.ARROW_DOWN).perform()
        ActionChains(driver).send_keys(Keys.ENTER).perform()
        ActionChains(driver).send_keys(Keys.ENTER).perform()
        time.sleep(5)
        driver.find_element_by_xpath("//*[@id='download_button']").click()
        time.sleep(2)
        driver.quit()
    except:
        driver.quit()
        print("Re-initializing")
        CHGDownloader()

def emailComplete():
    outlook = win32com.client.Dispatch('Outlook.Application')
    mail_item = outlook.CreateItem(0)
    mail_item.To = 'Email Address'
    mail_item.Cc = '%s' % OnCallName
    mail_item.Subject = "Change Report (%s)" %ReqD
    mail_item.Attachments.Add(Source='C:\\path\\to\\donwloads\\change_request.xlsx')
    body = "<p>Hi Team,</p><p>Please find the changes scheduled today below,</p>"+DF+"<br>%s" %HTML
    mail_item.HTMLBody = (body)
    mail_item.Send()

def emailEmpty():
    outlook = win32com.client.Dispatch('Outlook.Application')
    mail_item = outlook.CreateItem(0)
    mail_item.To = 'Email Address'
    mail_item.Cc = '%s' % OnCallName
    mail_item.Subject = "Change Report (%s)" % ReqD
    mail_item.Attachments.Add(Source='C:\\path\\to\\donwloads\\change_request.xlsx')
    body = "<p>Hi Team,</p><p>No changes are scheduled for today.</p><br>%s" % HTML
    mail_item.HTMLBody = (body)
    mail_item.Send()


CHGDownloader()
os.chdir("C:\\path\\to\\Downloads")
Sheet = pandas.read_excel("change_request.xlsx", sheet_name=0)
Sheet = Sheet[Sheet["State"] != "Canceled"]
Sheet = Sheet.drop(columns=["Assigned to"])
Sheet = Sheet[Sheet["Short Description / Title"].str.contains(r"SFG|B2B|sfg|b2b")]
DF = Sheet.to_html(index=False)
ReqD = datetime.strftime(datetime.now(), '%m/%d/%y')
Sign = open("C:\\path\\to\\KP Sign.html")
HTML = Sign.read()
OnCall = pandas.read_excel("\\path\\to\\DM_B2B_On_Call.xlsx", sheet_name=0)
OnCallEmail = OnCall['Email'][OnCall[OnCall['Date'] == ReqD].index.tolist()].tolist()
OnCallName = str(OnCallEmail[0])
if Sheet.empty:
    emailEmpty()
else:
    emailComplete()
time.sleep(3)
try:
    os.system("taskkill /f /im  outlook.exe")
except:
    print("No PID found")
os.remove("change_request.xlsx")

