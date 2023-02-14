import os
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager

class Browser:
    def openBrowse(self):
        driver.get(baseUrl)
        print(driver.title)


chrome_options = webdriver.ChromeOptions()
chrome_options.binary_location = os.environ.get("GOOGLE_CHROME_BIN")
chrome_options.add_argument("--headless")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--no-sandbox")
driver = webdriver.Chrome((ChromeDriverManager().install()), options=chrome_options)

driver.maximize_window()


baseUrl = "https://www.dsebd.org/"
if (baseUrl[0:8] == "https://"):
    baseurl = baseUrl
else:
    baseurl = "https://"+baseUrl

object = Browser()
object.openBrowse()



