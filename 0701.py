from selenium import webdriver
chrome_driver_path = 'C:/Users/admin/AppData/Local/Programs/Python/Python37/chromedriver.exe'
import time
# 设置 ChromeDriver 的路径
from selenium.webdriver.chrome.service import Service
service = Service(executable_path=chrome_driver_path)
options = webdriver.ChromeOptions()
driver = webdriver.Chrome(service=service, options=options)
driver.get('https://www.baidu.com')
time.sleep(50)
driver.quit()