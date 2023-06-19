from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.options import Options
import time

options = Options()
driver = webdriver.Edge(options=options)

driver.get("https://cn.bing.com/")

driver.implicitly_wait(5)

text_box = driver.find_element(by=By.ID, value="sb_form_q")
submit_button = driver.find_element(by=By.ID, value="search_icon")

text_box.send_keys("Selenium")
driver.implicitly_wait(5)

submit_button.click()
driver.implicitly_wait(5)

time.sleep(5)

driver.quit()
