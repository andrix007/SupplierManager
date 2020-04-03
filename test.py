import time
from selenium import webdriver

driver = webdriver.Chrome('C:/Users/Andrei Bancila/Desktop/SupplierListManagerProject/chromedriver/chromedriver.exe')  # Optional argument, if not specified will search path.
driver.get('http://btb.cnlmusic.kr/btb/bbs/login.php');
time.sleep(0.5)
driver.find_element_by_id("login_id").send_keys("NICHE")
driver.find_element_by_id("login_pw").send_keys("Horia@160")
driver.find_element_by_xpath("//*[@id=\"mb_login\"]/div/div/div/section/form/div[3]/input").click()
time.sleep(0.5)
driver.find_element_by_xpath("//*[@id=\"sidebar-menu\"]/div/ul/li[2]/a").click()
time.sleep(0.5)
driver.find_element_by_xpath("/html/body/div[1]/div/div[3]/div/div[2]/div/div/div[1]/div[2]/a").click()
time.sleep(1)
driver.quit()
