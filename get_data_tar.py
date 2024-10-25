import glob
import os
import pandas as pd
import shutil
from datetime import datetime, timedelta
import traceback
# import autoit
import time
from selenium.webdriver.common.action_chains import ActionChains
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

from setting_fbb import *
import setting_fbb

def clear_path(path):
    files = glob.glob(os.path.join(path,'*'))
    for f in files:
        os.remove(f)

def set_up():
    options = webdriver.ChromeOptions()
    prefs = {'download.default_directory': Data_path}
    options.add_experimental_option('prefs', prefs)
    options.add_argument('ignore-certificate-errors')
    options.add_argument('allow-running-insecure-content')
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("disable-infobars")
    options.add_argument("--disable-extension")
    options.add_experimental_option("detach", True)

    # Open Chrome
    dr = webdriver.Chrome(options=options)
    dr.implicitly_wait(30)
    return dr

def log_in(dr):
    dr.maximize_window()
    login_ins = setting_fbb.Login()
    login_ins.login()
    login_ins.webdriver()
    dr.get('https://tha-crm.wiz.ai/#/login')
    time.sleep(10)
    # Fill Username & Password
    dr.find_element(By.XPATH, '//*[@id="login-container"]/div[3]/form/div[1]/div/div/input').clear()
    time.sleep(1)
    dr.find_element(By.XPATH, '//*[@id="login-container"]/div[3]/form/div[1]/div/div/input').send_keys(login_ins.user)
    time.sleep(2)
    dr.find_element(By.XPATH, '//*[@id="login-container"]/div[3]/form/div[2]/div/span/div/input').clear()
    time.sleep(1)
    dr.find_element(By.XPATH, '//*[@id="login-container"]/div[3]/form/div[2]/div/span/div/input').send_keys(login_ins.password)
    time.sleep(2)

    # Slide to Login
    element = dr.find_element(By.CSS_SELECTOR, 'div.handler.icon-qianjin')
    offset_x = 350
    actions = ActionChains(dr)
    time.sleep(2)
    actions.click_and_hold(element).perform()
    time.sleep(2)
    actions.move_by_offset(offset_x, 0).perform()
    time.sleep(2)
    actions.release().perform()

    # Select Environment
    dr.find_element(By.XPATH, va).click()
    time.sleep(2)
    dr.find_element(By.XPATH, '/html/body/div[3]/div/div[3]/span/button[2]/span').click()
    time.sleep(2)
    return dr

def get_data(dr):
    dr.get('https://tha-crm.wiz.ai/#/smartReception/components/records/index')
    time.sleep(10)

    # Click Filter Button
    dr.find_element(By.XPATH,
                    '//*[@id="crm"]/div[1]/div[1]/div[2]/div[2]/div[2]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div/div[2]/button[1]/span/i').click()
    time.sleep(10)

    dr.find_element(By.XPATH,
                    "//input[@type='text' and @autocomplete='off' and @placeholder='Select time' and contains(@class, 'el-input__inner')]").click()
    time.sleep(3)

    wait = WebDriverWait(dr, 10)
    input_element = wait.until(
        EC.element_to_be_clickable((By.XPATH,
                                    "//input[@type='text' and @autocomplete='off' and @placeholder='Select date' and contains(@class, 'el-input__inner')]"))
    )

    # Get the current date in the format DDMMYYYY
    current_date = datetime.now().strftime('%d/%m/%Y')

    # Click the input element to focus on it (if needed)
    input_element.click()
    time.sleep(1)

    # Clear the input field if necessary
    input_element.clear()
    time.sleep(1)

    # Input the current date
    input_element.send_keys(current_date)
    time.sleep(1)

    # Click Ok
    dr.find_element(By.XPATH, '/html/body/div[3]/div[2]/button[2]/span').click()
    time.sleep(3)

    apply_filters_element = dr.find_element(By.XPATH, "//span[text()='Apply Filters']")
    apply_filters_element.click()
    time.sleep(10)

    export_element = dr.find_element(By.XPATH, "//span[text()='Export']")
    export_element.click()
    time.sleep(10)

    export_button = dr.find_element(By.XPATH,
                                    "//button[@class='el-button filter__btn el-button--primary el-button--small']//span[text()='Export']")
    export_button.click()
    time.sleep(30)

    # Find all "Download" buttons on the page
    download_buttons = dr.find_elements(By.XPATH,
                                        "/html/body/div[4]/div/div[2]/div/div[1]/div[4]/div[2]/table/tbody/tr[1]/td[7]/div/span/button[1]/span/i")

    # Click the first "Download" button
    if download_buttons:
        first_download_button = download_buttons[0]
        first_download_button.click()

        # Monitor the download folder for the file to be downloaded
        timeout = 180  # seconds
        start_time = time.time()
        while True:
            if any(file.endswith('.xlsx') for file in os.listdir(filedownload_path)):
                print('Download Success')
                break
            if time.time() - start_time > timeout:
                raise TimeoutError("File download timed out.")
            time.sleep(1)  # Polling interval

        time.sleep(20)

        # Close the browser
        dr.quit()


if __name__ == '__main__':
    # clear_path(filedownload_path)

    global_driver = set_up()
    time.sleep(1)
    log_in(global_driver)
    time.sleep(1)
    get_data(global_driver)


