import logging
import pandas as pd
from datetime import datetime, timedelta
import time

from selenium import webdriver
from selenium.webdriver.support.select import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common import exceptions
import pandas as pd
import glob

from setting_fbb import *
from map_data import *
import setting_fbb


def login():
    try:
        option = webdriver.ChromeOptions()
        pref = {'download.default_directory': Data_path}
        option.add_experimental_option('prefs', pref)
        option.add_argument('ignore-certificate-errors')
        option.add_argument("--no-sandbox")
        option.add_experimental_option("detach", True)
        option.add_argument("--disable-dev-shm-usage")
        driver = webdriver.Chrome(options= option)
        driver.implicitly_wait(30)
    except Exception as e:
        print(e)
    
    login_ins = setting_fbb.Login()
    login_ins.login()
    login_ins.webdriver()
    driver.get(login_ins.fbb)
    print("Open webdriver")
    driver. maximize_window()
    driver.find_element(By.NAME, 'account').send_keys(login_ins.user)
    driver.find_element(By.NAME, 'password').send_keys(login_ins.password)
    
    element = driver.find_element(By.CSS_SELECTOR, 'div.handler.iconfont-menu.icon-qianjin')
    offset_x = 350
    actions = ActionChains(driver)
    time.sleep(2)
    actions.click_and_hold(element).perform()
    time.sleep(2)
    actions.move_by_offset(offset_x, 0).perform()
    time.sleep(2)
    actions.release().perform()
    
    driver.find_element(By.XPATH, '/html/body/div[3]/div/div[2]/div[2]/div/div[2]/div/div/div/span[3]').click()
    driver.find_element(By.XPATH, '/html/body/div[3]/div/div[3]/span/button[2]/span').click()
    return driver

def search_master(driver ,value):
    finder_file = find_file(Master_path)
    file_master = finder_file.file_last_time()
    df = pd.read_excel(file_master, sheet_name= 'wiz')
    ai_outbound_element = driver.find_element(By.XPATH, "//div[@class='el-tooltip navBar-navItem-overflow']/span[text()=' AI Outbound']")
    ai_outbound_element.click()
    outbound_call_el = driver.find_element(By.XPATH, "//div[@class='el-tooltip navBar-navItem-overflow']/span[text() =' Outbound Call Task']")
    outbound_call_el.click()
    input_tesk_master = driver.find_element(By.XPATH, "//input[@class='el-input__inner']")
    input_tesk_master.clear()
    input_tesk_master.send_keys('Disney Plus HotStar' + Keys.ENTER)
    element = driver.find_element(By.XPATH, f"//div[contains(text(), '{value}')]")
    element.click()
    time.sleep(10)
    # export_button = WebDriverWait(driver, 20).until(
    # EC.presence_of_element_located((By.XPATH, "//button[contains(@class, 'el-button--default--bg') and contains(@class, 'el-button--default') and contains(@class, 'el-button--small')]//span[text()='Export']"))
    # )
    # export_button.click()
    js = """
            $(document).ready(function() {
                // ค้นหาและคลิกปุ่มแรกที่เปิดป๊อปอัพ
                $("span:contains('Export')").closest("button").click();

                // รอให้ป๊อปอัพโหลดเสร็จ
                setTimeout(function() {
                    $("button:contains('Export')").filter(".el-button--primary").click();
                }, 2000);
            });
         """
    driver.execute_script(js)
    
def Download_Data(driver, task_type):
    max_attempts = 5
    attempt = 0
    while attempt < max_attempts:
        try:
            time.sleep(5)
            # ค้นหาอิลิเมนต์ที่ต้องการ
            download_button = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "/html/body/div[3]/div/div[2]/div/div[1]/div[4]/div[2]/table/tbody/tr[1]/td[7]/div/span/button[1]"))
            )
            # ตรวจสอบข้อความภายในอิลิเมนต์
            button_text = download_button.text.strip()

            if "Download" in button_text:  # ถ้าข้อความเป็น "Download"
                driver.execute_script("arguments[0].click();", download_button)
                print(f"Click successful on attempt {attempt + 1}")
                break
            else:
                print(f"Button text is '{button_text}' on attempt {attempt + 1}, waiting for 'Download'")
                attempt += 1
        except Exception as e:
            print(e)
            print(f"Error on attempt {attempt + 1}: {e}")
            attempt += 1

    if attempt == max_attempts:
        print("Max attempts reached. The button text never changed to 'Download'.")
        
    # ตรวจสอบว่า jQuery ยังคงมี request active อยู่หรือไม่
    wait = 1
    while wait == 1:
        wait = driver.execute_script('return jQuery.active;')  # ถ้าไม่มี active requests, jQuery.active จะเป็น 0
        time.sleep(0.5)

        # รอเวลาหลังจากการดาวน์โหลดเริ่มต้น
    time.sleep(5)
    logging.debug('Downloading')

        # ตั้งเวลารอเพิ่มเติมเพื่อให้การดาวน์โหลดเสร็จสมบูรณ์
    download_complete = False
    for _ in range(30):  # ลองเช็คการดาวน์โหลดทุกๆ 2 วินาที นานสุด 1 นาที
        chech_xlsx = glob.glob(os.path.join(Data_path, '*.xlsx'))
        if chech_xlsx:
            download_complete = True
            break
        time.sleep(2)

    if download_complete:
        print('Download Success.')
        
        last_file = max(chech_xlsx, key=os.path.getctime)
        
        base_name, ext = os.path.splitext(last_file)
        
        new_file_name = f"{base_name}_{task_type}{ext}"
        new_file_path = os.path.join(Data_path, new_file_name)
        os.rename(last_file, new_file_path)
        print(f"Renamed file to: {new_file_name}")
    else:
        print('Download fail')

    logging.info('Quit web driver')
    driver.close()

def get_data(date):
    finder_file = find_file(Master_path)
    file_master = finder_file.file_last_time()
    df = pd.read_excel(file_master, sheet_name= 'wiz')
    # driver = login()
    for index, row in df.iterrows():
        task_name = row['Task Name']
        task_type = row['Type']
        task_promotion = row['Promotion Name']
        # search_master(driver, task_name)
        # Download_Data(driver, task_type)
        Play_premium(task_promotion, task_type, date)
        
    
if __name__ == '__main__':
    get_data()
    