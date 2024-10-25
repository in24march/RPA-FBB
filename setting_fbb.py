import os
from datetime import datetime, timedelta
from pathlib import Path

class Login:
    def __init__(self) -> None:
        pass
    
    def login(self):
        self.user = 'nichamot'
        self.password = 'Nummon@082024'
    
    def webdriver(self):
        self.fbb = "https://tha-crm.wiz.ai/#/login"
class find_file:
    def __init__(self,path):
        self.path = path
    def file_last_time(self):
        path_file = self.path
        file_excelfile = [file for file in os.listdir(path_file) if file.endswith('.xlsx')]
        if file_excelfile:
            file_ex_time = max(
                (os.path.join(path_file, file) for file in file_excelfile),
                key = os.path.getmtime
            )
            return file_ex_time

url = 'https://tha-crm.wiz.ai/#/login'
us = 'nichamot'
ps = 'Nummon@082024'

real_path = r"\\172.16.103.200\CSS_JobReport\Report Outbound Campaign Voice AI\Patch Data"
# test_real = r"D:\work_path\RPA_patch_data"
dir_path = os.path.dirname(os.path.abspath(__file__)) + os.sep

Data_path = dir_path + 'FBB DATA' + os.sep
Master_path = dir_path + 'Master' + os.sep
PP_path = dir_path + 'Play_Premium' + os.sep
Patch_data = dir_path + 'temp_patch' + os.sep
P_data = dir_path + 'Patch_dataq' + os.sep
txt_path = dir_path + 'edit_txt_patch' + os.sep

Path(Data_path).mkdir(parents=True, exist_ok= True)
Path(Master_path).mkdir(parents=True, exist_ok= True)
Path(PP_path).mkdir(parents=True, exist_ok= True)
Path(P_data).mkdir(parents=True, exist_ok= True)


date = datetime.now() - timedelta(days=0) # วันไฟล์ตรงนี้

current_date = date.strftime("%d-%m-%Y")    # เปลี่ยนวันที่ให้เหมาะสมกับชื่อโฟลเดอร์
date_folder_path = os.path.join(real_path, current_date)


Path(date_folder_path).mkdir(parents=True, exist_ok=True)
print(f"Create folder: {date_folder_path}")

folder_main1 = "Patch data"
folder_main2 = "Text patch data"
folder_main3 = "Play premium"

sub1_folder = os.path.join(date_folder_path, folder_main1)
Path(sub1_folder).mkdir(parents=True, exist_ok=True)
print(f"Create folder: {sub1_folder}")


sub2_folder = os.path.join(date_folder_path, folder_main2)
Path(sub2_folder).mkdir(parents=True, exist_ok=True)
print(f"Create folder: {sub2_folder}")

sub3_folder = os.path.join(date_folder_path, folder_main3)
Path(sub3_folder).mkdir(parents=True, exist_ok=True)
print(f"Create folder: {sub3_folder}")