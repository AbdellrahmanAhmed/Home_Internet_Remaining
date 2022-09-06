import ctypes
import os
import shutil
import sys
from datetime import date
from datetime import datetime, timedelta

from openpyxl import load_workbook
from openpyxl.styles import Alignment
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from winotify import Notification
from colorama import Fore

def some_job():
    global msg
    try:
        print("Decorated job")
        wb = load_workbook('Home Internet Remaining.xlsx')

        cell = 1

        today = date.today()
        now = datetime.now()

        month = today.strftime("%d/%m%Y")
        timeN = now.strftime("%H:%M:%S")
        todayN = today.strftime("%A")

        def writeCell(cell0):
            ws[f'A${cell0}'].value = today
            ws[f'B${cell0}'].value = month
            ws[f'C${cell0}'].value = todayN
            ws[f'D${cell0}'].value = timeN
            ws[f'E${cell0}'].value = remaining

        ws = wb.active
        cellValue = ws[f'A${cell}'].value

        chrome_options = Options()
        chrome_options.add_argument(
            "--user-data-dir=C:/Users/Abdel/AppData/Local/Google/Chrome/User Data/unique", )  # change to profile path
        chrome_options.add_argument('--profile-directory=Default')
        driver = webdriver.Chrome("C:\Program Files\Google\Chrome\Application\chromedriver.exe")

        driver.minimize_window()

        driver.get('https://my.te.eg/user/login')

        driver.implicitly_wait(3)

        service_Number = driver.find_element(by=By.XPATH, value='//*[@id="login-service-number-et"]')
        password = driver.find_element(by=By.XPATH, value='//*[@id="login-password-et"]')
        p_Button_Label = driver.find_element(by=By.XPATH, value='//*[@id="login-login-btn"]/span[2]')

        service_Number.send_keys('0663696485')
        password.send_keys('01552337484we')
        p_Button_Label.click()

        driver.implicitly_wait(5)

        old_Remaining = driver.find_element(by=By.XPATH,
                                            value='//*[@id="pr_id_7"]/div/div/div/div/div/app-gauge/div[2]/span[1]').text
        remaining = ""
        for c in old_Remaining:
            if c == ".":
                remaining = remaining + c
            elif c.isdigit():
                remaining = remaining + c

        while cellValue:
            cell = cell + 1
            cellValue = ws[f'A${cell}'].value

        writeCell(cell)

        cellList = ['A', 'B', 'C', 'D', 'E']

        for cellC in cellList:
            currentCell = ws[f'${cellC}${cell}']
            currentCell.alignment = Alignment(horizontal='center')

        lastRemaining = ws[f'E${cell - 1}'].value
        (used) = round(float(lastRemaining) - float(remaining), 2)
        # print(f'Last Used {round(used, 2)} GB')

        last_Remaining = ws[f'E${cell - 1}'].value
        remaining_days = date(2022, 9, 18) - today
        recommended_Remaining = round(float(remaining) / float(remaining_days.days), 2)

        tt = str(ws[f'D${cell - 1}'].value)
        after_hour_date_time = datetime.now() + timedelta(hours=1)

        bTime = datetime.strptime(tt, "%H:%M:%S")
        btTime = datetime.strptime(timeN, "%H:%M:%S") - bTime

        cT = []
        tSTR = f"{today - timedelta(days=1)} 00:00:00"


        for c in range(1, 1000):
            if str(ws[f'A${c}'].value) == str(tSTR):
                v = ws[f'E{c}'].value
                cT.append(float(v))
            elif not ws[f'A{c}'].value:
                if not cT:
                    cT.append(float(remaining))
                    break

        totalDayUsed = float(cT[-1]) - float(remaining)

        msg = f'''
        Remaining {remaining} GB 
        recommended consumption  {recommended_Remaining} GB 
        and Total Day Remaining {round(totalDayUsed, 2)}GB
        '''
        print(msg)

        # Next Reading {after_hour_date_time.strftime('%H:%M:%S')}

        driver.close()
        wb.save('Home Internet Remaining.xlsx')

        shutil.copy(f'Home Internet Remaining.xlsx', f'Home Internet Remaining {today}.xlsx')


        return str(msg)


    except Exception as e:
        print('***************************************************************')
        print(Fore.RED+str(e))
        print('***************************************************************')

        driver.close()
        some_job()

    return str(msg)


toast = Notification(
    app_id="Home Internet",
    title="تم حساب المتبقي من الإنترنت وتخزينه",
    msg=some_job(),
    duration="long",
)

toast.add_actions(label="للدخول إلى صفحتك الشخصية  ", launch="https://my.te.eg/user/login")

toast.show()

sys.exit()
