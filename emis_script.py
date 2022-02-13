from datetime import timedelta
from re import L
import time


from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys

from selenium.webdriver.chrome.service import Service as ChromeService # Similar thing for firefox also!
from subprocess import CREATE_NO_WINDOW # This flag will only be available in windows


import configparser
from plyer import notification
import os
import tkinter
import threading

import requests

def run_script():
    current_label['text'] = 'Checking Emis availability...'

    url = "https://emis.freeuni.edu.ge/index.php?sub_page=nishnebi"
    try:
        request_response = requests.head(url)
        status_code = request_response.status_code
    except:
        status_code = -1
    website_is_up = status_code == 200

    if not website_is_up:
        current_label['text'] = 'Emis is offline'
        notification.notify('Emis is offline', 'Can\'t run script, Emis is currently offline', app_icon = 'data/notification.ico')
        return

    current_label['text'] = 'Fetching credentials...'
    
    
    # SERVICES
    chrome_service = ChromeService('data/chromedriver')
    chrome_service.creationflags = CREATE_NO_WINDOW

    # OPTIONS
    option = webdriver.ChromeOptions()
    option.add_argument('headless')
    option.add_argument("--window-size=1920x1080");
    
    current_label['text'] = 'Logging in...'
    driver = webdriver.Chrome('data/chromedriver.exe', options = option, service=chrome_service)  # Optional argument, if not specified will search path.
    try:
        driver.get(url)

        element = WebDriverWait(driver, 10).until( # Wait for page to load
            lambda x: x.find_element_by_class_name('login_with_google'))

        login_button = driver.find_element_by_class_name('login_with_google')
        login_button.click() # Click on login button

        #
        # LOGIN PAGE PROCESSING
        #
        parentWindowHandle = driver.current_window_handle # Get parent window
        print("Parent Window ID is : " + parentWindowHandle)
        loginWindowHandle = driver.window_handles[1] # Get login window
        print("Login Window ID is : " + loginWindowHandle)

        driver.switch_to.window(loginWindowHandle) # Switch to login window



        email = driver.find_element_by_xpath('/html/body/div/div[2]/div[2]/div[1]/form/div/div/div/div/div/input[1]')
        #email = driver.find_element_by_name('identifier') # NORMAL MODE
        email.send_keys(my_name)
        #driver.find_element_by_id('identifierNext').click() # NORMAL MODE
        driver.find_element_by_xpath('/html/body/div/div[2]/div[2]/div[1]/form/div/div/input').click()

        element = WebDriverWait(driver, 10).until(
                #EC.element_to_be_clickable((By.NAME, 'password')) # NORMAL MODE
                EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[2]/div/div[2]/form/span/div/div[2]/input'))
            )

        password = driver.find_element_by_xpath('/html/body/div[1]/div[2]/div/div[2]/form/span/div/div[2]/input')
        #password = driver.find_element_by_name('password') # NORMAL MODE
        password.send_keys(my_code)
        #driver.find_element_by_id('passwordNext').click() # NORMAL MODE
        driver.find_element_by_xpath('/html/body/div[1]/div[2]/div/div[2]/form/span/div/input[2]').click()

        driver.switch_to.window(parentWindowHandle) # Switch to parent window
        
        #---------------------------------------------------
        current_label['text'] = 'Checking grades...'
        element = WebDriverWait(driver, 10).until( # Wait for page to load
            lambda x: x.find_element_by_class_name('sem_block'))


        currSemester = driver.find_element_by_xpath('//html/body/table/tbody/tr/td[2]/div/div[1]/div[2]') # Points to semester block


        #
        # Getting information for new_grades.txt
        #
        new_grades = open('data/new_grades.txt', 'w', encoding='utf-8')
        subjectElements = driver.find_elements_by_xpath('//html/body/table/tbody/tr/td[2]/div/div[1]/div[2]/table/tbody/tr') # List of elements pointing to their respective subjects

        for i in range(len(subjectElements)): # Click on 'შეფასება' for every subject and save info in new_grades.txt
            #subjectName = driver.find_element_by_xpath('//html/body/table/tbody/tr/td[2]/div/div[1]/div[2]/table/tbody/tr[' + str(i+1) + ']/td[1]').get_attribute('innerHTML').split(' ', 1)[1]
            #subjectGrade = driver.find_element_by_xpath('//html/body/table/tbody/tr/td[2]/div/div[1]/div[2]/table/tbody/tr[' + str(i+1) + ']/td[4]').get_attribute('innerHTML').replace('-', '0')
            
            subjectMore = driver.find_element_by_xpath('//html/body/table/tbody/tr/td[2]/div/div[1]/div[2]/table/tbody/tr[' + str(i+1) + ']/td[6]')
            
            element = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable(subjectMore)
            )
            
            subjectMore.click()
            time.sleep(2)
            
            subjectName = driver.find_element_by_class_name('st_book_item').get_attribute('innerHTML')
            new_grades.write(subjectName + ': ')
            grades = driver.find_elements_by_class_name('st_est_item')
            for j in range(len(grades)):
                new_grades.write(' ' + grades[j].get_attribute('innerHTML'))
            new_grades.write('\n')
            webdriver.ActionChains(driver).send_keys(Keys.ESCAPE).perform()
        driver.quit()
        new_grades.close()

        current_label['text'] = 'Comparing to old grades...'
        if not os.path.exists('data/old_grades.txt'):
            os.rename('data/new_grades.txt', 'data/old_grades.txt')
            current_label['text'] = 'Done'
            return

        # Check for old_grades file and compare both of them to see any inconsistencies
        old_grades = open('data/old_grades.txt', 'r', encoding='utf-8')
        new_grades = open('data/new_grades.txt', 'r', encoding='utf-8')
        notification_message = ''

        old_line = old_grades.readline()
        new_line = new_grades.readline()

        while new_line:
            if new_line != old_line:
                notification_message = notification_message + new_line.split(':')[0] + ' has just been updated!\n'
            old_line = old_grades.readline()
            new_line = new_grades.readline()

        old_grades.close()
        new_grades.close()
        old_grades = open('data/old_grades.txt', 'w', encoding='utf-8')
        new_grades = open('data/new_grades.txt', 'r', encoding='utf-8')

        old_grades.writelines(new_grades.readlines())

        old_grades.close()
        new_grades.close()
        os.remove('data/new_grades.txt')
        
        if notification_message != '':
            notification.notify('Emis Update', notification_message, app_icon = 'data/notification.ico')
        current_label['text'] = 'Done'
        return
    except:
        driver.quit()
        current_label['text'] = 'Failed'
        notification.notify('Emis is offline', 'Can\'t run script, Emis is currently offline', app_icon = 'data/notification.ico')
        return


def change_status(stat):
    if stat == 'running':
        status_value.config(text = 'Running')
        status_value.config(fg = 'lightgreen')
    else:
        status_value.config(text = 'Not running')
        status_value.config(fg = 'red')
        last_updated_value.config(text = time.strftime("%H:%M:%S", time.localtime()))


def auto_run_button_handle():
    if auto_run_button['text'] == 'Start Autorun':
        auto_run_button['text'] = 'Stop Autorun'
        auto_run_button['fg'] = 'red'
        print('Started auto run')
        th = threading.Thread(target=run_script_auto)
        th.start()
        #run_script_auto()
    else:
        auto_run_button['text'] = 'Start Autorun'
        auto_run_button['fg'] = 'lightgreen'
        print('Stopped auto run')

def run_script_auto():
    if auto_run_button['text'] == 'Stop Autorun':
        threading.Timer(my_timer, run_script_auto).start()
        if (run_button['state'] == tkinter.NORMAL):
            print('Trying to run script')
            button_click_handle()
            return
        else:
            print('ERROR: Script already started, ignoring call')
    else:
        print('Autorun disabled, process halted')


def run_script_handle():
    run_script()
    
    change_status('stopped')
    run_button.config(state=tkinter.NORMAL)

def button_click_handle():
    run_button.config(state=tkinter.DISABLED)
    th = threading.Thread(target=change_status('running'))
    th.start()
    th = threading.Thread(target=run_script_handle)
    th.start()

if __name__ == '__main__':
    tk = tkinter.Tk()
    tk.configure(bg='#303436')
    tk.title('Emis App')
    tk.geometry('300x100')
    tk.columnconfigure(0, weight=1)
    tk.columnconfigure(1, weight=1)
    tk.columnconfigure(2, weight=1)
    stopAutoRun = True

    status_label = tkinter.Label(text = 'Status:', fg = 'white', bg = '#303436') # label for script status
    status_value = tkinter.Label(text = 'Not running', fg = 'red', bg = '#303436') # shows whether script is running
    run_button = tkinter.Button(text = 'Run script', fg = 'lightgreen', bg = '#303436', command = button_click_handle) #Run script button
    last_updated_label = tkinter.Label(text = 'Last updated:', fg = 'white', bg = '#303436') # label for last update
    last_updated_value = tkinter.Label(text = 'No update yet', fg = 'white', bg = '#303436')
    auto_run_button = tkinter.Button(text = 'Start Autorun', fg = 'lightgreen', bg = '#303436', command = auto_run_button_handle) #Run script button
    current_label = tkinter.Label(text = 'Doing nothing', fg = 'white', bg = '#303436')
    timer_label = tkinter.Label(text = '00:00', fg = 'white', bg = '#303436')
    status_label.grid(column=0, row=0)
    status_value.grid(column=1, row=0)
    run_button.grid(column=2, row=0)
    last_updated_label.grid(column=0, row=1)
    last_updated_value.grid(column=1, row=1)
    auto_run_button.grid(column=2, row=1)
    current_label.grid(column=0, row=2)
    timer_label.grid(column=2,row=2)

    conf = configparser.RawConfigParser()
    confPath = r'data/config.config'
    conf.read(confPath)
    my_name = conf.get('config', 'my_name')
    my_code = conf.get('config', 'my_code')
    my_timer = int(conf.get('config', 'timer'))
    timer_label['text'] = timedelta(seconds=my_timer)
    tk.mainloop()