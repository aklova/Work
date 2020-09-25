from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver import DesiredCapabilities
from getpass import getpass
import glob, os
import time
import ssl
import shutil
import subprocess
import ctypes
from subprocess import Popen, PIPE, STDOUT
import sys
import win32ts

shutil.rmtree('C:\#DellBios', ignore_errors=True)

#-----Finding downloaded file in C:\#DellBios\ and running it later--------
def downloaded():
    for fn in glob.glob("C:\#DellBios\*.exe"):
        arguments = ("/s /r")
        subprocess.call(fn + ' ' + arguments)

#-------------------------------------------

#-----is Locked-------------------------
def isComputerLocked():
    time.sleep(30)
    process_name='LogonUI.exe'
    callall='TASKLIST'
    outputall=subprocess.check_output(callall)
    outputstringall=str(outputall)
    if process_name in outputstringall:
        return 'Locked'
        isComputerLocked()
        #print(isLocked)
    else:
        return 'Unlocked'
        #print(isLocked)

#-------------------------------------------

#-----Message popup-------------------------
def newMessage():

    sessionID = win32ts.WTSGetActiveConsoleSessionId()
    messageTitle = 'Western Union End User Computing: BIOS Update Notification'

    count = 3
    while count > -1:
        if isComputerLocked() != 'Unlocked':
            print('Computer is locked, waiting for user.')
        else:
            print('Computer is unlocked, sending message.')
            if win32ts.WTSSendMessage(0, sessionID, messageTitle, 'Your BIOS is out of date and needs to be updated.\n\n Your Computer will reboot to perform the update, please save your work and click "OK" to continue, or click "Cancel" to postpone for 10 minutes.\n\n You can postpone {} more times.'.format(count), 1, 1800, True) == 1:
                print('User pushed OK button, starting program.')
                disablingBitlocker()
                downloaded()
                break
            else:
                if count == 0:
                    print('User pushed OK button, starting program.')
                    disablingBitlocker()
                    downloaded()
                    break
                else:
                    print('User pushed Cancel, postponing program.')
                    time.sleep(600)
                    count = count - 1
                    continue

#---------------------------------------------

#-----Bitlocker disable command--------------
def disablingBitlocker():
    system_root = os.environ.get('SystemRoot', 'C:\Windows')
    system32 = 'System32'
    manage_bde = os.path.join(system_root, system32, 'manage-bde.exe')


    p = Popen([manage_bde, '-status', 'c:'], stdout=PIPE, stderr=STDOUT)
    lines = p.stdout.readlines()
    #print(lines)

    data = ''
    for line in lines:
        line = line.decode('utf-8')
        data += line
    #print(data)

    if 'Protection On' in data:
        print('Bitlocker is enabled. Disabling Bitlocker')
        p = Popen([manage_bde, '-protectors', '-disable', 'c:'], stdout=PIPE, stderr=STDOUT)

#---------------------------------------------
#--------Checking if versions doesn't match---
def checkingIfMatch():
    bio = os.popen('wmic bios get smbiosbiosversion').read().replace("\n","").replace(" ","").replace(" ","")
    bio = bio.lstrip('SMBIOSBIOSVersion')
    #print(bio)

    for file in glob.glob("C:\#DellBios\*.exe"):
        if bio in file:
            print('The newest BIOS version is installed, exiting script.')
            shutil.rmtree('C:\#DellBios', ignore_errors=True)
            sys.exit()
        else:
            print('Newer BIOS version downloaded, starting installation.')

#---------------------------------------------


def downloadingBios():
    #user_agent = 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/60.0.3112.50 Safari/537.36'
    #ctx = ssl.create_default_context()
    #ctx.check_hostname = False
    #ctx.verify_mode = ssl.CERT_NONE
    capabilities = DesiredCapabilities.CHROME.copy()
    capabilities['acceptSslCerts'] = True
    capabilities['acceptInsecureCerts'] = True

    options = webdriver.ChromeOptions()
    preferences = {"download.default_directory": "C:\#DellBios", "safebrowsing.enabled": "false"}
    options.add_experimental_option("prefs", preferences)
    #options.add_argument("--window-position=-32000,-32000")
    #options.add_argument("--headless")
    #options.add_argument("--window-size=0,0")




    tag = os.popen("wmic bios get serialnumber").read().replace("\n","").replace(" ","").replace(" ","")
    tag = tag.lstrip('SerialNumber')


    driver = webdriver.Chrome(options=options, desired_capabilities=capabilities)
    #driver.minimize_window()
    driver.get('https://www.dell.com/support/home/lt/en/ltbsdt1/?app=drivers')
    time.sleep(10)

    try:
        service_tag = driver.find_element_by_name('entry-main-input')
    except:
        driver.quit()
        print("Something went wrong, couldn't download BIOS, try again later.")
        sys.exit()

    service_tag.send_keys(tag)

    login_btn = driver.find_element_by_class_name('input-group-btn')
    login_btn.submit()

    time.sleep(10)
    current_website = driver.current_url
    #print(current_website)
    driver.quit()

    tag = os.popen("wmic computersystem get model").read().replace("\n","")
    x = tag.split()
    serialNumber = x[1] + '_'

    driver = webdriver.Chrome(options=options)
    #driver.minimize_window()
    driver.get(current_website)
    time.sleep(10)
    count = 0

    for a in driver.find_elements_by_xpath('.//a'):
        if a.get_attribute('href') == None or a.get_attribute('href') == '':
            continue
        else:
            links = a.get_attribute('href')
            #print(links)
            if serialNumber in links:
                count = count + 1
                link = links
                if count > 1:
                    break
                #print(links)
    #print(count)
    #print(link)
    #webbrowser.open(link)
    driver.get(link)
    time.sleep(10)
    driver.quit()






#----Running downloaded file----------
#downloaded()


#-----Running program with message (includes running bios upgrade)----
downloadingBios()
checkingIfMatch()
newMessage()
