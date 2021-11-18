import platform
import chromedriver_autoinstaller
from webdriver_manager.chrome import ChromeDriverManager
import subprocess
from selenium.webdriver.chrome.options import Options
from selenium import webdriver


def openChrome():
    print("current platform : " + platform.system())
    try:
        if(platform.system() == 'Windows'):
            try:          
                subprocess.Popen(r'C:\Program Files\Google\Chrome\Application\chrome.exe --remote-debugging-port=9222 --user-data-dir="C:\chrometemp"')
            except:
                subprocess.Popen(r'C:\Program Files (x86)\Google\Chrome\Application\chrome.exe --remote-debugging-port=9222 --user-data-dir="C:/ChromeTEMP"')

            option = Options()
            option.add_experimental_option("debuggerAddress", "127.0.0.1:9222")

            chrome_ver = chromedriver_autoinstaller.get_chrome_version().split('.')[0]
            try:
                driver = webdriver.Chrome(f'./{chrome_ver}/chromedriver.exe', options=option)
            except:
                chromedriver_autoinstaller.install(True)
                driver = webdriver.Chrome(f'./{chrome_ver}/chromedriver.exe', options=option)
            driver.implicitly_wait(10)

        elif(platform.system() == 'Darwin'):
            subprocess.Popen(r'/Applications/Google\ Chrome.app/Contents/MacOS/Google\ Chrome --remote-debugging-port=9222 --user-data-dir="~/ChromeProfile"', shell=True)

            chromedriver_autoinstaller.install()

            co = Options()
            co.add_experimental_option('debuggerAddress', '127.0.0.1:9222')
            driver = webdriver.Chrome(options=co)
            driver.implicitly_wait(10)
    except:
        driver = webdriver.Chrome(ChromeDriverManager().install())

    return driver