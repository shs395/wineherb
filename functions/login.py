from selenium import webdriver

# 텍스트파일에서 아이디 비밀번호 불러오기
def openAccountInfo(siteName):
    try:
        with open(f"../keys/{siteName}_account.txt") as f:
            lines = f.readlines()
            id_ = lines[0].strip()
            pw_ = lines[1].strip()
    
        return (id_, pw_)

    except Exception as e:
        print(e)
    
def loginKlwines(driver):
    # klwines 계정정보 불러오기
    (klwines_id, klwines_pw) = openAccountInfo('klwines')
    #로그인 안되어 있는 경우
    try:
        url = 'https://www.klwines.com/account/login'
        driver.get(url)
        driver.find_element_by_id('Email').send_keys(klwines_id)
        driver.find_element_by_id('Password').send_keys(klwines_pw)
        driver.find_element_by_xpath('/html/body/div[2]/div/div[2]/div/div/div[2]/form/div[1]/table/tbody/tr[3]/td/input').click()
    # 로그인 정보가 남아 로그인 되어 있는 경우
    except:
        pass

def loginWineBid(driver):
    # winebid 계정정보 불러오기
    (winebid_id, winebid_pw) = openAccountInfo('winebid')
    try:
        url = 'https://www.winebid.com/SignIn'
        driver.get(url)
        driver.find_element_by_id('formModel_EmailAddress').send_keys(winebid_id)
        driver.find_element_by_id('formModel_Password').send_keys(winebid_pw)
        driver.find_element_by_id('SignIn').click()
    except:
        pass


def loginWineSearcher(driver):
    # winesearcher 계정정보 불러오기
    (winesearcher_id, winesearcher_pw) = openAccountInfo('winesearcher')
    try:
        url = 'https://www.wine-searcher.com/sign-in'
        driver.get(url)
        driver.find_element_by_id('loginmodel-username').send_keys(winesearcher_id)
        driver.find_element_by_id('loginmodel-password').send_keys(winesearcher_pw)
        driver.find_element_by_id('pv_submit_F').click()
    # 로그인 정보가 남아 로그인 되어 있는 경우
    except:
        pass