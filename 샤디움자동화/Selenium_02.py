import Main as M

from selenium import webdriver

from selenium.webdriver.common.by import By

from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains

from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

import pyautogui
import clipboard
import time
import os
import random
import string
import openpyxl as op

IDPW = "12121212"
WALLET_ADDRESS = "0x0354111da43a5014dB6754e18b6f0DEc64D9A7d9"

#======================랜덤 샤디움 네이밍 생성======================

def SHARDEUM_NAME():
    SHARDEUM_NAME=10	# 문자의 개수(문자열의 크기)
    RANDOM_str = ""	# 문자열
    for i in range(SHARDEUM_NAME):
        RANDOM_str += str(random.choice(string.ascii_letters + string.digits))
    return(RANDOM_str)

##########################################################
#####################크롬 프로필 리스트화###################
##########################################################

# 크롬 계정데이터 모여있는 폴더 chrome://version
directory_chrome = r"C:\Users\thqud\AppData\Local\Google\Chrome\User Data"

# 해당 문자가 포함된 폴더만 리스트화 시키겠다 (프로필 뒤에 공백 필수)
include_text = 'Profile '

# 디렉토리명 필터링
DIR_PROFILE_LIST = (
    [d for d in os.listdir(directory_chrome) 
    if os.path.isdir(os.path.join(directory_chrome, d)) and include_text in d]
)

#폴더명 리스트화
print(len(DIR_PROFILE_LIST))
for i in range(len(DIR_PROFILE_LIST)):
    print(DIR_PROFILE_LIST[i])


##########################################################
#####################셀레니움 기본 세팅#####################
##########################################################
options = Options()
user_data = r"C:\Users\thqud\AppData\Local\Google\Chrome\User Data"
options.add_argument(f"user-data-dir={user_data}")
options.add_argument("--profile-directory=" + DIR_PROFILE_LIST[2])
options.add_experimental_option("detach", True)  # 화면이 꺼지지 않고 유지
options.add_argument("--start-maximized")  # 최대 크기로 시작

service = Service(ChromeDriverManager().install()) # 웹드라이버 설치
driver = webdriver.Chrome(service=service, options=options) # 웹드라이버 불러오기
action = ActionChains(driver)


##########################################################
#####################자동화 코드 시작######################
##########################################################
# 0 - 메타마스크 미리 로그인 (팝업방지용도)

driver.get("chrome-extension://nkbihfbeogaeaoehlefnkodbefgpgknn/home.html") #웹사이트 이동
M.TIME(5)
action.send_keys(IDPW).key_down(Keys.ENTER).pause(5).perform()
driver.find_element(By.CLASS_NAME,"selected-account__clickable").click()
WALLET_ADDRESS = clipboard.paste() #지갑주소 변수에 저장
action.reset_actions() #액션 종료

for i in range(2):

    # 20 - DotSHM 연결
    # 만약 안된다면 네트워크를 다시 샤디움 1.0으로 바꿔주세요
    M.TIME(3)
    driver.get("https://dotshm.me") #사이트 이동
    M.TIME(5)
    action.send_keys("Ll6EYQZD").perform() #10번 과정에서 사용했던 샤디움 도메인 그대로 불러오기
    action.reset_actions() #액션 종료
    driver.find_element(By.ID,"headlessui-menu-button-:R3el6:").click() #검색
    M.TIME(3)
    driver.find_element(By.ID,"headlessui-menu-items-:r0:").click() #Register
    M.TIME(6)
    for j in range(2):
        try:
            # 21 - DotSHM 메타마스크 연결
            driver.switch_to.window(driver.window_handles[1]) #크롬 팝업창으로 이동
            driver.find_element(By.CLASS_NAME,"button.btn--rounded.btn-primary").click() #다음
            M.TIME(3)
            driver.find_element(By.CLASS_NAME,"button.btn--rounded.btn-primary.page-container__footer-button").click() #연결
            M.TIME(7)
            driver.switch_to.window(driver.window_handles[0]) #원래 창으로 이동
            M.CLICK_MOSUE(1081, 821) #MATIC부분 흰색영역 클릭
            action.send_keys(Keys.TAB).send_keys(Keys.ENTER).pause(4).send_keys((Keys.TAB) * 2).send_keys(Keys.ENTER).pause(10).perform()
            action.reset_actions() #액션 종료
        except:
            print("# 21 DotSHM 메타마스크 연결부 실패")
            break

        try:
            # 22 - DotSHM 가스비 요구
            driver.switch_to.window(driver.window_handles[1]) #크롬 팝업창으로 이동
            driver.find_element(By.CLASS_NAME,"button.btn--rounded.btn-primary.page-container__footer-button").click() #다음
            driver.switch_to.window(driver.window_handles[0]) #크롬 팝업창으로 이동
            action.pause(10).send_keys((Keys.TAB) * 2).send_keys(Keys.ENTER).pause(5).perform()
            action.reset_actions() #액션 종료
        except:
            print("# 22 DotSHM 가스비 요구부분 실패")
            break