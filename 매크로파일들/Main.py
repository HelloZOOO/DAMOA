
import os
import subprocess

import pyautogui
import keyboard
import time

import openpyxl as op
import clipboard

import random
import string

#======================랜덤 EMAIL 생성======================

def RANDOM_EMAIL():
    RANDOM_EMAIL=20	# 문자의 개수(문자열의 크기)
    RANDOM_str = ""	# 문자열
    for i in range(RANDOM_EMAIL):
        RANDOM_str += str(random.choice(string.ascii_letters + string.digits))
    return(RANDOM_str + '@damoim.com')

#======================랜덤 EMAIL 생성======================

#파일 실행 함수 정의
def run(path):
    # 디렉토리 설정
    os.chdir(PATH_CHROME)
    # 파일 실행
    os.startfile(path)

#time 단축키
def TIME(how):
    pyautogui.sleep(how)

#tab 을 자동으로 눌러주는 함수
def TAB(how):
    for i in range(how) :
        pyautogui.hotkey('tab',interval = 0.02)

#alt + D 를 자동으로 눌러주는 함수
def ALT_D(how):
    for i in range(how) :  
        pyautogui.hotkey('alt', 'd')

#ctrl + A 를 자동으로 눌러주는 함수
def CTRL_A(how):
    for i in range(how) :  
        pyautogui.hotkey('ctrl', 'a')

#alt + F 를 자동으로 눌러주는 함수
def ALT_F4(how):
    for i in range(how) :
        pyautogui.hotkey('alt', 'f4')

#enter를 자동으로 눌러주는 함수
def ENTER(how):
    for i in range(how) :
        pyautogui.hotkey('enter')

#전체화면 단축키
def FULL_SCREEN(how):
    for i in range(how) :
        pyautogui.hotkey('f11')

#화면 깨끗하게
def CLEAR_SCREEN(how):
    for i in range(how) :
        pyautogui.hotkey('win','d')
        
#크롬 새탭
def NEW_TAB(how):
    for i in range(how) :
        pyautogui.hotkey('ctrl','t')

#크롬 탭뒤로
def BACK_TAB(how):
    for i in range(how) :
        pyautogui.hotkey('ctrl','shift','tab')

#마우스 클릭
def CLICK_MOSUE(x1, y1):
    pyautogui.click(x1, y1, button='left',clicks=1)

#문자입력
def WRITE(WORD):
    pyautogui.write(WORD)
    
#복사
def CTRL_C(how):
    pyautogui.hotkey('ctrl','c')

#클립보드카피
def CLIPBOARD(WORD):
    clipboard.copy(WORD)

#크롬강제종료
def KILL_CHROME():
    os.system('taskkill /f /im chrome.exe ') #프로세스명을 사용한 프로세스 종료