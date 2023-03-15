import os
import subprocess

import pyautogui
import keyboard
import time

import openpyxl as op
import clipboard

os.getcwd()

WIDTH, HEIGHT = pyautogui.size()

##크롬 프로필 폴더 위치 역슬래시\ 가 아닌 /슬래시로 폴더 구분
PATH_CHROME = (r"C:/CHROME_ACCOUNTS")

#폴더 내 파일 리스트화
FILE_LIST_RANGE = os.listdir(PATH_CHROME)
FILE_LIST_LNK_RANGE = [file for file in FILE_LIST_RANGE if file.endswith(".lnk")]

CHROME_NUMBER = ""

#파일 실행 함수 정의
def run(path):
    # 디렉토리 설정
    os.chdir(PATH_CHROME)
    # 파일 실행
    os.startfile(path)


#=======================단축키 함수 정의==========================

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

#=======================단축키 함수 부분 끝==========================

#=======================실행 / 반복 부분 ===========================
#몇번부터 실행?
ALL_RANGE = len(FILE_LIST_LNK_RANGE)
n = int(input("몇번부터 시작? 1 ~ %d: " %ALL_RANGE))
print(type(n),n)
n -= 1 #사용자 편의를 위해 0부터 시작이 아닌 1부터 시작

for i in FILE_LIST_LNK_RANGE:
    #특정 파일 링크 - 파일에는 한글 공백 특수기호가 없어야합니다
    if(n < len(FILE_LIST_LNK_RANGE)):
        executeFile = PATH_CHROME + '/' + FILE_LIST_LNK_RANGE[n]
    else:
        break

    #몇번째 크롬계정인지 확인
    CHROME_NUMBER = FILE_LIST_LNK_RANGE[n]
    print(CHROME_NUMBER) #계정 번호 출력

    TIME(1)
    CLEAR_SCREEN(1)
    run(executeFile) #크롬실행
    
    TIME(10) #메타마스크 뜰때까지 대기
    NEW_TAB(1)
    
    #============================메타마스크 계정 생성==========================

    TIME(1.5) #잠시 대기시간을 주자
    ALT_D(1)
    WRITE("https://chrome.google.com/webstore/detail/buster-captcha-solver-for/mpbjkejclgfgadiemmefgebjfooflfhl/related")
    ENTER(1)
    TIME(5)
    
    FULL_SCREEN(1)
    TAB(6)
    ENTER(1) #확장프로그램 설치
    TIME(1)

    TAB(1)
    ENTER(1) #확장프로그램추가
    TIME(5)
    
    ALT_F4(1) #크롬종료
    KILL_CHROME()
    CLEAR_SCREEN(1)
    TIME(1.5)

    #============================메타마스크 계정 생성==========================


    #한번 더 반복
    n += 1

print("Buster자동설치프로그램종료")