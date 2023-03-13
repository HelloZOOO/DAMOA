import os

import random

# 크롬 계정데이터 모여있는 폴더 chrome://version
directory_chrome = r"C:\Users\thqud\AppData\Local\Google\Chrome\User Data"

# 해당 문자가 포함된 폴더만 리스트화 시키겠다 (프로필 뒤에 공백 필수)
include_text = 'Profile '

# 디렉토리명 필터링
DIR_PROFILE_LIST = (
    [d for d in os.listdir(directory_chrome) 
    if os.path.isdir(os.path.join(directory_chrome, d)) and include_text in d]
)

#난수생성기
def RANDOM_NUM():
    num = round(random.uniform(5, 15), 1)
    return num

#폴더명 리스트화
print(len(DIR_PROFILE_LIST))
for i in range(len(DIR_PROFILE_LIST)):
    print(DIR_PROFILE_LIST[i])
