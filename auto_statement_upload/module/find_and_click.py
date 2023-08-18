# 특정 이미지를 찾아서 클릭하는 기능
# 이미지를 최대 2번 찾도록 데코레이터 설정을 취한다. 2번째 또한 이미지를 못 찾을 경우엔 에러 취급함.

########## import list ##############################################################################################################################
##### Library import 
import pyautogui as gui             # 운영체제 제어
import os
import sys                          # 시스템 정보
import logging
import pyperclip
from PIL import Image


# 공통 경로
img_dir_path = os.path.abspath(os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'img')) + os.sep
# logger
applogger = logging.getLogger("app")


########## Function ##############################################################################################################################
# 이미지 gray_scale 변환
def gray_scale_img(img_nm):
    gray_image = Image.open(os.path.join(img_dir_path, img_nm)).convert("L")
    confidence = 0.999
    return gui.locateOnScreen(gray_image, grayscale=True, confidence=confidence)


# 이미지 찾아서 클릭
def img_click(img_nm):
    try:
        location_img = gray_scale_img(img_nm) # 이미지 gray_scale 변환
        
        if location_img is not None:
            center = gui.center(location_img)
            gui.click(center)
        else: 
            gui.alert(f'찾는 이미지 : {img_nm}\n찾고자 하는 이미지가 존재하지 않습니다. \n관리자 확인이 필요합니다.')
            applogger.debug('img_click ERROR')
            sys.exit()
    except Exception as e:
        applogger.debug('img_click ERROR : ', str(e))


# 이미지 찾음유무 flag 확인
def find_img_flag(img_nm):
    img = gray_scale_img(img_nm)
    # img = gui.locateOnScreen(img_dir_path + img_nm)
    return True if img is not None else False


# 이미지 더블클릭
def img_db_click(img_nm):
    img = gray_scale_img(img_nm)

    if img is not None: 
        center = gui.center(img)
        gui.doubleClick(center, interval=0.5)
    else: 
        gui.alert(f'찾는 이미지 : {img_nm}\n찾고자 하는 이미지가 존재하지 않습니다. \n관리자 확인이 필요합니다.')
        applogger.debug('img_db_click ERROR')
        sys.exit()


# 이미지의 오른쪽 150px 이동하여 클릭
def img_right_150_click(img_nm):
    img = gray_scale_img(img_nm)

    if img is not None:
        moveX = (img.left + img.width) + 150
        moveY = img.top + img.height // 2
        gui.click(moveX, moveY)
    else:
        gui.alert(f'찾는 이미지 : {img_nm}\n찾고자 하는 이미지가 존재하지 않습니다. \n관리자 확인이 필요합니다.')
        applogger.debug('img_right_150_click ERROR')
        sys.exit()


# 커스텀된 사업명을 찾아서 클릭
def customization_payroll_project_img_click(self):
    img_path = self.file_payroll_project_img_path.toPlainText()
    length   = len(img_path.rsplit(os.sep))
    img_nm   = img_path.rsplit(os.sep)[length-1]
    img      = gray_scale_img(img_nm)

    if img is not None:
        center = gui.center(img)
        gui.click(center)
    else: 
        gui.alert(f'찾는 이미지 : {img_nm}\n찾고자 하는 이미지가 존재하지 않습니다. \n관리자 확인이 필요합니다.')
        applogger.debug('customization_payroll_project_img_click ERROR')
        sys.exit()


# 커스텀된 회계연도를 찾아서 왼쪽 클릭
def customization_payroll_year_img_click(self):
    img_path = self.file_payroll_year_img_path.toPlainText()
    length   = len(img_path.rsplit(os.sep))
    img_nm   = img_path.rsplit(os.sep)[length-1]
    img      = gray_scale_img(img_nm)

    if img is not None:
        moveX = img.left - 10
        moveY = img.top + img.height // 2
        gui.click(moveX, moveY)
    else:
        gui.alert(f'찾는 이미지 : {img_nm}\n찾고자 하는 이미지가 존재하지 않습니다. \n관리자 확인이 필요합니다.')
        applogger.debug('customization_payroll_year_img_click ERROR')
        sys.exit()
    

# 이미지 찾아서 이미지의 오른쪽 위치 클릭 기능
def img_right_click(img_nm):
    img = gray_scale_img(img_nm)

    if img is not None:
        moveX = (img.left + img.width) + 10
        moveY = img.top + img.height // 2
        gui.click(moveX, moveY)
    else:
        gui.alert(f'찾는 이미지 : {img_nm}\n찾고자 하는 이미지가 존재하지 않습니다. \n관리자 확인이 필요합니다.')
        applogger.debug('img_right_click ERROR')
        sys.exit()


# 이미지의 바텀에서 살짝 오른쪽 클릭
def img_bottom_right_in_click(img_nm):
    img = gray_scale_img(img_nm)

    if img is not None:
        moveX = img.left + img.width - 10
        moveY = img.top + img.height + 10
        gui.click(moveX, moveY)
    else:
        gui.alert(f'찾는 이미지 : {img_nm}\n찾고자 하는 이미지가 존재하지 않습니다. \n관리자 확인이 필요합니다.')
        applogger.debug('img_bottom_right_in_click ERROR')
        sys.exit()


###########################################################################
####### 결의서/전표 함수 ###################################################
###########################################################################

# 사용자가 업로드한 사업명 이미지 경로를 찾아서 가운데 클릭
def customization_project_img_click(self):
    img_path = self.file_project_img_path.toPlainText()
    length = len(img_path.rsplit(os.sep))
    img_nm = img_path.rsplit(os.sep)[length-1]

    img = gray_scale_img(img_nm)

    if img is not None:
        center = gui.center(img)
        gui.click(center)
    else:
        gui.alert(f'찾는 이미지: {img_nm}\n찾고자 하는 이미지가 존재하지 않습니다. \n관리자 확인이 필요합니다.')
        applogger.debug('customization_project_img_click ERROR')
        sys.exit()


# 이미지를 찾아서 이미지의 오른쪽 안 끝을 클릭하는 기능
def img_right_in_click(img_nm):
    img = gray_scale_img(img_nm)

    if img is not None:
        moveX = img.left + img.width - 20
        moveY = img.top + img.height // 2
        gui.click(moveX, moveY)
    else:
        gui.alert(f'찾는 이미지 : {img_nm}\n찾고자 하는 이미지가 존재하지 않습니다. \n관리자 확인이 필요합니다.')
        applogger.debug('img_right_in_click ERROR')
        sys.exit()


# 이미지 찾아서 이미지의 왼쪽 위치 클릭하는 기능
def img_left_click(img_nm):
    img = gray_scale_img(img_nm)

    if img is not None:
        moveX = img.left - 30
        moveY = img.top + img.height // 2
        gui.click(moveX, moveY)
    else:
        gui.alert(f'찾는 이미지 : {img_nm}\n찾고자 하는 이미지가 존재하지 않습니다. \n관리자 확인이 필요합니다.')
        applogger.debug('img_left_click ERROR')
        sys.exit()


# 계정과목코드가 '장기요양급여수입'인 경우 반영을 픽스하기 위한 기능
# def pick_account_반영(data, type, self):
#     xy_info = self.img_xy_info
    
#     if type == 'account_subject':
#         xy_info_click(xy_info['popup_account_cd']) # 팝업_계정코드목록
#         gui.hotkey('ctrl', 'a')
#         gui.press('backspace')
#         pyperclip.copy(data)
#         gui.hotkey('ctrl', 'v')
#         gui.press('enter')
#         gui.press('enter')
#     elif type == 'opponent_subject':
#         xy_info_click(*xy_info['popup_opponent_account_subject_cd']) # 상대계정 코드명 입력
#         gui.hotkey('ctrl', 'a')
#         gui.press('backspace')
#         pyperclip.copy(data)
#         gui.hotkey('ctrl', 'v')
#         gui.press('enter')
#         gui.press('enter')

def pick_account_반영(data, type):
    img_left_click('조회.png') # 포커스 초기화 클릭
    gui.sleep(1.0)
    
    if type == 'account_subject':
        img_click('팝업_계정코드목록.png')
        gui.sleep(1.0)
        gui.press('enter') # 선택
        gui.sleep(0.5)
    elif type == 'opponent_subject':
        img_click('코드명_장기요양급여수입.png')
        gui.hotkey('ctrl', 'a')
        gui.press('backspace')
        pyperclip.copy(data)
        gui.hotkey('ctrl', 'v')
        gui.press('enter')
        gui.sleep(1.0)
        gui.press('enter') # 선택
        gui.sleep(0.5)


# 화면 정 가운데 클릭
def screen_center_click():
    # 화면 크기 가져오기
    screen_width, screen_height = gui.size()

    # 정 가운데 좌표 계산
    click_x = screen_width // 2
    click_y = screen_height // 2

    # 클릭 실행
    gui.click(click_x, click_y)


def xy_info_click(xy_info):
    gui.moveTo(*xy_info) # key값의 xy좌표로 이동
    gui.click()          # 이동한 좌표에서 클릭