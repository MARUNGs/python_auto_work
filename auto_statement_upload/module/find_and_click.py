# 특정 이미지를 찾아서 클릭하는 기능

########## import list ##############################################################################################################################
##### Library import 
import pyautogui as gui             # 운영체제 제어
import os



# 공통 경로
img_dir_path = os.path.dirname(__file__).replace('module', 'img') + os.sep
# 기본 딜레이 설정
gui.PAUSE = 0.2


#8# 이미지 찾아서 이동
def move_to_img(img_nm, self):
    img = gui.locateOnScreen(img_dir_path + img_nm)

    if img is not None: 
        center = gui.center(img)
        gui.click(center)
    else:
        gui.alert(f'찾는 이미지 : {img_nm}\n찾고자 하는 이미지가 존재하지 않습니다. \n관리자 확인이 필요합니다.')
# def move_to_img End #


#9# 이미지 더블클릭
def img_db_click(img_nm, self):
    img = gui.locateOnScreen(img_dir_path + img_nm)

    if img is not None: 
        center = gui.center(img)
        gui.doubleClick(center)
    else:
        gui.alert(f'찾는 이미지 : {img_nm}\n찾고자 하는 이미지가 존재하지 않습니다. \n관리자 확인이 필요합니다.')
# def img_db_click End #