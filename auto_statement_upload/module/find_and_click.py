# 특정 이미지를 찾아서 클릭하는 기능

########## import list ##############################################################################################################################
##### Library import 
import pyautogui as gui             # 운영체제 제어
import os
import sys                          # 시스템 정보


# 공통 경로
img_dir_path = os.path.dirname(__file__).replace('module', 'img') + os.sep
# 기본 딜레이 설정
gui.PAUSE = 0.2


########## Function ##############################################################################################################################
# 이미지 찾아서 클릭
def img_click(img_nm):
    img = gui.locateOnScreen(img_dir_path + img_nm)

    if img is not None: 
        center = gui.center(img)
        gui.click(center, interval=0.5)
    else:
        gui.alert(f'찾는 이미지 : {img_nm}\n찾고자 하는 이미지가 존재하지 않습니다. \n관리자 확인이 필요합니다.')


# 이미지 찾음유무 flag 확인
def find_img_flag(img_nm):
    img = gui.locateOnScreen(img_dir_path + img_nm)
    return True if img is not None else False


# 이미지 더블클릭
def img_db_click(img_nm):
    img = gui.locateOnScreen(img_dir_path + img_nm)

    if img is not None: 
        center = gui.center(img)
        gui.doubleClick(center, interval=0.5)
    else:
        gui.alert(f'찾는 이미지 : {img_nm}\n찾고자 하는 이미지가 존재하지 않습니다. \n관리자 확인이 필요합니다.')


# 이미지의 오른쪽 150px 이동하여 클릭
def img_right_150_click(img_nm):
    img = gui.locateOnScreen(img_dir_path + img_nm)

    if img is not None:
        moveX = (img.left + img.width) + 150
        moveY = img.top + img.height // 2
        gui.click(moveX, moveY, interval=0.5)
    else:
        gui.alert(f'찾는 이미지 : {img_nm}\n찾고자 하는 이미지가 존재하지 않습니다. \n관리자 확인이 필요합니다.')
        sys.exit()


# 커스텀된 사업명을 찾아서 클릭
def customization_payroll_project_img_click(self):
    img_path = self.file_payroll_project_img_path.toPlainText()
    length = len(img_path.rsplit(os.sep))
    img_nm = img_path.rsplit(os.sep)[length-1]

    click_img = gui.locateOnScreen(img_path)

    if click_img is not None:
        center = gui.center(click_img)
        gui.click(center, interval=0.5)
    else:
        gui.alert(f'찾는 이미지 : {img_nm}\n찾고자 하는 이미지가 존재하지 않습니다. \n관리자 확인이 필요합니다.')
        return False
    

# 이미지 찾아서 이미지의 오른쪽 위치 클릭 기능
def img_right_click(img_nm):
    img = gui.locateOnScreen(img_dir_path + img_nm)

    if img is not None:
        moveX = (img.left + img.width) + 10
        moveY = img.top + img.height // 2
        gui.click(moveX, moveY)
    else:
        gui.alert(f'찾는 이미지 : {img_nm}\n찾고자 하는 이미지가 존재하지 않습니다. \n관리자 확인이 필요합니다.')
        sys.exit()



# 이미지의 바텀에서 살짝 오른쪽 클릭
def img_bottom_right_in_click(img_nm):
    img = gui.locateOnScreen(img_dir_path + img_nm)

    if img is not None:
        moveX = img.left + img.width - 10
        moveY = img.top + img.height + 10
        gui.click(moveX, moveY, interval=0.5)
    else:
        gui.alert(f'찾는 이미지 : {img_nm}\n찾고자 하는 이미지가 존재하지 않습니다. \n관리자 확인이 필요합니다.')
        sys.exit()