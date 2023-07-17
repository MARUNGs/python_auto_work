# 결의서/전표 자동저장 작업
import os
import pyautogui as gui                          # 운영체제 제어
from PyQt5.QtWidgets import QTableWidgetItem
import pyperclip                                 # 데이터 복사 및 붙여넣기 
import time
import logging                                   # 로그
import sys                                       # 시스템 정보


# 공통 경로
img_dir_path = os.path.dirname(__file__).replace('module', 'img') + os.sep


'''
    @param self      # PyQT5
    @param excelList # makeExcelData()를 통해 갖고있는 데이터
'''
def auto_save(self, excel_list):
    try:
        max_col_cnt = len(excel_list[0])
        
        for row_i in range(0, len(excel_list)) :
            rows = excel_list[row_i]

            gui.hotkey('alt', 'f3') # 조회
            time.sleep(0.5)
            gui.hotkey('alt', 'f2') # 신규
            time.sleep(0.5)
            
            for i in range(0, max_col_cnt):
                data = rows[i]

                if i==0:
                    img_right_click('결의구분타이틀.png')
                    if data == '수입': 
                        img_click('수입결의서TXT.png')
                    else: 
                        img_click('지출결의서TXT.png')
                    continue
                elif i==2:
                    img_right_click('사업타이틀.png')
                    customization_project_img_click(self)
                    continue
                elif i==3:
                    # 결의일자 활성화
                    img_right_click('결의일자타이틀.png')
                    # 결의일자 삽입
                    gui.hotkey('ctrl', 'a')
                    gui.press('backspace')
                    pyperclip.copy(data)
                    gui.hotkey('ctrl', 'v')
                    continue
                elif i==4:
                    img_click('계정코드박스.png')
                    gui.press('tab')
                    pyperclip.copy(data)
                    gui.hotkey('ctrl', 'v')
                    img_click('계정코드박스.png') # 포커스 초기화하여 '대상자' 항목을 나타내기 위함

                    # 본인부담금수입인 경우, 대상자 항목이 오픈되고 가장 첫번째 대상자를 선택한다.
                    if data == '본인부담금수입':
                        img_right_in_click('대상자.png')
                        time.sleep(1)
                        gui.press('enter')

                    continue
                elif i==5:
                    img_click('결의서적요.png')
                    gui.hotkey('ctrl', 'a')
                    gui.press('backspace')
                    pyperclip.copy(data)
                    gui.hotkey('ctrl', 'v')
                    continue
                elif i==6 or i==7:
                    img_right_click('금액타이틀.png')
                    # 수입이던지, 지출이던지간에 어차피 금액은 다시 세팅될것이므로 그냥 두자...
                    gui.hotkey('ctrl', 'a')
                    gui.press('backspace')
                    pyperclip.copy(data)
                    gui.hotkey('ctrl', 'v')
                    continue
                elif i==9:
                    if data != '':
                        img_click('상대계정박스.png')
                        gui.press('tab')
                        pyperclip.copy(data)
                        gui.hotkey('ctrl', 'v')
                        time.sleep(0.2)
                        continue
                    else:
                        continue

            if i == max_col_cnt-1:
                time.sleep(0.5)
                imgLeftClick('조회.png') # 포커스 초기화 클릭
                time.sleep(0.5)
                gui.hotkey('alt', 'f8') # 저장
                time.sleep(0.5)
                screen_center_click() # 포커스 초기화 클릭
                time.sleep(0.5)
                imgLeftClick('저장취소.png')
                time.sleep(0.5)                
                screen_center_click() # 포커스 초기화 클릭
                time.sleep(0.5)                
                img_click('성공저장확인.png')                
                time.sleep(0.5)                
                status_change_true(self, row_i) #### 성공 확인됨. Flag값 수정하기
                
                time.sleep(0.5)
            
    except Exception as e:
        logging.debug('자동저장(auto_save) Exception : ', e)
# def auto_save End #



#1 이미지 찾아서 가운데 클릭 기능
def img_click(img_nm):
    img = gui.locateOnScreen(img_dir_path + img_nm)

    if img is not None: 
        center = gui.center(img)
        gui.click(center)
    else:
        gui.alert(f'찾는 이미지 : {img_nm}\n찾고자 하는 이미지가 존재하지 않습니다. \n관리자 확인이 필요합니다.')
        sys.exit()
# def img_click End #


#2 이미지 찾아서 이미지의 오른쪽 위치 클릭 기능
def img_right_click(img_nm):
    img = gui.locateOnScreen(img_dir_path + img_nm)

    if img is not None:
        moveX = (img.left + img.width) + 10
        moveY = img.top + img.height // 2
        gui.click(moveX, moveY)
    else:
        gui.alert(f'찾는 이미지 : {img_nm}\n찾고자 하는 이미지가 존재하지 않습니다. \n관리자 확인이 필요합니다.')
        sys.exit()
# def img_right_click End #


# 이미지 찾아서 이미지의 왼쪽 위치 클릭 기능
def imgLeftClick(img_nm):
    img = gui.locateOnScreen(img_dir_path + img_nm)

    if img is not None:
        moveX = img.left - 30
        moveY = img.top + img.height // 2
        gui.click(moveX, moveY)
    else:
        gui.alert(f'찾는 이미지 : {img_nm}\n찾고자 하는 이미지가 존재하지 않습니다. \n관리자 확인이 필요합니다.')
        sys.exit()


#3 사용자가 올린 사업명 이미지 경로를 찾아서 가운데 클릭
'''
    @param self : PyQt5
'''
def customization_project_img_click(self):
    img_path = self.file_project_img_path.toPlainText()
    img_nm = self.file_project_img_nm.toPlainText().split('.')[0]

    click_img = gui.locateOnScreen(img_path)

    if click_img is not None: 
        center = gui.center(click_img)
        gui.click(center)
    else:
        gui.alert(f'찾는 이미지 : {img_nm}\n찾고자 하는 이미지가 존재하지 않습니다. \n관리자 확인이 필요합니다.')
        return False


#4 (사용안함, 계좌 관련된 내용 삭제예정)사용자가 올린 계좌명 이미지 경로를 찾아서 가운데 클릭
def customization_manage_img_click(self):
    find_manage_img = gui.locateOnScreen(img_dir_path + '계좌번호(선택).png')
    img_nm = self.file_project_img_nm.toPlainText().split('.')[0]

    # 만약, 사업이 세팅되어 계좌번호가 자동적으로 세팅되어 있지 않다면 이미지를 찾아서 클릭할 것.
    if find_manage_img == None:
        img_path = self.file_manage_img_path.toPlainText()
        img_right_click('계좌번호타이틀.png')
        click_img = gui.locateOnScreen(img_path)

        if click_img is not None:
            center = gui.center(click_img)
            gui.click(center)
        else:
            gui.alert(f'찾는 이미지 : {img_nm}\n찾고자 하는 이미지가 존재하지 않습니다. \n관리자 확인이 필요합니다.')
            sys.exit()
    else:
        gui.alert(f'찾는 이미지 : {img_nm}\n찾고자 하는 이미지가 존재하지 않습니다. \n관리자 확인이 필요합니다.')
        sys.exit()
            
 

#5 화면 정 가운데 클릭
def screen_center_click():
    # 화면 크기 가져오기
    screen_width, screen_height = gui.size()

    # 정 가운데 좌표 계산
    click_x = screen_width // 2
    click_y = screen_height // 2

    # 클릭 실행
    gui.click(click_x, click_y)


#6 상태 테이블값 true로 변환
def status_change_true(self, row_i): self.status_tb.setItem(row_i, 0, QTableWidgetItem('Success'))


#7 이미지를 찾아서 이미지의 오른쪽 끝을 클릭하는 기능
def img_right_in_click(img_nm):
    img = gui.locateOnScreen(img_dir_path + img_nm)

    if img is not None:
        moveX = img.left + img.width - 20
        moveY = img.top + img.height // 2
        gui.click(moveX, moveY)
    else:
        gui.alert(f'찾는 이미지 : {img_nm}\n찾고자 하는 이미지가 존재하지 않습니다. \n관리자 확인이 필요합니다.')
        sys.exit()