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
                    # 장기요양급여수입인 경우 팝업창이 오픈되는데 반영을 기준으로 선택하도록 한다.
                    elif data == '장기요양급여수입':
                        pick_account_반영(data, 'account_subject')
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
                    
                    gui.hotkey('ctrl', 'a')
                    gui.press('backspace')

                    if rows[0] == '수입': pyperclip.copy(rows[6])
                    elif rows[0] == '지출': pyperclip.copy(rows[7])
                    
                    gui.hotkey('ctrl', 'v')
                    continue
                elif i==9:
                    if data != '':
                        img_click('상대계정박스.png')
                        gui.press('tab')
                        pyperclip.copy(data)
                        gui.hotkey('ctrl', 'v')
                        time.sleep(0.2)

                        # '장기요양급여수입' 계정과목같은 경우는, 마음손에서 반영/미반영을 별도로 처리하지 않을 때 상대계정코드목록 팝업창이 오픈하게 된다.
                        # 따라서, 코드/명 항목이 존재하면 인풋에 값을 입력하고 검색한다.
                        if data == '장기요양급여수입':
                            pick_account_반영(data, 'opponent_subject')
                        continue
                    else:
                        continue

            if i == max_col_cnt-1:
                time.sleep(0.5)
                img_left_click('조회.png') # 포커스 초기화 클릭
                time.sleep(0.5)
                gui.hotkey('alt', 'f8') # 저장
                time.sleep(0.5)
                screen_center_click() # 포커스 초기화 클릭
                time.sleep(0.5)
                img_left_click('저장취소.png')
                time.sleep(0.5)                
                screen_center_click() # 포커스 초기화 클릭
                time.sleep(0.5)                
                img_click('성공저장확인.png')                
                time.sleep(0.5)                
                status_change_true(self, row_i) #### 성공 확인됨. Flag값 수정하기
                
                time.sleep(0.5)
            
    except Exception as e:
        logging.debug('전표정보 자동저장(auto_save) Exception : ', e)
# def auto_save End #






# 급여대장 자동업로드 수행
'''
    @param self      # PyQT5
    @param excelList # makeExcelData()를 통해 갖고있는 데이터
'''
def payroll_auto_save(self, excel_list):
    try:
        max_col_cnt = len(excel_list[0])
        
        for row_i in range(0, len(excel_list)) :
            rows = excel_list[row_i]

            gui.hotkey('alt', 'f3') # 조회
            time.sleep(0.5)

            # 사업명 리프레시
            img_right_150_click('급여대장_사업명.png')
            if find_img_flag('급여대장_사업명_선택하세요.png'): 
                img_click('급여대장_사업명_선택하세요.png')

            gui.press('enter') # 팝업창 뜸.. 제거하기 위한 엔터
            time.sleep(0.5)

            img_click('행추가.png')
            
            # 팝업창 오픈
            img_click('선택.png')
            img_click('전체선택체크박스.png')

            # 사업명 선택
            img_right_150_click('급여대장_사업명.png')
            customization_payroll_project_img_click(self) # 사업명 선택
            img_click('지출결의서_등록.png') # 지출결의서 등록

            gui.alert('지출결의서의 회계연도를 설정 후 선택하신 뒤, 해당 안내창의 \'확인\'을 눌러주세요.')

            # 확인을 누르면 다음 매크로 수행

            for i in range(3, max_col_cnt):
                '''
                    급여대장은 지출결의서만 관리하므로 기본으로 세팅되는 구분, 사업구분은 작업하지 않고
                    거래일자부터 작성하면 된다.

                    *** 미리 세팅되는 항목: 사업, 자금원천, 지출, 계좌
                '''
                data = rows[i]

                

                if i==3:
                    img_right_click('급여대장_결의일자.png')
                    # 결의일자 삽입
                    gui.hotkey('ctrl', 'a')
                    gui.press('backspace')
                    pyperclip.copy(data)
                    gui.hotkey('ctrl', 'v')

                    img_click('행추가.png')
                    continue
                elif i==4:
                    img_bottom_right_in_click('급여대장_계정과목_타이틀.png')
                    img_click('급여대장_코드명.png')
                    gui.hotkey('ctrl', 'a')
                    gui.press('backspace')
                    pyperclip.copy(data)
                    gui.hotkey('ctrl', 'v')
                    time.sleep(0.5)

                    #선택

                    gui.press('enter')
                    time.sleep(0.5)

                    for idx in range(0,3):
                        # 다음 항목 활성화
                        gui.press('tab')

                    continue
                elif i==7:
                    img_bottom_right_in_click('급여대장_금액.png')
                    gui.hotkey('ctrl', 'a')
                    gui.press('backspace')
                    pyperclip.copy(data)
                    gui.hotkey('ctrl', 'v')
                    time.sleep(0.5)

                    # 다음 항목(적요) 활성화
                    gui.press('tab')

                    # 만약 적요를 작성할거라면 별도로 처리하던지 여기서 이어서 작성하던지.
                    continue
                elif i==9:
                    img_bottom_right_in_click('급여대장_상대계정과목_타이틀.png')
                    gui.hotkey('ctrl', 'a')
                    gui.press('backspace')
                    pyperclip.copy(data)
                    gui.hotkey('ctrl', 'v')
                    time.sleep(0.5)

                    # 팝업창을 오픈하기 위한 tab
                    gui.press('tab')

                    # 인건비 반영을 기준으로 그냥 엔터 침
                    gui.press('enter') if data == '장기요양급여수입' else None

                    
                    continue
            
            if i == max_col_cnt-1:
                time.sleep(0.5)
                img_click('급여대장_저장.png')
                gui.press('enter') # 저장여부 '확인'
                time.sleep(0.5)
                gui.press('enter') # 성공저장 확인
                time.sleep(0.5)
        time.sleep(0.5)
    except Exception as e:
        logging.debug('급여대장 자동저장(payroll_auto_save) Exception : ', e)
        sys.exit()








############### FUNCTION ############################################################################################################################################
#1 이미지 찾아서 가운데 클릭 기능
def img_click(img_nm):
    img = gui.locateOnScreen(img_dir_path + img_nm)

    if img is not None: 
        center = gui.center(img)
        gui.click(center, interval=0.5)
    else:
        gui.alert(f'찾는 이미지 : {img_nm}\n찾고자 하는 이미지가 존재하지 않습니다. \n관리자 확인이 필요합니다.')
        sys.exit()
# def img_click End #


#1-2 이미지 찾음유무 flag 확인
def find_img_flag(img_nm):
    img = gui.locateOnScreen(img_dir_path + img_nm)
    return True if img is not None else False


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


#2-2 이미지 찾아서 이미지의 오른쪽 위치 50 클릭 기능
def img_right_150_click(img_nm):
    img = gui.locateOnScreen(img_dir_path + img_nm)

    if img is not None:
        moveX = (img.left + img.width) + 150
        moveY = img.top + img.height // 2
        gui.click(moveX, moveY, interval=0.5)
    else:
        gui.alert(f'찾는 이미지 : {img_nm}\n찾고자 하는 이미지가 존재하지 않습니다. \n관리자 확인이 필요합니다.')
        sys.exit()
# def img_right_click End #


# 이미지 찾아서 이미지의 왼쪽 위치 클릭 기능
def img_left_click(img_nm):
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


#8 계정과목코드가 '장기요양급여수입'인 경우 반영을 픽스하기 위한 기능
def pick_account_반영(data, type):
    img_left_click('조회.png') # 포커스 초기화 클릭
    time.sleep(1.0)
    
    if type == 'account_subject':
        img_click('팝업_계정코드목록.png')
        time.sleep(1.0)
        gui.press('enter') # 선택
        time.sleep(0.5)
    elif type == 'opponent_subject':
        img_click('코드명_장기요양급여수입.png')
        gui.hotkey('ctrl', 'a')
        gui.press('backspace')
        pyperclip.copy(data)
        gui.hotkey('ctrl', 'v')
        gui.press('enter')
        time.sleep(1.0)
        gui.press('enter') # 선택
        time.sleep(0.5)



# 급여대장의 사업명 선택
def customization_payroll_project_img_click(self):
    img_path = self.file_payroll_project_img_path.toPlainText()
    img_nm = self.file_payroll_project_img_nm.toPlainText().split('.')[0]

    click_img = gui.locateOnScreen(img_path)

    if click_img is not None:
        center = gui.center(click_img)
        gui.click(center, interval=0.5)
    else:
        gui.alert(f'찾는 이미지 : {img_nm}\n찾고자 하는 이미지가 존재하지 않습니다. \n관리자 확인이 필요합니다.')
        return False
# def customization_payroll_project_img_click End #




def img_bottom_right_in_click(img_nm):
    img = gui.locateOnScreen(img_dir_path + img_nm)

    if img is not None:
        moveX = img.left + img.width - 10
        moveY = img.top + img.height + 10
        gui.click(moveX, moveY, interval=0.5)
    else:
        gui.alert(f'찾는 이미지 : {img_nm}\n찾고자 하는 이미지가 존재하지 않습니다. \n관리자 확인이 필요합니다.')
        sys.exit()
# def img_bottom_right_in_click End #