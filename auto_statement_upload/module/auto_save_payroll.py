# 급여대장 자동업로드 작업
import os
import pyautogui as gui                          # 운영체제 제어
from PyQt5.QtWidgets import QTableWidgetItem
import pyperclip                                 # 데이터 복사 및 붙여넣기 
import time
import logging                                   # 로그
import sys                                       # 시스템 정보

from . import find_and_click


# 공통 경로
img_dir_path = os.path.dirname(__file__).replace('module', 'img') + os.sep
# logger
applogger = logging.getLogger("app")


'''
    @param self       # PyQT5
    @param excel_list # makeExcelData()를 통해 갖고있는 데이터
'''
# def auto_save_payroll(self, excel_list):
#     try:
#         xy_info = self.img_xy_info # x,y 좌표 정보
#         max_col_cnt = len(excel_list[0]) if len(excel_list) > 0 else 0
#         for row_i in range(0, len(excel_list)) :
#             rows = excel_list[row_i]

#             gui.hotkey('alt', 'f3') # 조회

#             # 사업명 리프레시
#             find_and_click.img_right_150_click('급여대장_사업명.png') #find_and_click.xy_info_click(xy_info['payroll_project_combobox'])
#             find_and_click.img_click('급여대장_사업명_선택하세요.png') if find_and_click.find_img_flag('급여대장_사업명_선택하세요.png') else None # find_and_click.xy_info_click(xy_info['payroll_project_refresh']) if find_and_click.find_img_flag('급여대장_사업명_선택하세요.png') else None
#             gui.press('enter') # 팝업창 뜸.. 제거하기 위한 엔터
#             time.sleep(0.5)

#             find_and_click.img_click('행추가.png') #find_and_click.xy_info_click(xy_info['payroll_add_row_btn'])
            
#             # 팝업창 오픈
#             gui.press('tab')
#             gui.press('enter') # find_and_click.img_click('선택.png') #find_and_click.xy_info_click(xy_info['payroll_staff_select_btn'])
#             find_and_click.img_click('전체선택체크박스.png') #find_and_click.xy_info_click(xy_info['payroll_all_check'])

#             # 사업명 선택
#             find_and_click.img_right_150_click('급여대장_사업명.png') #find_and_click.xy_info_click(xy_info['payroll_project_combobox'])
#             find_and_click.customization_payroll_project_img_click(self) # 사업명 선택
#             gui.press('tab') 
#             gui.press('enter') # find_and_click.img_click('지출결의서_등록.png') #find_and_click.xy_info_click(xy_info['payroll_expense_registration']) # 지출결의서 등록

#             # 회계연도의 왼쪽 클릭
#             find_and_click.customization_payroll_year_img_click(self)
#             find_and_click.img_click('회계연도_선택.png') #find_and_click.xy_info_click(xy_info['payroll_account_year_select_btn']) # 회계연도 선택

#             time.sleep(1.0)

#             for i in range(3, max_col_cnt):
#                 '''
#                     급여대장은 지출결의서만 관리하므로 기본으로 세팅되는 구분, 사업구분은 작업하지 않고
#                     거래일자부터 작성하면 된다.

#                     *** 미리 세팅되는 항목: 사업, 자금원천, 지출, 계좌
#                 '''
#                 data = rows[i]

#                 #결의일자#
#                 if i==3:
#                     find_and_click.img_right_click('급여대장_결의일자.png') #find_and_click.xy_info_click(xy_info['payroll_cashier_dt'])
#                     # 결의일자 삽입
#                     gui.hotkey('ctrl', 'a')
#                     gui.press('backspace')
#                     pyperclip.copy(data)
#                     gui.hotkey('ctrl', 'v')

#                     find_and_click.img_click('행추가.png') #find_and_click.xy_info_click(xy_info['payroll_statement_add_row_btn'])
#                     continue
#                 #계정명#
#                 elif i==4:
#                     # tab으로 하면 안내메세지가 발생하여 계정과목은 이미지로 찾음.
#                     find_and_click.img_bottom_right_in_click('급여대장_계정과목_타이틀.png')

#                     for idx in range(0, 4): gui.press('tab') # find_and_click.img_click('급여대장_코드명.png')
#                     gui.hotkey('ctrl', 'a')
#                     gui.press('backspace')
#                     pyperclip.copy(data)
#                     gui.hotkey('ctrl', 'v')
#                     for idx in range(0, 2): gui.press('enter')

#                     # 다음 항목 활성화
#                     for idx in range(0,3): gui.press('tab')
#                     continue
#                 #금액#
#                 elif i==7:
#                     gui.hotkey('ctrl', 'a')
#                     gui.press('backspace')
#                     pyperclip.copy(data)
#                     gui.hotkey('ctrl', 'v')
#                     gui.sleep(0.2)
#                     for idx in range(0,2): gui.press('tab')
#                     continue
#                 #상대계정#
#                 elif i==9:
#                     gui.hotkey('ctrl', 'a')
#                     gui.press('backspace')
#                     pyperclip.copy(data)
#                     gui.hotkey('ctrl', 'v')
#                     gui.press('tab') # 팝업창을 오픈하기 위한 tab

#                     gui.sleep(1.0)
#                     # 인건비 반영을 기준으로 그냥 엔터 침
#                     if data == '장기요양급여수입': gui.press('enter')
#                     continue
            
#             if i == max_col_cnt-1:
#                 find_and_click.img_click('급여대장_저장.png') #find_and_click.xy_info_click(xy_info['payroll_save'])
#                 for idx in range(0, 2):
#                     gui.press('enter') # 팝업창 확인
#                     gui.sleep(0.2)
#                 # gui.press('enter') # 저장여부 '확인'
#                 # time.sleep(0.2)
#                 # gui.press('enter') # 성공저장 확인                time.sleep(0.2)

#                 # 급여대장 지출결의서 닫기
#                 for idx in range(0, 5): gui.press('tab')
#                 gui.press('enter')

#                 # find_and_click.img_click('급여대장_닫기.png') #find_and_click.xy_info_click(xy_info['close'])
#                 status_change_true(self, row_i) #### 성공 확인됨. Flag값 수정하기

#         status_change_true(self, len(excel_list)) #### 사실 여기까지 왔으면 성공취급함. 마지막줄 Flag값 수정하기
#     except Exception as e:
#         applogger.debug('auto save statement ERROR MSG : ', str(e))
#         sys.exit()


#### 리팩토링
def auto_save_payroll(self, excel_list):
    try:
        # 사업을 딱 1번만 설정함. 어차피 같은 사업처리할 것이기 때문.
        find_and_click.img_click('급여대장_사업명.png')
        gui.press('tab')

        if find_and_click.find_img_flag('급여대장_사업명_선택하세요2.png'):
            project_num = int(self.project_num.text())
            for _ in range(0, project_num): 
                gui.press('down')
                gui.press('enter') # 안내창 닫기
        
        gui.sleep(0.5)

        #데이터 수 만큼 움직이기#
        for row_i in range(0, len(excel_list)):
            rows = excel_list[row_i]

            gui.hotkey('alt', 'f3') # 조회

            #행추가#
            find_and_click.img_click('급여대장_사업명.png')
            gui.keyDown('shift')
            for _ in range(0, 4): gui.press('tab')
            gui.keyUp('shift')
            gui.press('enter')

            gui.sleep(0.7) # 팝업창 오픈

            #대상자 선택버튼 클릭#
            gui.press('tab')
            gui.press('enter')

            #전체 선택박스 체크#
            find_and_click.img_click('전체선택체크박스.png')

            # 위에서 먼저 설정했으므로 지금은 필요없음
            # #사업 선택#
            # find_and_click.img_right_150_click('급여대장_사업명.png')
            # find_and_click.customization_payroll_project_img_click(self) # 사업명 선택
            # gui.sleep(0.5)

            #회계연도 선택#
            find_and_click.img_click('급여대장_사업명.png')
            for _ in range(0, 2): gui.press('tab')
            gui.press('enter')
            find_and_click.customization_payroll_year_img_click(self)
            find_and_click.img_click('회계연도_선택.png')

            #기존소스 #회계연도 선택#
            # gui.press('tab')
            # gui.press('enter')
            # find_and_click.customization_payroll_year_img_click(self)
            # find_and_click.img_click('회계연도_선택.png')

            time.sleep(1.0) #팝업창 오픈
            
            #결의일자#
            for _ in range(0, 6): gui.press('tab')
            gui.hotkey('ctrl', 'a')
            gui.press('backspace')
            pyperclip.copy(rows[3])
            gui.hotkey('ctrl', 'v')

            #행추가#
            for _ in range(0, 10): gui.press('tab')
            gui.press('enter')

            #계정명 : 직접 데이터를 입력하면 안내창이 뜨므로, 아이콘을 눌러서 처리해야 한다.#
            find_and_click.img_bottom_right_in_click('급여대장_계정과목_타이틀.png')
            for _ in range(0, 4): gui.press('tab')
            pyperclip.copy(rows[4])
            gui.hotkey('ctrl', 'v')
            for _ in range(0, 2): gui.press('enter') # 확인, 창닫기

            #금액#
            for _ in range(0, 3): gui.press('tab')
            pyperclip.copy(rows[7]) # 지출금액만 취급하므로 idx 7번만 사용함.
            gui.hotkey('ctrl', 'v')

            #상대계정#
            for _ in range(0,2): gui.press('tab')
            pyperclip.copy(rows[9]) # 지출금액만 취급하므로 idx 7번만 사용함.
            gui.hotkey('ctrl', 'v')
            gui.press('tab')

            gui.sleep(0.5)
            if rows[9] == '장기요양급여수입': gui.press('enter')
            gui.sleep(1.0)


            #1건 마무리 프로세스#
            find_and_click.img_click('급여대장_저장.png')
            for _ in range(0, 2):
                gui.press('enter') # 팝업창 확인
                gui.sleep(0.2)
            # 급여대장 지출결의서 닫기
            for _ in range(0, 5): gui.press('tab')
            gui.press('enter')

            #성공 확인됨. Flag값 수정하기#
            status_change_true(self, rows) 

        # status_change_true(self, len(excel_list)) #### 사실 여기까지 왔으면 성공취급함. 마지막줄 Flag값 수정하기
    except Exception as e: applogger.debug('auto save statement ERROR MSG : ', str(e))


# 상태 테이블값 true로 변환
def status_change_true(self, rows):
    # 기존소스 # self.status_tb.setItem(row_i, 0, QTableWidgetItem('Success'))
    data      = rows
    excel_tb  = self.excel_tb
    status_tb = self.status_tb

    # excel_tb 위젯을 for문 돌아서 결의번호(10)가 동일하면 그 index를 이용하여 status_tb의 fail를 success로 변경하기
    for r_idx in range(status_tb.rowCount()): 
        if data[10] == excel_tb.item(r_idx, 10).text(): status_tb.setItem(r_idx, 0, QTableWidgetItem('Success'))