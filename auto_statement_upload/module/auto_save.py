# 결의서/전표 자동저장 작업
import os
import pyautogui as gui                          # 운영체제 제어
from PyQt5.QtWidgets import QTableWidgetItem
import pyperclip                                 # 데이터 복사 및 붙여넣기 
import time
import logging                                   # 로그

from . import find_and_click
from . import auto_save_payroll


# 공통 경로
img_dir_path = os.path.dirname(__file__).replace('module', 'img') + os.sep
# logger
applogger = logging.getLogger("app")

'''
    @param self      # PyQT5
    @param excel_obj # makeExcelData()를 통해 갖고있는 데이터(map 형태)
    ▶      excel_obj = {
                'title_list': [],             # 타이틀
                'income_list': [],            # 수입
                'expense_list': [],           # 지출
                'personnel_expense_list': []  # 인건비(지출)
            }
'''
def auto_save(self, excel_obj):
    try:
        ### 입력해야 할 엑셀 데이터들을 변수로 선언
        income_list = excel_obj['income_list']
        expense_list = excel_obj['expense_list']
        personnel_expense_list = excel_obj['personnel_expense_list']

        ### 화면 이동 : 데이터가 존재하는지 먼저 확인하고 일반전표 데이터가 있으면 간편입력으로 화면이동
        if len(income_list)  > 0 or len(expense_list) > 0: 
            find_and_click.img_click('회계.png')
            find_and_click.img_click('결의및전표관리.png')
            find_and_click.img_click('결의서전표간편입력.png')
            gui.press('enter')

            ### 일반전표(수입, 지출) 처리
            auto_save_simple(self, income_list)
            gui.sleep(2)
            auto_save_simple(self, expense_list)
            gui.sleep(2)

        if len(personnel_expense_list) > 0:
            find_and_click.img_click('간편입력.png')
            find_and_click.img_click('급여대장등록.png')
            gui.press('enter')
            time.sleep(1)

            ### 인건비(지출) 처리
            auto_save_payroll.auto_save_payroll(self, personnel_expense_list)
        
        # Success 갯수 체크
        status_tb     = self.status_tb
        success_count = 0
        fail_count    = 0
        for idx in range(0, status_tb.rowCount()):
            if status_tb.item(idx,0).text() in 'Success': success_count += 1
            else:                                         fail_count += 1

        gui.alert(f'전표정보 자동업로드 등록이 완료되었습니다. \n성공횟수: {success_count} \n실패횟수: {fail_count}')
    except Exception as e:
        logging.error('전표정보 자동저장(auto_save) Exception : ', str(e))



##### 성능을 위하여 리팩토링
def auto_save_simple(self, excel_list):
    global loop_flag
    

    try:
        for row_i in range(0, len(excel_list)) :
            # for문 대신 컨트롤할 변수
            rows = excel_list[row_i]

            for _ in range(0, 2): gui.hotkey('alt', 'f3') # 조회
            for _ in range(0, 2): gui.hotkey('alt', 'f2') # 신규
            
            #결의일자#
            gui.hotkey('ctrl', 'a')
            gui.press('backspace')
            pyperclip.copy(rows[3])
            gui.hotkey('ctrl', 'v')

            for _ in range(0, 2): gui.press('tab')
            
            #결의구분#
            if rows[0] == '수입':
                gui.press('down')
            else: 
                gui.press('down')
                gui.press('down')

            gui.press('tab')

            #사업 ---> 이거 하려면 화면단부터 수정해야 함#
            project_num = int(self.project_num.text())
            if project_num != None: 
                for _ in range(0, project_num): gui.press('down')

            gui.press('tab')

            #금액
            # - 반납구분이 '반납'인 경우 '-' 붙여서 금액 입력할 것.
            gui.hotkey('ctrl', 'a')
            gui.press('backspace')
            
            # '수입'일 때
            if rows[0] == '수입':
                ## '반납'인 경우, 마이너스 금액으로 작성
                if rows[1] == '반납': pyperclip.copy('-' + rows[7])
                ## '정상'인 경우, 금액 그대로 작성
                else: pyperclip.copy(rows[6])
            # '지출'일 때
            else:
                ## '반납'인 경우, 마이너스 금액으로 작성
                if rows[1] == '반납': pyperclip.copy('-' + rows[6])
                ## '정상'인 경우, 금액 그대로 작성
                else: pyperclip.copy(rows[7])
            gui.hotkey('ctrl', 'v')
            
            #계정코드#
            for _ in range(0, 3): gui.press('tab')
            gui.press('enter') # 팝업창 오픈
            for _ in range(0, 4): gui.press('tab')

            gui.sleep(1.0)

            #계정코드명 수정: W4C 프로그램에서는 '공공요금 및 제세공과금'에 대한 과목명이 다르게 관리되므로 별도로 변경함
            if rows[4] == '공공요금 및 제세공과금':
                pyperclip.copy('공공요금 및 각종 세금공과금')
            else:
                pyperclip.copy(rows[4])

            gui.hotkey('ctrl', 'v') # 계정코드 입력
            for _ in range(0, 2): gui.press('enter')

            #대상자#
            if rows[4] == '본인부담금수입':
                for _ in range(0, 3): gui.press('tab')
                for _ in range(0, 2): gui.press('enter') # 팝업창 오픈, 선택까지

            gui.sleep(1.0)

            #상대계정(지출결의서)#
            if rows[9] != '':
                for _ in range(0, 3): gui.press('tab')
                gui.press('enter') # 팝업창 오픈

                gui.sleep(1.0)

                for _ in range(0, 3): gui.press('tab')
                pyperclip.copy(rows[9])
                gui.hotkey('ctrl', 'v')
                for _ in range(0, 2): gui.press('enter') # 무조건 첫번째 상대계정 선택

            #적요#
            if len(rows[5]) > 0:
                for _ in range(0, 6): gui.press('tab')
                pyperclip.copy(rows[5])
                gui.hotkey('ctrl', 'v')
                gui.press('tab')

            gui.sleep(0.5)

            #1건 저장 프로세스#
            gui.hotkey('alt', 'f8') # 저장
            gui.sleep(0.2)
            for _ in range(0, 2):
                gui.press('enter')
                gui.sleep(0.2)
            
            #성공 확인됨. Flag값 수정하기#
            status_change_true(self, rows)
                
    except Exception as e: applogger.debug('auto save statement ERROR MSG : ', str(e))




# 상태 테이블값 true로 변환
def status_change_true(self, rows):
    data      = rows
    excel_tb  = self.excel_tb
    status_tb = self.status_tb

    # excel_tb 위젯을 for문 돌아서 결의번호(10)가 동일하면 그 index를 이용하여 status_tb의 fail를 success로 변경하기
    for r_idx in range(status_tb.rowCount()):
        if data[10] == excel_tb.item(r_idx, 10).text(): status_tb.setItem(r_idx, 0, QTableWidgetItem('Success'))