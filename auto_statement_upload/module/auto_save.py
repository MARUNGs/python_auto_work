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
            time.sleep(0.5)

            ### 일반전표(수입, 지출) 처리
            auto_save_simple(self, income_list)
            time.sleep(1)
            auto_save_simple(self, expense_list)


        if len(personnel_expense_list) > 0:
            find_and_click.img_click('간편입력.png')
            find_and_click.img_click('급여대장등록.png')
            gui.press('enter')
            time.sleep(1)

            ### 인건비(지출) 처리
            auto_save_payroll.auto_save_payroll(self, personnel_expense_list)
            
        status_tb     = self.payroll_status_tb
        success_count = 0
        fail_count    = 0
        for idx in range(0, status_tb.rowCount()):
            if status_tb.item(idx,0).text() in 'Success': success_count += 1
            else:                                         fail_count += 1

        gui.alert(f'전표정보 자동업로드 등록이 완료되었습니다. \n성공횟수: {success_count} \n실패횟수: {fail_count}')
    except Exception as e:
        logging.error('전표정보 자동저장(auto_save) Exception : ', str(e))


'''
    @param self      # PyQT5
    @param excelList # makeExcelData()를 통해 갖고있는 데이터
'''
def auto_save_simple(self, excel_list):
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
                    find_and_click.img_right_click('결의구분타이틀.png')
                    find_and_click.img_click('수입결의서TXT.png') if data == '수입' else find_and_click.img_click('지출결의서TXT.png')
                    continue
                elif i==2:
                    find_and_click.img_right_click('사업타이틀.png')
                    find_and_click.customization_project_img_click(self)
                    continue
                elif i==3:
                    # 결의일자 활성화
                    find_and_click.img_right_click('결의일자타이틀.png')
                    # 결의일자 삽입
                    gui.hotkey('ctrl', 'a')
                    gui.press('backspace')
                    pyperclip.copy(data)
                    gui.hotkey('ctrl', 'v')
                    continue
                elif i==4:
                    find_and_click.img_click('계정코드박스.png')
                    gui.press('tab')
                    pyperclip.copy(data)
                    gui.hotkey('ctrl', 'v')
                    find_and_click.img_click('계정코드박스.png') # 포커스 초기화하여 '대상자' 항목을 나타내기 위함

                    # 본인부담금수입인 경우, 대상자 항목이 오픈되고 가장 첫번째 대상자를 선택한다.
                    if data == '본인부담금수입':
                        find_and_click.img_right_in_click('대상자.png')
                        time.sleep(1)
                        gui.press('enter')
                    # 장기요양급여수입인 경우 팝업창이 오픈되는데 반영을 기준으로 선택하도록 한다.
                    elif data == '장기요양급여수입':
                        find_and_click.pick_account_반영(data, 'account_subject')
                    continue
                elif i==5:
                    find_and_click.img_click('결의서적요.png')
                    gui.hotkey('ctrl', 'a')
                    gui.press('backspace')
                    pyperclip.copy(data)
                    gui.hotkey('ctrl', 'v')
                    continue
                elif i==6 or i==7:
                    find_and_click.img_right_click('금액타이틀.png')
                    
                    gui.hotkey('ctrl', 'a')
                    gui.press('backspace')

                    if rows[0] == '수입': pyperclip.copy(rows[6])
                    elif rows[0] == '지출': pyperclip.copy(rows[7])
                    
                    gui.hotkey('ctrl', 'v')
                    continue
                elif i==9:
                    if data != '':
                        find_and_click.img_click('상대계정박스.png')
                        gui.press('tab')
                        pyperclip.copy(data)
                        gui.hotkey('ctrl', 'v')
                        time.sleep(0.2)

                        # '장기요양급여수입' 계정과목같은 경우는, 마음손에서 반영/미반영을 별도로 처리하지 않을 때 상대계정코드목록 팝업창이 오픈하게 된다.
                        # 따라서, 코드/명 항목이 존재하면 인풋에 값을 입력하고 검색한다.
                        find_and_click.pick_account_반영(data, 'opponent_subject') if data == '장기요양급여수입' else None
                        # if data == '장기요양급여수입':
                        #     find_and_click.pick_account_반영(data, 'opponent_subject')
                        continue
                    else:
                        continue

            if i == max_col_cnt-1:
                time.sleep(0.5)
                find_and_click.img_left_click('조회.png') # 포커스 초기화 클릭
                time.sleep(0.5)
                gui.hotkey('alt', 'f8') # 저장
                time.sleep(0.5)
                # find_and_click.screen_center_click() # 포커스 초기화 클릭
                # time.sleep(0.5)
                gui.press('enter')
                # find_and_click.img_left_click('저장취소.png')
                time.sleep(0.5)
                # find_and_click.screen_center_click() # 포커스 초기화 클릭
                # time.sleep(0.5)       
                gui.press('enter')         
                # find_and_click.img_click('성공저장확인.png')                
                time.sleep(0.5)                
                status_change_true(self, row_i) #### 성공 확인됨. Flag값 수정하기
                
                time.sleep(0.5)
    except Exception as e:
        logging.error('전표정보 자동저장(auto_save) Exception : ', str(e))


# #6 상태 테이블값 true로 변환
def status_change_true(self, row_i): self.payroll_status_tb.setItem(row_i, 0, QTableWidgetItem('Success'))
# def status_change_true(self, row_i): self.status_tb.setItem(row_i, 0, QTableWidgetItem('Success'))