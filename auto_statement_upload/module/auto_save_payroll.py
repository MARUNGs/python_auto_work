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


'''
    @param self       # PyQT5
    @param excel_list # makeExcelData()를 통해 갖고있는 데이터
'''
def auto_save_payroll(self, excel_list):
    try:
        max_col_cnt = len(excel_list[0]) if len(excel_list) > 0 else 0

        for row_i in range(0, len(excel_list)) :
            rows = excel_list[row_i]

            gui.hotkey('alt', 'f3') # 조회
            time.sleep(0.5)

            # 사업명 리프레시
            find_and_click.img_right_150_click('급여대장_사업명.png')
            find_and_click.img_click('급여대장_사업명_선택하세요.png') if find_and_click.find_img_flag('급여대장_사업명_선택하세요.png') else None

            gui.press('enter') # 팝업창 뜸.. 제거하기 위한 엔터
            time.sleep(0.5)

            find_and_click.img_click('행추가.png')
            
            # 팝업창 오픈
            find_and_click.img_click('선택.png')
            find_and_click.img_click('전체선택체크박스.png')

            # 사업명 선택
            find_and_click.img_right_150_click('급여대장_사업명.png')
            find_and_click.customization_payroll_project_img_click(self) # 사업명 선택
            find_and_click.img_click('지출결의서_등록.png') # 지출결의서 등록

            # 회계연도의 왼쪽 클릭
            find_and_click.customization_payroll_year_img_click(self)
            find_and_click.img_click('회계연도_선택.png') # 회계연도 선택

            time.sleep(1.0)

            for i in range(3, max_col_cnt):
                '''
                    급여대장은 지출결의서만 관리하므로 기본으로 세팅되는 구분, 사업구분은 작업하지 않고
                    거래일자부터 작성하면 된다.

                    *** 미리 세팅되는 항목: 사업, 자금원천, 지출, 계좌
                '''
                data = rows[i]

                

                if i==3:
                    find_and_click.img_right_click('급여대장_결의일자.png')
                    # 결의일자 삽입
                    gui.hotkey('ctrl', 'a')
                    gui.press('backspace')
                    pyperclip.copy(data)
                    gui.hotkey('ctrl', 'v')

                    find_and_click.img_click('행추가.png')
                    continue
                elif i==4:
                    find_and_click.img_bottom_right_in_click('급여대장_계정과목_타이틀.png')
                    find_and_click.img_click('급여대장_코드명.png')
                    gui.hotkey('ctrl', 'a')
                    gui.press('backspace')
                    pyperclip.copy(data)
                    gui.hotkey('ctrl', 'v')
                    time.sleep(0.5)

                    #선택

                    gui.press('enter')
                    time.sleep(1.0)
                    gui.press('enter')
                    time.sleep(1.0)

                    # 다음 항목 활성화
                    for idx in range(0,3): gui.press('tab')
 
                    continue
                elif i==7:
                    # tab으로 찾은 금액 항목에 입력.
                    gui.hotkey('ctrl', 'a')
                    gui.press('backspace')
                    pyperclip.copy(data)
                    gui.hotkey('ctrl', 'v')
                    time.sleep(0.5)

                    # 다음 항목(적요) 활성화
                    # 만약 적요를 작성할거라면 별도로 처리하던지 여기서 이어서 작성하던지 해야할 듯

                    # 다음 항목 활성화
                    for idx in range(0,2): gui.press('tab')

                    continue
                elif i==9:
                    # tab으로 찾은 상대계정 항목에 입력
                    gui.hotkey('ctrl', 'a')
                    gui.press('backspace')
                    pyperclip.copy(data)
                    gui.hotkey('ctrl', 'v')
                    time.sleep(0.5)

                    # 팝업창을 오픈하기 위한 tab
                    gui.press('tab')
                    time.sleep(0.5)

                    # 인건비 반영을 기준으로 그냥 엔터 침
                    if data == '장기요양급여수입':
                        time.sleep(1.0)
                        gui.press('enter')
                        time.sleep(0.5)
                    
                    continue
            
            if i == max_col_cnt-1:
                find_and_click.img_click('급여대장_저장.png')
                gui.press('enter') # 저장여부 '확인'
                time.sleep(0.5)
                gui.press('enter') # 성공저장 확인
                time.sleep(0.5)
                find_and_click.img_click('급여대장_닫기.png')
                time.sleep(0.5)
                status_change_true(self, row_i) #### 성공 확인됨. Flag값 수정하기
                
                time.sleep(0.5)

        status_change_true(self, len(excel_list)) #### 사실 여기까지 왔으면 성공취급함. 마지막줄 Flag값 수정하기
    except Exception as e:
        logging.error('급여대장 자동저장(payroll_auto_save) Exception : ', str(e))
        sys.exit()


# 상태 테이블값 true로 변환
def status_change_true(self, row_i): self.status_tb.setItem(row_i, 0, QTableWidgetItem('Success'))
