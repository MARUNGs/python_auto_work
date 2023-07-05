# 주제: 마음손 전표정보 업로드 자동화

# [ 작업순서 ]
# 1. 업로드할 파일명은 미리 복사하여 입력창에 작성해두어야 한다.
# 2. 입력한 파일명을 기준으로 파일을 검색한다.
# 3. 검색된 파일명이 맞으면 enter 키 입력을 수행한다.
# 4. 파일명에 존재하는 확장자를 추출하여 xlsx인지 xls인지 확인한다.
# 5. xls인 경우, 호환성 여부 안내창을 인식한다.



########## import list ##############################################################################
##### Library import 
import os               # 운영체제 정보
# from tkinter import *   # 파이썬 UI 구현
import pyautogui as gui # 마우스 & 키보드 제어
import sys              # 시스템 정보
from PyQt5.QtWidgets import * # PyQt5 GUI
from PyQt5 import uic   # .ui 파일 호출
from PyQt5.QAxContainer import *
from PyQt5.QtGui import *
import psycopg2 as pg # PostgreSQL 연동
import re # 정규식 표현
import time
from openpyxl import Workbook # 엑셀 
import win32com.client as win32 # 윈도우 앱을 활용할 수 있게 해주는 모듈

##### function import
from func_dir import change_status_fn # 상태값 변경 함수 모음



########## 참고한 블로그 ###############################################################################
# PyQt5 사용법 : https://coding-kindergarten.tistory.com/60
# 파이썬과 PostgreSQL 연결 : https://edudeveloper.tistory.com/131
# Tkinter 위젯 배치 :  https://camplee.tistory.com/32
# Tkinter 위젯 x,y 배치 : https://cosmosproject.tistory.com/610
# Tkinter 여러가지 설정(읽어보기) : https://m.blog.naver.com/PostView.naver?isHttpsRedirect=true&blogId=yheekeun&logNo=220701094001
# 파일이 열려있는가 확인 : https://cyworld.tistory.com/3053
# f문자열 포맷팅 사용법 : https://www.daleseo.com/python-f-strings/
# 파일이 python에서 열려있는지 확인하는 방법 : https://cyworld.tistory.com/3053


########### 전역처리 ########################################################################################
main_ui = uic.loadUiType(os.path.dirname(__file__) + os.sep + 'test.ui')[0] # 파일경로


########### class function ##############################################################################
'''
    __init__(self) : Setting Window Base Form
'''
class window__base__setting(QMainWindow, main_ui) :
    def __init__(self) :
        super().__init__()

        # 버튼 기능 연결
        self.set_ui()
        self.find_btn.clicked.connect(self.find_fn)
        self.start_btn.clicked.connect(self.start_fn)
        self.stop_btn.clicked.connect(self.stop_fn)

    # ui 세팅
    def set_ui(self): self.setupUi(self)

    # 파일 업로드
    def find_fn(self):
        try:
            file_path = QFileDialog.getOpenFileName(self)
            file_nm = os.path.basename(file_path[0])

            # 파일명이 .xlsx 또는 .xls 문자열이 포함하지 않으면 이벤트를 종료한다.
            if ('.xlsx' in file_nm) or ('.xls' in file_nm):
                # 파일명/경로 세팅
                self.file_nm.setText(file_nm)
                self.file_path.setText(file_path[0])
            else: 
                gui.alert('xlsx 또는 xls 확장자만 허용합니다.')
        except:
            gui.alert('파일업로드 과정에서 오류가 발생했습니다. \n관리자 확인이 필요합니다.')

    # 자동업로드 시작
    def start_fn(self):
        if gui.confirm('자동화 업무를 실행하시겠습니까?'):
            change_status_fn.starting(self)

            # if self.check and self.check_open_file :
            if check(self) and self.check_open_file :
                self.start_auto
            else :
                return False
        else:
            gui.alert('자동화 업무 실행을 취소합니다.')
            print('fail start')


    # 자동화 실행 전 확인(before)
    # def check(self):
    #     try: 
    #         # 위에서 첨부파일 체크를 수행했으니 아래에서는 w4c_cd만 수행하면 됨
    #         check_w4c_cd = self.w4c_cd.toPlainText().replace(' ', '')

    #         # 첨부파일 text, w4c_cd text 빈 값 확인
    #         if self.file_nm.toPlainText().replace(' ', '') != '' and self.file_path.toPlainText().replace(' ', '') != '' and check_w4c_cd != '':
    #             # w4c_cd 정규표현식 확인
    #             if len(check_w4c_cd) == 11 and re.match('[a-zA-z0-9]', check_w4c_cd):
    #                 # DB에 저장된 코드값과 동일한지 확인
    #                 conn = pg.connect(host='192.168.0.11', dbname='test_hearthands', user='postgres', password='123qwe```', port=54332)
    #                 with conn:
    #                     cur = conn.cursor()
    #                     # PreparedStatement 생성
    #                     stmt = cur.mogrify('SELECT w4c_code FROM common.org_info WHERE w4c_code = %s', (check_w4c_cd, ))
    #                     # PreparedStatement 실행
    #                     cur.execute(stmt)
    #                     result = cur.fetchone()
    #                     conn.commit()
    #                     # self.start_auto() ################## 시작
    #                     return True
    #             else:
    #                 gui.alert('첨부파일 및 자동화 정보가 올바르지 않습니다. \n확인 후 다시 작업을 수행하세요.')
    #                 return False
    #     except Exception as e:
    #         print(e)
    #         gui.alert('자동화 업무 수행 전 확인단계에서 오류가 발생했습니다. \n업로드한 자료 및 희망e음 인증코드를 확인하세요.')
    

    # 파일이 열려있는지 확인.
    def check_open_file(self) :
        try :
            file_nm = self.file_nm.toPlainText()
            file_path = self.file_path.toPlainText()
            xl = win32.gencache.EnsureDispatch('Excel.Application')
            
            # 열려있는 파일 중 특정 Excel 이름과 일치하는 파일이 없으면 새 파일 오픈
            if xl.Workbooks.Count > 0 :
                for excel in xl.Workbooks :
                    if not excel.Name == file_nm :
                        xl.Workbooks.Open(Filename = file_path)
                        xl.Visible = True #화면에 표시
            else :
                # 안 열려져 있으면 오픈
                xl.Workbooks.Open(Filename = file_path)
                xl.Visible = True

            time.sleep(0.2)
            gui.hotkey("ctrl", "F10")

            return True
        except Exception as e:
            print(e)
        


    # 자동화 실행(after)
    def start_auto(self):
        '''
            자동화 업무 시작 !!!
            # 전제조건 1. 엑셀파일과 W4C 프로그램이 실행되어 있어야 한다.
        '''

        try :
            print('start_auto')
            #1
            
            #2 
        except Exception as e:
            print(e)

        
        

    # 자동업로드 중지
    def stop_fn(self):
        gui.alert('자동화 업무를 중단합니다.')
        change_status_fn.ending(self)




########## function ###################################################################################
# 자동업로드 시작 전 체크
def check(self) :
    try: 
        # 위에서 첨부파일 체크를 수행했으니 아래에서는 w4c_cd만 수행하면 됨
        check_w4c_cd = self.w4c_cd.toPlainText().replace(' ', '')

        # 첨부파일 text, w4c_cd text 빈 값 확인
        if self.file_nm.toPlainText().replace(' ', '') != '' and self.file_path.toPlainText().replace(' ', '') != '' and check_w4c_cd != '':
            # w4c_cd 정규표현식 확인
            if len(check_w4c_cd) == 11 and re.match('[a-zA-z0-9]', check_w4c_cd):
                # DB에 저장된 코드값과 동일한지 확인
                conn = pg.connect(host='192.168.0.11', dbname='test_hearthands', user='postgres', password='123qwe```', port=54332)
                with conn:
                    cur = conn.cursor()
                    # PreparedStatement 생성
                    stmt = cur.mogrify('SELECT w4c_code FROM common.org_info WHERE w4c_code = %s', (check_w4c_cd, ))
                    # PreparedStatement 실행
                    cur.execute(stmt)
                    result = cur.fetchone()
                    conn.commit()
                    # self.start_auto() ################## 시작
                    return True
            else:
                gui.alert('첨부파일 및 자동화 정보가 올바르지 않습니다. \n확인 후 다시 작업을 수행하세요.')
                return False
    except Exception as e:
        print(e)
        gui.alert('자동화 업무 수행 전 확인단계에서 오류가 발생했습니다. \n업로드한 자료 및 희망e음 인증코드를 확인하세요.')

########## Start Program(PyQt5 Designer) ###################################################################################
'''
    프로그램 시작
'''
if 'upload.py' in __file__ :
    app = QApplication(sys.argv)
    window = window__base__setting()
    window.show()
    app.exec_()
else :
    gui.alert('프로그램 시작 과정에서 문제 발생')