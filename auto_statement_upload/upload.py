# 주제1 : 마음손 전표정보 업로드 자동화
# 주제2 : 마음손 급여대장 업로드 자동화


########## MEMO ##############################################################################
#1# 파이썬의 명명규칙을 따라서 작명하였음. 규칙을 준수할 것.
#  [ex 1] 변수 및 함수 이름: my_variable, calculate_sum()
#  [ex 2] 클래스 이름: MyClass, MyException




########## import list ##############################################################################
##### Library import 
import os        # 운영체제 정보
import pyautogui as gui # 운영체제 제어
import sys              # 시스템 정보
from PyQt5.QtWidgets import *  # PyQt5 GUI
from PyQt5 import uic   # .ui 파일 호출
from PyQt5.QAxContainer import *
from PyQt5.QtGui import *
import psycopg2 as pg # PostgreSQL 연동
import re # 정규식 표현
import openpyxl # 엑셀 
import win32com.client as win32 # 윈도우 앱을 활용할 수 있게 해주는 모듈
import logging # 로그


##### Module import 
from module import auto_save             # 자동업로드 및 저장기능 수행
import module.xls_to_xlsx as xls_to_xlsx # 엑셀 확장자 변경




########### 전역처리 ########################################################################################
# 파일경로
mainUi = uic.loadUiType(os.path.dirname(__file__) + os.sep + 'upload_form.ui')[0]

# 로그 설정
logFormat = logging.Formatter('%(asctime)s [%(levelname)s] %(message)s')
logger = logging.getLogger()
logger.setLevel(logging.DEBUG)

# 기본 딜레이 설정
gui.PAUSE = 0.2

########### class function ##############################################################################
class window__base__setting(QMainWindow, mainUi) :
    def __init__(self) :
        super().__init__()

        # 버튼 기능 연결
        self.set_ui()
        self.find_btn.clicked.connect(self.find_fn)                        # 전표정보 - 첨부 엑셀파일
        self.start_btn.clicked.connect(self.start_fn)                      # 전표정보 - 시작
        self.stop_btn.clicked.connect(self.stop_fn)                        # 전표정보 - 종료
        self.find_projectImg_btn.clicked.connect(self.find_project_img)     # 전표정보 - 첨부 사업이미지
        self.find_manageImg_btn.clicked.connect(self.find_manage_img)       # 전표정보 - 첨부 계좌이미지

        # self.file_payroll_btn.clicked.connect(self.find_fn)                # 급여대장 - 첨부 엑셀파일
        # self.start_payroll_btn.clicked.connect(self.start_payroll_fn)       # 급여대장 - 시작

    # def __init__ End #


    # ui 세팅
    def set_ui(self): self.setupUi(self)
    # def set_ui End #


    #1 파일 업로드
    def find_fn(self):
        try:
            filePath = QFileDialog.getOpenFileName(self)
            fileNm = os.path.basename(filePath[0])

            if ('.xlsx' in fileNm) or ('.xls' in fileNm):
                if 'xls' == fileNm.split('.')[1]: 
                    xls_to_xlsx.xls_to_xlsx(self)                   # 파일변환 작업
                    self.file_nm.setText(fileNm + 'x')
                    self.file_path.setText(filePath[0] + 'x')
                else:
                    self.file_nm.setText(fileNm)
                    self.file_path.setText(filePath[0])
            else: 
                gui.alert('xlsx 또는 xls 확장자만 허용합니다.')
        except Exception as e: 
            gui.alert('파일업로드 과정에서 오류가 발생했습니다. \n관리자 확인이 필요합니다.')
            logging.debug(e)
            sys.exit()
    # def find_fn End #


    #2-1 전표 간편입력 자동업로드 시작
    def start_fn(self):
        logging.info('----- 전표 간편입력 START -----')

        if gui.confirm('전표정보 자동업로드 업무를 실행하시겠습니까?'):
            starting(self)
            if check(self): start_auto(self) # 확인사항 조건이 맞으면 자동업로드 시작
            else: self.stopFn
        else:
            gui.alert('전표정보 자동업로드 실행을 취소합니다.')
    # def start_fn End #


    #2-2 급여대장 자동업로드 시작
    def start_payroll_fn(self):
        logging.info('----- 급여대장 START -----')

    # def startPayrollFn End


    #3 자동업로드 중지
    def stop_fn(self):
        gui.alert('자동화 업무를 중단합니다.')
        ending(self)
        sys.exit()
    # def stopFn End #


    #4 사업명 이미지 업로드
    def find_project_img(self):
        try:
            filePath = QFileDialog.getOpenFileName(self)
            fileNm = os.path.basename(filePath[0])

            if '.png' in fileNm:
                self.file_project_img_nm.setText(fileNm)
                self.file_project_img_path.setText(filePath)
            else:
                gui.alert('png 확장자 이미지만 허용합니다.')
        except Exception as e:
            gui.alert('사업명 이미지 파일업로드 과정에서 오류가 발생했습니다. \n관리자 확인이 필요합니다.')
            logging.debug(e)
            sys.exit()
    # def findProjectImg End #


    #5 계좌 이미지 업로드(필요없을 수도 있음)
    def find_manage_img(self):
        try:
            filePath = QFileDialog.getOpenFileName(self)
            fileNm = os.path.basename(filePath[0])

            if '.png' in fileNm:
                self.file_manage_img_nm.setText(fileNm)
                self.file_manage_img_path.setText(filePath)
            else:
                gui.alert('png 확장자 이미지만 허용합니다.')
        except Exception as e:
            gui.alert('계좌명 이미지 파일업로드 과정에서 오류가 발생했습니다. \n관리자 확인이 필요합니다.')
            logging.debug(e)
            sys.exit()
########## function ###################################################################################
#1# 자동화 실행 전에 w4c_cd가 DB에 등록된 정보와 일치하는지 확인.
def check(self):
    check_w4c_cd = self.w4c_cd.toPlainText().replace(' ', '')         #1 사용자가 입력한 희망e음 코드
    check_file_path = self.file_path.toPlainText()                    #2 전표정보 첨부파일 경로
    check_project_img_path = self.file_project_img_path.toPlainText() #3 사업명이미지 경로
    check_manage_img_path = self.file_manage_img_path.toPlainText()   #4 계좌명이미지 경로

    try: 
        if check_file_path.replace(' ', '') != '' and check_project_img_path != '' and check_manage_img_path != '' and check_w4c_cd != '':
            # w4c_cd 정규표현식 확인
            if len(check_w4c_cd) == 11 and re.match('[a-zA-z0-9]', check_w4c_cd):
                conn = pg.connect(host='192.168.0.11', dbname='test_hearthands', user='postgres', password='123qwe```', port=54332) # DB정보

                with conn:
                    cur = conn.cursor()
                    stmt = cur.mogrify('SELECT w4c_code FROM common.org_info WHERE w4c_code = %s', (check_w4c_cd, )) # PreparedStatement 생성
                    cur.execute(stmt) # PreparedStatement 실행
                    result = cur.fetchall()

                    if len(result) > 0 and check_w4c_cd in result[0] : 
                        return check_open_file(self)
                    else : 
                        gui.alert('희망e음 인증코드가 확인되지 않습니다. \n확인 후 다시 작업을 수행하세요.')
                        return False
            else:
                gui.alert('첨부파일 및 희망e음코드가 올바르지 않습니다. \n확인 후 다시 작업을 수행하세요.')
                return False
    except Exception as e:
        gui.alert('자동화 업무 수행 전 확인단계에서 오류가 발생했습니다. \n업로드한 자료 및 희망e음 인증코드를 확인하세요.')
        logging.debug('check Exception : ', e)
        sys.exit()
# def check End #



#2# 파일이 열려있는지 확인.
def check_open_file(self) :
    try :
        file_nm = self.file_nm.toPlainText()
        file_path = self.file_path.toPlainText()
        xl = win32.Dispatch('Excel.Application')

        if len(gui.getWindowsWithTitle(file_nm.split('.')[0])) < 1: # 아예 엑셀프로그램이 열려있지 않으면 오픈
            xl.Workbooks.Open(Filename = file_path)
            xl.Visible = True
            return True
        
            
        if xl.Workbooks.Count > 0 :     # 열려있는 파일 중 특정 Excel 이름과 일치하는 파일이 없으면 새 파일 오픈
            for excel in xl.Workbooks :
                if not excel.Name == file_nm :
                    xl.Workbooks.Open(Filename = file_path)
                    xl.Visible = True
        else:                           # 안 열려져 있으면 오픈
            xl.Workbooks.Open(Filename = file_path)
            xl.Visible = True

        return True
    except Exception as e:
        logging.debug('---- 해당 첨부파일 열림 확인 오류 ----', e)
        sys.exit()
# def check_open_file End #



#3# 자동화 실행(after)
def start_auto(self):
    logging.info('----- 전표정보 자동업로드 업무 실행 -----')

    # Active
    excel_window = gui.getWindowsWithTitle('마음손거래내역')[0] # 파일명 호출
    if excel_window.isActive == False: excel_window.activate() # 파일 활성화
    
    # Action
    excel_list = make_excel_data(self)       #1 조회한 엑셀 데이터 생성
    title_list = excel_list[0]
    make_table(self, title_list, excel_list) #2 조회한 엑셀 데이터를 가지고 테이블 생성

    # Active
    w4c_window = gui.getWindowsWithTitle('사회복지시설정보시스템(1W)')[0] # 프로그램 호출
    if w4c_window.isActive == False: w4c_window.activate()              # 프로그램 활성화

    # Action
    auto_save.auto_save(self, excel_list)  #2 결의서/전표 자동저장 작업

    ending(self)    #3 다 끝나면 종료
# def start_auto(self) End #



#4# 엑셀데이터 생성
''' 
    @param self
    @return excelList 
'''
def make_excel_data(self):
    try:
        wb = openpyxl.load_workbook(self.file_path.toPlainText())
        sheet = wb[wb._sheets[0].title]
        max_col_cnt = sheet.max_column
        excel_list = [] # 객체를 담을 리스트


        for rows in sheet.iter_rows() :
            '''
                한 행의 데이터를 담을 딕셔너리 자료형 - 엑셀 항목 기준
                01. incomeExpenseCode : 수입지출구분
                    (반납구분은 무시해도 될 듯.)
                02. cashierDate : 거래일자
                03. accountSubject : 계정과목
                04. summary : 적요
                05. incomeAmt : 수입금액
                06. expenseAmt : 지출금액
                07. capitalSource : 자금원천
                08. opponentSubject : 상대계정
                09. resolutionNo : 결의번호
                10. project : 사업구분(사업명)
                11. manage : 계좌명
            '''
            dataList = []

            for i in range(0, max_col_cnt):
                inputData = None
                cell = rows[i]

                if str(cell.value) == 'None': inputData = ''
                else: inputData = str(cell.value).replace(' 00:00:00', '') # 거래일자 시분초 제거

                dataList.insert(i, inputData) # list 형태로 삽입해야 함..
            # for in range End #

            excel_list.insert(cell.row - 1, dataList) # 0 index부터 삽입
        # for in End #

        return excel_list
    except Exception as e:
        logging.debug('엑셀 데이터 생성 실패 : ', e)
        sys.exit()
# def make_excel_data End #


#5# 테이블 생성
def make_table(self, title_list, excel_list):
    try:
        wb = openpyxl.load_workbook(self.file_path.toPlainText())
        sheet = wb[wb._sheets[0].title]
        max_col_cnt = sheet.max_column
        max_row_cnt = sheet.max_row - 1  # 타이틀을 제외한 데이터 row수
        excel_tb = self.excel_tb         # 엑셀 테이블
        status_tb = self.status_tb       # 상태 테이블

        # 테이블 세팅
        excel_tb.setColumnCount(max_col_cnt)
        excel_tb.setRowCount(max_row_cnt)
        excel_tb.setHorizontalHeaderLabels(title_list) # list 형태로 넣기
        del excel_list[0]
        status_tb.setRowCount(max_row_cnt)


        # 테이블 내 엑셀데이터 기본설정
        for i in range(0, max_row_cnt) :
            data = excel_list[i]

            for j in range(0, len(data)): 
                excel_tb.setItem(i, j, QTableWidgetItem(data[j]))
            # for in range End #
        # for in range End #


        # 상태 테이블 기본설정
        for i in range(0, max_row_cnt) :
            status_tb.setItem(i, 0, QTableWidgetItem('Fail'))
        # for in range End #
    except Exception as e:
        logging.debug('엑셀 테이블 생성 실패 : ', e)
        sys.exit()
# def make_table(self) End


#6# '실행중'으로 상태변경
def starting(self) :
    self.status_text.setText('실행중')
    self.status_text.setStyleSheet('color: red')
# def starting(self) End #

#7# '종료'으로 상태변경
def ending(self) :
    self.status_text.setText('종료')
    self.status_text.setStyleSheet('Color: black')
# def ending(self) End #


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