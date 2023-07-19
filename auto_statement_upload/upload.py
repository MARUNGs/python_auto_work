# 주제1 : 마음손 전표정보 업로드 자동화
# 주제2 : 마음손 급여대장 업로드 자동화


########## MEMO ##############################################################################
#1# 파이썬의 명명규칙을 따라서 작명하였음. 규칙을 준수할 것.
#  [ex 1] 변수 및 함수 이름: my_variable, calculate_sum()
#  [ex 2] 클래스 이름: MyClass, MyException




########## import list ##############################################################################
##### Library import 
import os                          # 운영체제 정보
import pyautogui as gui            # 운영체제 제어
import sys                         # 시스템 정보
from PyQt5.QtWidgets import *      # PyQt5 GUI
from PyQt5 import uic              # .ui 파일 호출
from PyQt5.QAxContainer import *
from PyQt5.QtGui import *
import openpyxl                    # 엑셀 
import logging                     # 로그


##### Module import 
from module import auto_save             # 자동업로드 및 저장기능 수행
from module import check                 # 시작 전 확인기능 수행
import module.xls_to_xlsx as xls_to_xlsx # 엑셀 확장자 변경




########### 전역처리 ########################################################################################
# 파일경로
mainUi = uic.loadUiType(os.path.dirname(__file__) + os.sep + 'upload_form.ui')[0]

# 공통이미지경로
img_dir_path = os.path.dirname(__file__) + os.sep + 'img' + os.sep

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
        self.find_projectImg_btn.clicked.connect(self.find_project_img_fn)    # 전표정보 - 첨부 사업이미지
        self.find_manageImg_btn.clicked.connect(self.find_manage_img_fn)      # 전표정보 - 첨부 계좌이미지 <<<<<<<<<<<<<<<<<< 관련 기능 곧 삭제 예정
        self.move_simple_menu_btn.clicked.connect(self.move_simple_menu_fn)   # 전표정보 - 간편입력 메뉴로 이동

        self.move_payroll_menu_btn.clicked.connect(self.move_payroll_menu_fn) # 급여대장 등록 메뉴로 이동
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
            if check.check(self): start_auto(self) # 확인사항 조건이 맞으면 자동업로드 시작
            else: self.stopFn
        else:
            gui.alert('전표정보 자동업로드 실행을 취소합니다.')
    # def start_fn End #


    #2-2 급여대장 자동업로드 시작
    def start_payroll_fn(self):
        logging.info('----- 급여대장 START -----')

    # def start_payroll_fn End


    #3 자동업로드 중지
    def stop_fn(self):
        gui.alert('자동화 업무를 중단합니다.')
        ending(self)
        sys.exit()
    # def stop_fn End #


    #4 사업명 이미지 업로드
    def find_project_img_fn(self):
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
    # def find_project_img_fn End #


    #5 계좌 이미지 업로드(필요없을 수도 있음)
    def find_manage_img_fn(self):
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
    # def fine_manage_img_fn End #


    #6 전표정보 - 간편입력 메뉴로 이동
    def move_simple_menu_fn(self):
        move_to_img('회계.png', self)
        img_db_click('결의및전표관리.png', self)
        img_db_click('결의서전표간편입력.png', self)
    # def move_simple_menu_fn End #


    #7 급여대장 등록 메뉴로 이동
    def move_payroll_menu_fn(self):
        img_db_click('간편입력.png', self)
        img_db_click('급여대장등록.png', self)
    # def move_payroll_menu_fn End #
########## function ###################################################################################








#3-1# 자동화 실행(after) >> 간편입력
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


#3-2# 자동화 실행 >> 급여대장
def start_payroll_auto(self):
    print('작업예정')
# def start_payroll_auto End #


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