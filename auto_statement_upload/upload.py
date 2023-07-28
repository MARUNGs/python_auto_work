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
import logging                     # 로그


##### Module import 
from module import auto_save             # 결의서/전표정보 자동업로드 및 저장기능 수행
from module import auto_save_payroll     # 급여대장 자동업로드 및 저장기능 수행
from module import check                 # 시작 전 확인기능 수행
from module import make_excel_data_table # 엑셀 데이터 생성기능 수행
from module import find_and_click        # 이미지 찾아서 클릭기능 수행
import module.xls_to_xlsx as xls_to_xlsx # 엑셀 확장자 변경




########### 전역처리 ########################################################################################
# 파일경로
main_ui = uic.loadUiType(os.path.dirname(__file__) + os.sep + 'upload_form.ui')[0]

# 공통이미지경로
img_dir_path = os.path.dirname(__file__) + os.sep + 'img' + os.sep

# 로그 설정
logging.getLogger().setLevel(logging.DEBUG) # 로그레벨 설정
log_file = 'app.log'
file_handler = logging.FileHandler(log_file)
formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
file_handler.setFormatter(formatter)
logging.getLogger().addHandler(file_handler)


# 기본 딜레이 설정
gui.PAUSE = 0.2

########### class function ##############################################################################
class window__base__setting(QMainWindow, main_ui) :
    def __init__(self) :
        super().__init__()

        # 버튼 기능 연결
        self.set_ui()
        self.find_btn.clicked.connect(self.find_fn)                                         # 전표정보 - 첨부 엑셀파일
        self.start_btn.clicked.connect(self.start_fn)                                       # 전표정보 - 시작
        self.stop_btn.clicked.connect(self.stop_fn)                                         # 전표정보 - 종료
        self.find_projectImg_btn.clicked.connect(self.find_project_img_fn)                  # 전표정보 - 첨부 사업이미지
        self.move_simple_menu_btn.clicked.connect(self.move_simple_menu_fn)                 # 전표정보 - 간편입력 메뉴로 이동

        self.move_payroll_menu_btn.clicked.connect(self.move_simple_menu_fn)               # 급여대장 등록 메뉴로 이동
        self.find_payroll_btn.clicked.connect(self.find_fn)                                 # 급여대장 - 첨부 엑셀파일
        self.find_payroll_project_img_btn.clicked.connect(self.find_project_img_fn)         # 급여대장 - 첨부 사업이미지
        self.find_payroll_year_img_btn.clicked.connect(self.find_year_img_fn)             # 급여대장 - 첨부 회계연도 이미지
        self.start_payroll_btn.clicked.connect(self.start_payroll_fn)                       # 급여대장 - 시작


    # ui 세팅
    def set_ui(self): self.setupUi(self)


    #1 파일 업로드
    def find_fn(self):
        try:
            file_path = QFileDialog.getOpenFileName(self)
            file_nm = os.path.basename(file_path[0])

            # 현재 활성화중인 탭의 정보에 따라서 file_path 정보값이 달라지도록 유도할 것.
            if(self.tabs.tabText(self.tabs.currentIndex()) in '전표정보'):   change_file_path = self.file_path
            elif(self.tabs.tabText(self.tabs.currentIndex()) in '급여대장'): change_file_path = self.file_payroll_path

            if ('.xlsx' in file_nm) or ('.xls' in file_nm):
                if 'xls' == file_nm.split('.')[1]: 
                    xls_to_xlsx.xls_to_xlsx(file_path[0]) # 파일변환 작업
                    change_file_path.setText(file_path[0] + 'x')
                else: change_file_path.setText(file_path[0])
            else: gui.alert('xlsx 또는 xls 확장자만 허용합니다.')
        except Exception as e: 
            gui.alert('파일업로드 과정에서 오류가 발생했습니다. \n관리자 확인이 필요합니다.')
            logging.error(str(e))


    #2-1 전표 간편입력 자동업로드 시작
    def start_fn(self):
        logging.info('----- Check And Start -----')

        if gui.confirm('전표정보 자동업로드 업무를 실행하시겠습니까?') == 'OK':
            starting(self)

            # 확인사항 조건이 맞으면 자동업로드 시작
            start_auto(self) if check.check(self) else self.stop_fn
        else: gui.alert('전표정보 자동업로드 실행을 취소합니다.')


    #2-2 급여대장 자동업로드 시작
    def start_payroll_fn(self):
        logging.info('----- Payroll Auto Write START -----')

        if gui.confirm('급여대장 자동업로드 업무를 실행하시겠습니까?') == 'OK':
            starting(self)

            # 확인사항 조건이 맞으면 자동업로드 시작
            start_payroll_auto(self) if check.payroll_check(self) else self.stop_fn
        else: gui.alert('급여대장 자동업로드 실행을 취소합니다.')


    #3 자동업로드 중지
    def stop_fn(self):
        gui.alert('자동화 업무를 중단합니다.')
        ending(self)
        sys.exit()


    #4 사업명 이미지 업로드
    def find_project_img_fn(self):
        try:
            file_path = QFileDialog.getOpenFileName(self)

            # 현재 활성화중인 탭의 정보에 따라서 file_path 정보값이 달라지도록 유도할 것.
            tab_txt = self.tabs.tabText(self.tabs.currentIndex())
            if(tab_txt in '전표정보'):   change_file_path = self.file_project_img_path
            elif(tab_txt in '급여대장'): change_file_path = self.file_payroll_project_img_path

            change_file_path.setText(file_path[0]) if ('.png' in file_path[0]) else gui.alert('png 확장자 이미지만 허용합니다.')
        except Exception as e:
            gui.alert('사업명 이미지 파일업로드 과정에서 오류가 발생했습니다. \n관리자 확인이 필요합니다.')
            logging.error(str(e))


    # 회계연도 이미지 업로드
    def find_year_img_fn(self):
        try:
            file_path = QFileDialog.getOpenFileName(self)
            change_file_path = self.file_payroll_year_img_path
            change_file_path.setText(file_path[0]) if ('.png' in file_path[0]) else gui.alert('png 확장자 이미지만 허용합니다.')
        except Exception as e:
            gui.alert('회계연도 이미지 파일업로드 과정에서 오류가 발생했습니다. \n관리자 확인이 필요합니다.')
            logging.error(str(e))


    # 현재 탭에 맞는 메뉴로 이동
    def move_simple_menu_fn(self):
        w4c_window = gui.getWindowsWithTitle('사회복지시설정보시스템(1W)')[0] # 프로그램 호출
        if w4c_window.isActive == False: w4c_window.activate()              # 프로그램 활성화

        # 현재 활성화중인 탭의 정보에 따라서 file_path 정보값이 달라지도록 유도할 것.
        tab_txt = self.tabs.tabText(self.tabs.currentIndex())
        if(tab_txt in '전표정보'):
            find_and_click.img_click('회계.png')
            find_and_click.img_db_click('결의및전표관리.png')
            find_and_click.img_db_click('결의서전표간편입력.png')
        elif(tab_txt in '급여대장'):
            find_and_click.img_db_click('간편입력.png')
            find_and_click.img_db_click('급여대장등록.png')
########## function ###################################################################################
#3-1# 자동화 실행(after) >> 간편입력
def start_auto(self):
    logging.info('----- Start Statement Simple Auto Write -----')

    # Action
    tab_i = self.tabs.currentIndex()
    excel_list = make_excel_data_table.make_excel_data(self, tab_i)       #1 조회한 엑셀 데이터 생성
    make_excel_data_table.make_table(self, excel_list, tab_i) #2 조회한 엑셀 데이터를 가지고 테이블 생성

    # Active
    w4c_window = gui.getWindowsWithTitle('사회복지시설정보시스템(1W)')[0] # 프로그램 호출
    if w4c_window.isActive == False: w4c_window.activate()              # 프로그램 활성화

    # Action
    auto_save.auto_save(self, excel_list)  #2 결의서/전표 자동저장 작업
    ending(self)    #3 다 끝나면 종료


#3-2# 자동화 실행 >> 급여대장
def start_payroll_auto(self):
    logging.info('----- Start Payroll  -----')
    
    # Action
    tab_i = self.tabs.currentIndex()
    excel_list = make_excel_data_table.make_excel_data(self, tab_i)       #1 조회한 엑셀 데이터 생성
    make_excel_data_table.make_table(self, excel_list, tab_i) #2 조회한 엑셀 데이터를 가지고 테이블 생성

    # Active
    w4c_window = gui.getWindowsWithTitle('사회복지시설정보시스템(1W)')[0]
    if w4c_window.isActive == False: w4c_window.activate()

    # Active 1 : 급여항목등록여부 확인하기
    find_and_click.img_click('급여대장_급여항목등록.png')

    if find_and_click.find_img_flag('급여대장_급여항목_순번.png'):
        # Action 2 : 창 닫고 급여대장 전표 자동저장 작업
        find_and_click.img_click('창닫기.png')
        auto_save_payroll.auto_save_payroll(self, excel_list)
    else: gui.alert('급여대장 자동업로드시 급여항목이 등록되어야 사용이 가능합니다.')

    ending(self)
# def start_payroll_auto End #


#6# '실행중'으로 상태변경
def starting(self) :
    self.status_text.setText('실행중')
    self.status_text.setStyleSheet('color: red')

#7# '종료'으로 상태변경
def ending(self) :
    self.status_text.setText('종료')
    self.status_text.setStyleSheet('Color: black')


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