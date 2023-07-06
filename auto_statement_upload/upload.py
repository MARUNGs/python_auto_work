# 주제: 마음손 전표정보 업로드 자동화

# [ 작업순서 ]
# 1. 업로드할 파일명은 미리 복사하여 입력창에 작성해두어야 한다.
# 2. 입력한 파일명을 기준으로 파일을 검색한다.
# 3. 검색된 파일명이 맞으면 enter 키 입력을 수행한다.
# 4. 파일명에 존재하는 확장자를 추출하여 xlsx인지 xls인지 확인한다.
# 5. xls인 경우, 호환성 여부 안내창을 인식한다.



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
import time
import openpyxl # 엑셀 
import win32com.client as win32 # 윈도우 앱을 활용할 수 있게 해주는 모듈



########## 참고한 블로그 ###############################################################################
# PyQt5 사용법 : https://coding-kindergarten.tistory.com/60
# 파이썬과 PostgreSQL 연결 : https://edudeveloper.tistory.com/131
# Tkinter 위젯 배치 :  https://camplee.tistory.com/32
# Tkinter 위젯 x,y 배치 : https://cosmosproject.tistory.com/610
# Tkinter 여러가지 설정(읽어보기) : https://m.blog.naver.com/PostView.naver?isHttpsRedirect=true&blogId=yheekeun&logNo=220701094001
# f문자열 포맷팅 사용법 : https://www.daleseo.com/python-f-strings/
# 파일이 python에서 열려있는지 확인하는 방법 : https://cyworld.tistory.com/3053
# 이미지 좌표 확인 및 가운데 클릭하는 방법 : https://forward.tistory.com/entry/pyautogui-%EC%9D%B4%EB%AF%B8%EC%A7%80-%EC%A2%8C%ED%91%9C-%EC%95%8C%EC%95%84%EB%82%B4-%EB%A7%88%EC%9A%B0%EC%8A%A4%EB%A1%9C-%ED%81%B4%EB%A6%AD

########### 전역처리 ########################################################################################
mainUi = uic.loadUiType(os.path.dirname(__file__) + os.sep + 'upload_form.ui')[0] # 파일경로





########### class function ##############################################################################
class window__base__setting(QMainWindow, mainUi) :
    def __init__(self) :
        super().__init__()

        # 버튼 기능 연결
        self.set_ui()
        self.find_btn.clicked.connect(self.findFn)
        self.start_btn.clicked.connect(self.startFn)
        self.stop_btn.clicked.connect(self.stopFn)
    # def __init__ End #


    # ui 세팅
    def set_ui(self): self.setupUi(self)
    # def set_ui End #


    #1 파일 업로드
    def findFn(self):
        try:
            filePath = QFileDialog.getOpenFileName(self)
            fileNm = os.path.basename(filePath[0])

            # 파일명이 .xlsx 또는 .xls 문자열이 포함하지 않으면 이벤트를 종료한다.
            if ('.xlsx' in fileNm) or ('.xls' in fileNm):
                # 파일명/경로 세팅
                self.file_nm.setText(fileNm)
                self.file_path.setText(filePath[0])
            else: 
                gui.alert('xlsx 또는 xls 확장자만 허용합니다.')
        except Exception as e:
            gui.alert('파일업로드 과정에서 오류가 발생했습니다. \n관리자 확인이 필요합니다.')
            return False
    # def findFn End #


    #2 자동업로드 시작
    def startFn(self):
        if gui.confirm('자동화 업무를 실행하시겠습니까?'):
            starting(self)

            # 확인사항 조건이 맞으면 자동업로드 시작
            if check(self) and checkOpenFile(self): startAuto(self)
            else: self.stopFn
        else:
            gui.alert('자동화 업무 실행을 취소합니다.')
            print('fail start')
    # def startFn End #


    #3 자동업로드 중지
    def stopFn(self):
        gui.alert('자동화 업무를 중단합니다.')
        ending(self)
        return False
    # def stopFn End #

    
        

        
        

    




########## function ###################################################################################
# 자동화 실행 전에 w4c_cd가 DB에 등록된 정보와 일치하는지 확인.
def check(self):
    checkW4cCd = self.w4c_cd.toPlainText().replace(' ', '') # 사용자가 입력한 w4c_cd
    checkFileNm = self.file_nm.toPlainText() # 사용자가 호출한 첨부파일명
    checkFilePath = self.file_path.toPlainText() # 사용자가 호출한 첨부파일 경로

    try: 
        if checkFileNm.replace(' ', '') != '' and checkFilePath.replace(' ', '') != '' and checkW4cCd != '':
            # w4c_cd 정규표현식 확인
            if len(checkW4cCd) == 11 and re.match('[a-zA-z0-9]', checkW4cCd):
                conn = pg.connect(host='192.168.0.11', dbname='test_hearthands', user='postgres', password='123qwe```', port=54332) # DB정보
                with conn:
                    cur = conn.cursor()
                    stmt = cur.mogrify('SELECT w4c_code FROM common.org_info WHERE w4c_code = %s', (checkW4cCd, )) # PreparedStatement 생성
                    cur.execute(stmt) # PreparedStatement 실행
                    result = cur.fetchall()

                    if len(result) > 0 and checkW4cCd in result[0] : 
                        return True
                    else : 
                        gui.alert('희망e음 인증코드가 확인되지 않습니다. \n확인 후 다시 작업을 수행하세요.')
                        return False
            else:
                gui.alert('첨부파일 및 자동화 정보가 올바르지 않습니다. \n확인 후 다시 작업을 수행하세요.')
                return False
    except Exception as e:
        print(e)
        gui.alert('자동화 업무 수행 전 확인단계에서 오류가 발생했습니다. \n업로드한 자료 및 희망e음 인증코드를 확인하세요.')
# def check End #



# 파일이 열려있는지 확인.
def checkOpenFile(self) :
    try :
        fileNm = self.file_nm.toPlainText()
        filePath = self.file_path.toPlainText()
        xl = win32.gencache.EnsureDispatch('Excel.Application')
        
        if xl.Workbooks.Count > 0 :     # 열려있는 파일 중 특정 Excel 이름과 일치하는 파일이 없으면 새 파일 오픈
            for excel in xl.Workbooks :
                if not excel.Name == fileNm :
                    xl.Workbooks.Open(Filename = filePath)
                    xl.Visible = True
        else:                           # 안 열려져 있으면 오픈
            xl.Workbooks.Open(Filename = filePath)
            xl.Visible = True

        return True
    except Exception as e:
        print(e)
# def checkOpenFile End #



# 자동화 실행(after)
def startAuto(self):
    excelWindow = gui.getWindowsWithTitle('마음손거래내역')[0] # 파일명 호출
    if excelWindow.isActive == False: excelWindow.activate() # 파일 활성화

    # Action
    makeTable(self) # 조회한 엑셀 데이터를 가지고 테이블 생성
    autoSave(self) # 결의서/전표 자동저장 작업
    ending(self) # 다 끝나면 종료
    
# def startAuto(self) End #



# 테이블 생성
def makeTable(self):
    try:
        wb = openpyxl.load_workbook(self.file_path.toPlainText())
        sheet = wb[wb._sheets[0].title]
        maxColumnCnt = sheet.max_column
        maxRowCnt = sheet.max_row - 1 # 타이틀을 제외한 데이터 row수
        excelList = [] # 객체를 담을 리스트


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

            for i in range(0, maxColumnCnt):
                inputData = None
                cell = rows[i]

                if str(cell.value) == 'None': inputData = ''
                else: inputData = str(cell.value)

                dataList.insert(i, inputData) # list 형태로 삽입해야 함..
            # for in range End #

            excelList.insert(cell.row - 1, dataList) # 0 index부터 삽입
        # for in End #



        excelTb = self.excel_tb # 엑셀 테이블
        statusTb = self.status_tb # 상태 테이블
        

        # 테이블 세팅
        excelTb.setColumnCount(maxColumnCnt)
        excelTb.setRowCount(maxRowCnt)
        excelTb.setHorizontalHeaderLabels(excelList[0]) # list 형태로 넣기
        del excelList[0] # 타이틀만 있는 리스트 삭제

        statusTb.setRowCount(maxRowCnt)


        # 테이블 내 엑셀데이터 기본설정
        for i in range(0, maxRowCnt) :
            data = excelList[i]

            for j in range(0, len(data)): 
                excelTb.setItem(i, j, QTableWidgetItem(data[j]))
            # for in range End #
        # for in range End #


        # 상태 테이블 기본설정
        for i in range(0, maxRowCnt) :
            statusTb.setItem(i, 0, QTableWidgetItem('False'))
        # for in range End #
    except Exception as e:
        print(e)
# def makeTable(self) End


# 결의서/전표 자동저장 작업
def autoSave(self):
    try:
        print('auto save')
    except Exception as e:
        print(e) 
# def autoSave End #




# '실행중'으로 상태변경
def starting(self) :
    self.status_text.setText('실행중')
    self.status_text.setStyleSheet('color: red')
# def starting(self) End #

# '종료'으로 상태변경
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