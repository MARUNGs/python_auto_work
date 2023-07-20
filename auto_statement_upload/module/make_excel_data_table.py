# 엑셀파일을 기준으로 데이터를 생성하거나, 테이블을 생성하는 작업

########## import list ##############################################################################################################################
##### Library import 
from PyQt5.QtWidgets import *      # PyQt5 GUI
import psycopg2 as pg               # PostgreSQL 연동
import win32com.client as win32     # 윈도우 앱을 활용할 수 있게 해주는 모듈
import re                           # 정규식 표현
import pyautogui as gui             # 운영체제 제어
import logging                      # 로그
import sys                          # 시스템 정보
import openpyxl                     # 엑셀 


########## function #################################################################################################################################
#4# 엑셀데이터 생성
''' 
    @param self
    @return excelList 
'''
def make_excel_data(self, title):
    try:
        if title=='statement':
            wb = openpyxl.load_workbook(self.file_path.toPlainText())
        elif title=='payroll': #급여대장
            wb = openpyxl.load_workbook(self.file_payroll_path.toPlainText())
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
def make_table(self, title_list, excel_list, title):
    try:
        if title=='statement':
            wb = openpyxl.load_workbook(self.file_path.toPlainText())
        elif title=='payroll':
            wb = openpyxl.load_workbook(self.file_payroll_path.toPlainText())
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