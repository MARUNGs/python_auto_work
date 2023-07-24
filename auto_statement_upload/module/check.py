#자동화 실행 전에 w4c_cd가 DB에 등록된 정보와 일치하는지 확인하는 모듈

########## import list ##############################################################################################################################
##### Library import 
import os
import psycopg2 as pg               # PostgreSQL 연동
import win32com.client as win32     # 윈도우 앱을 활용할 수 있게 해주는 모듈
import re                           # 정규식 표현
import pyautogui as gui             # 운영체제 제어
import logging                      # 로그
import sys                          # 시스템 정보


########## function #################################################################################################################################
conn = pg.connect(host='192.168.0.11', dbname='test_hearthands', user='postgres', password='123qwe```', port=54332) # DB정보


def check(self):
    check_w4c_cd = self.w4c_cd.toPlainText().replace(' ', '')         #1 사용자가 입력한 희망e음 코드
    check_file_path = self.file_path.toPlainText()                    #2 전표정보 첨부파일 경로
    check_project_img_path = self.file_project_img_path.toPlainText() #3 사업명이미지 경로

    try: 
        if check_file_path.replace(' ', '') != '' and check_project_img_path != '' and check_w4c_cd != '':
            # w4c_cd 정규표현식 확인
            if len(check_w4c_cd) == 11 and re.match('[a-zA-z0-9]', check_w4c_cd):
                with conn:
                    cur = conn.cursor()
                    stmt = cur.mogrify('SELECT w4c_code FROM common.org_info WHERE w4c_code = %s', (check_w4c_cd, )) # PreparedStatement 생성
                    cur.execute(stmt) # PreparedStatement 실행
                    result = cur.fetchall()

                    if (len(result) > 0) and (check_w4c_cd in result[0]): return check_open_file(self)
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
        file_path = self.file_path.toPlainText()
        length = len(file_path.rsplit(os.sep))
        file_nm = file_path.rsplit(os.sep)[length-1]
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



# 급여대장 
def payroll_check(self):
    check_file_path = self.file_payroll_path.toPlainText()                            # 엑셀파일
    check_payroll_project_img_path = self.file_payroll_project_img_path.toPlainText() # 사업명이미지
    check_w4c_cd = self.payroll_w4c_cd.toPlainText().replace(' ', '')                 # 희망e음 코드

    try:
        if check_file_path.replace(' ', '') != '' and check_payroll_project_img_path != '' and check_w4c_cd != '':
            if len(check_w4c_cd) == 11 and re.match('[a-zA-z0-9]', check_w4c_cd):
                if code_check_DB(check_w4c_cd) == True:
                    return check_open_payroll_file(self)
                else:
                    gui.alert('희망e음 코드가 확인되지 않습니다. \n확인 후 다시 작업을 수행하세요.')
                    return False
            else:
                gui.alert('희망e음 인증코드가 확인되지 안습니다. \n확인 후 다시 작업을 수행하세요.')
                return False
        else:
            gui.alert('각 파일정보와 희망e음 코드 확인이 어렵습니다. \n확인 후 다시 작업을 수행하세요.')
            return False
    except Exception as e:
        gui.alert('자동화 업무 수행 전 확인단계에서 오류가 발생했습니다. \n관리자 확인이 필요합니다.')
        logging.debug('급여대장 확인 기능 오류: ', e)
        sys.exit()



# 급여대장 파일 열렸는지 확인
def check_open_payroll_file(self):
    try:
        file_path = self.file_payroll_path.toPlainText()
        length = len(file_path.rsplit(os.sep))
        file_nm = file_path.rsplit(os.sep)[length-1]
        xl = win32.Dispatch('Excel.Application')

        if len(gui.getWindowsWithTitle(file_nm.split('.')[0])) < 1: # 엑셀프로그램이 열려있지 않으면 오픈
            xl.Workbooks.Open(Filename=file_path)
            xl.Visible = True
            return True
        

        if xl.Workbooks.Count > 0: # 열려있는 파일 중 특정 Excel 이름과 일치하는 파일이 없으면 새 파일 오픈
            for excel in xl.Workbooks:
                if not excel.Name == file_nm:
                    xl.Workbooks.Open(Filename=file_path)
                    xl.Visible = True
        else:
            xl.Workbooks.Ope(Filename=file_path)
            xl.Visible=True

        return True
    except Exception as e:
        logging.debug('----- 해당 첨부파일 열림 확인 오류 -----', e)
        sys.exit()
# def check_open_payroll_file End #






# DB 정보 확인
'''
    @param check_w4c_cd      ### 희망e음 코드
    @return True / False
'''
def code_check_DB(check_w4c_cd):
    

    with conn:
        cur = conn.cursor()
        stmt = cur.mogrify('SELECT w4c_code FROM common.org_info WHERE w4c_code = %s', (check_w4c_cd, )) # PreparedStatement 생성
        cur.execute(stmt) # PreparedStatement 실행
        result = cur.fetchall()

        if len(result) > 0 and check_w4c_cd in result[0] : 
            return True
        else:
            return False