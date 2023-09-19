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
import base64
from sshtunnel import SSHTunnelForwarder # SSH DB 연결
import paramiko
import io
import tkinter as tk



applogger = logging.getLogger("app")
global_logger = logging.getLogger()

########## function #################################################################################################################################
'''
    @param check_w4c_cd
    @return result
'''
def DB_setting_and_select_result(check_w4c_cd, self):
    applogger.debug('all check before')

    try:
        # pem 파일 경로
        pem_path = __file__.replace(__file__.rsplit(os.sep)[len(__file__.rsplit(os.sep))-1], '') + 'hearthands-aws.pem'

        with open(pem_path, 'rb') as f: blob = base64.b64encode(f.read())
        pem_decode = blob.decode('utf-8')

        applogger.debug("-------------------------------------------------------------------------------------")
        applogger.debug("1. " + pem_decode)
        applogger.debug("-------------------------------------------------------------------------------------")

        SSH_KEY_BLOB_DECODED = base64.b64decode(pem_decode)
        applogger.debug("-------------------------------------------------------------------------------------")
        applogger.debug("2. " + SSH_KEY_BLOB_DECODED)
        applogger.debug("-------------------------------------------------------------------------------------")

        SSH_KEY = SSH_KEY_BLOB_DECODED.decode('utf-8')
        applogger.debug("-------------------------------------------------------------------------------------")
        applogger.debug("3. "+ SSH_KEY)
        applogger.debug("-------------------------------------------------------------------------------------")

        pkey = paramiko.RSAKey.from_private_key(io.StringIO(SSH_KEY))
        applogger.debug("-------------------------------------------------------------------------------------")
        applogger.debug("4. "+ pkey)
        applogger.debug("-------------------------------------------------------------------------------------")

        # 개발용
        # conn = pg.connect(
        #     dbname='test_hearthands',
        #     user='postgres',
        #     password='123qwe```',
        #     host='192.168.0.11',
        #     port=54332
        # )

        # with conn:
        #     cur      = conn.cursor()
        #     input_id = self.id_2.toPlainText()

        #     stmt = cur.mogrify('''
        #         SELECT A.id,
        #                 A.password,
        #                 C.w4c_code
        #             FROM common.login_user A
        #         INNER JOIN common.login_user_and_org B
        #                 ON A.login_user_idno = B.login_user_idno
        #                 AND A.id = %s
        #         INNER JOIN common.org_info C
        #                 ON B.org_idno = C.org_idno
        #                 AND C.w4c_code = %s
        #     ''', (input_id, check_w4c_cd, ))
        #     cur.execute(stmt) # PreparedStatement 실행
        #     result = cur.fetchall()

        #     return result


        with SSHTunnelForwarder(
            ssh_address_or_host=('15.165.39.46', 22),
            ssh_username='ec2-user',
            ssh_pkey=pkey,
            remote_bind_address=('127.0.0.1', 5432),
            local_bind_address=('127.0.0.1', 5432)
        ) as server:
            conn = pg.connect(
                host='ec2-15-165-39-46.ap-northeast-2.compute.amazonaws.com', 
                dbname='hearthands', 
                user='postgres', 
                password='hearthandsLive2022', 
                port=5432
            )

            with conn:
                cur      = conn.cursor()
                input_id = self.id_2.toPlainText()

                stmt = cur.mogrify('''
                    SELECT A.id,
                           A.password,
                           C.w4c_code
                      FROM common.login_user A
                    INNER JOIN common.login_user_and_org B
                            ON A.login_user_idno = B.login_user_idno
                           AND A.id = %s
                    INNER JOIN common.org_info C
                            ON B.org_idno = C.org_idno
                           AND C.w4c_code = %s
                ''', (input_id, check_w4c_cd, ))
                cur.execute(stmt) # PreparedStatement 실행

                applogger.debug("-------------------------------------------------------------------------------------")
                applogger.debug("5. execute")
                applogger.debug("-------------------------------------------------------------------------------------")

                result = cur.fetchall()

                return result
    except Exception as e:
        applogger.debug('DB Connect ERROR MSG : ', e)
        global_logger.error('DB 연결 오류!', e)


def all_check(self):
    check_file_path                = self.file_path.toPlainText()                                   # 엑셀파일
    check_w4c_cd                   = self.w4c_cd.toPlainText().replace(' ', '')                     # 희망e음 코드
    check_year                     = self.file_payroll_year_img_path.toPlainText().replace(' ', '') # 급여대장 회계연도
    check_project_num              = self.project_num.text().replace(' ', '')                       # 간편입력 사업 순서

    try:
        p = re.compile('[0-9]{1,2}')
        if p.match(check_project_num) is None:
            gui.alert('사업명 순서는 순서에 맞게 숫자만 입력해야 합니다. (2글자까지 제한)')
            return False
        

        if (
            check_file_path.replace(' ', '') != '' and
            check_w4c_cd                     != '' and
            check_year                       != '' and
            check_project_num                != ''
        ):
            p2 = re.compile('[a-zA-z0-9]')
            if (
                len(check_w4c_cd) == 11               and
                p2.match(check_w4c_cd) is not None
            ):
                # db 연결 함수                
                result = DB_setting_and_select_result(check_w4c_cd, self)

                ''' 0: id, 1: pw(hash), 2: w4c_code '''
                if (
                    len(result)              > 0            and
                    self.id_2.toPlainText() == result[0][0] and
                    check_w4c_cd            == result[0][2]
                ): 
                    return True
                    # return check_open_file(self) ## 해당 기능은 정상기능됨(엑셀 파일 굳이 열 필요 없어져서 기능은 사용안함)
                else:
                    gui.alert('입력 정보가 확인되지 않습니다. \n확인 후 다시 작업을 수행하세요.')
                    return False
            else:
                gui.alert('입력한 희망e음코드 규칙이 올바르지 않습니다. \n확인 후 다시 작업을 수행하세요.')
                return False
        else:
            gui.alert('각 정보와 희망e음 코드 확인이 어렵습니다. \n확인 후 다시 작업을 수행하세요.')
            return False 
    except Exception as e:
        gui.alert('자동화 업무 수행 전 확인단계에서 오류가 발생했습니다. \n관리자 확인이 필요합니다.')
        global_logger.error('전체 확인기능 오류: ', str(e))


def check_open_file(self):
    try:
        ################ 굳이 엑셀파일을 열 필요가 없을 것 같아서 리턴만 작업.
        ################ 기능상 문제는 없음, 만약 엑셀파일을 열어야 한다면 아래 소스를 주석해제할 것.
        file_path = self.file_path.toPlainText()
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
        logging.error('----- 해당 첨부파일 열림 확인 오류 -----', str(e))
        sys.exit()