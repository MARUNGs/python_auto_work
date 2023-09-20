#자동화 실행 전에 w4c_cd와 로그인 정보가 존재하는지 확인하는 모듈#

import os
import psycopg2 as pg
import win32com.client as win32
import re
import pyautogui as gui
import logging
import sys
import base64
import requests



app_logger = logging.getLogger('app')

def select_and_return_result(self) :
    file_path = self.file_path.toPlainText()
    w4c_cd = self.w4c_cd.toPlainText()
    year = self.file_payroll_year_img_path.toPlainText().replace(' ', '') 
    project_num = self.project_num.text().replace(' ', '')
    id_txt = self.id.toPlainText()

    p = re.compile('[0-9]{1,2}')

    if p.match(project_num) is None :
        gui.alert('사업명 순서는 숫자만 입력해야 합니다. (2글자 제한)')
        return False
    
    if (
        file_path.replace(' ', '') != '' and
        w4c_cd                     != '' and
        year                       != '' and
        id_txt                     != '' and
        project_num                != ''
    ) :
        
        p2 = re.compile('[a-zA-z0-9]')

        if (
            len(w4c_cd) == 11            and 
            p2.match(w4c_cd) is not None
        ) :
            
            url = 'http://hearthands.btog.co.kr/macro/checkUserAndW4cCode.vsj'
            body = {
                'id': id_txt.replace(' ', ''),
                'w4cCd': w4c_cd.replace(' ', '')
            }
            
            try :
                # API 호출 - POST #
                response = requests.post(url, json=body)

                if response.status_code == 200 :
                    data = response.json()['data']

                    if data['id'] == id_txt and data['w4cCd'] == w4c_cd:
                        ## 데이터가 있으면 확인완료
                        return True
                    else :
                        gui.alert('희망e음 코드와 마음손 id가 마음손 시스템에 등록되어 있지 않습니다. \n마음손 시스템에서 확인하세요.')
                        return True
                else :
                    ## 서버 API 호출이 안 되면 실패
                    gui.alert('API 호출이 실패되었습니다. \n관리자 확인이 필요합니다.')
                    return False
            except Exception as e :
                app_logger.debug(e)
                return False
