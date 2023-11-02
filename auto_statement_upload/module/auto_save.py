# 결의서/전표 자동저장 작업
import os
import pyautogui as gui                          # 운영체제 제어
from PyQt5.QtWidgets import QTableWidgetItem
import pyperclip                                 # 데이터 복사 및 붙여넣기 
import time
import logging                                   # 로그

from . import find_and_click
from . import state


# 공통 경로
img_dir_path = os.path.dirname(__file__).replace('module', 'img') + os.sep
# logger
applogger = logging.getLogger("app")

##### function ###########################################################################################
def auto_save(self, excel_obj):
    try:
        ### 입력해야 할 엑셀 데이터들을 변수로 선언
        income_list = excel_obj['income_list']
        expense_list = excel_obj['expense_list']
        personnel_expense_list = excel_obj['personnel_expense_list']

        row_index = 0 # 전체 리스트를 카운트하기 위한 변수

        ### 화면 이동 : 데이터가 존재하는지 먼저 확인하고 일반전표 데이터가 있으면 간편입력으로 화면이동
        if len(income_list)  > 0 or len(expense_list) > 0:
            if not find_and_click.img_click('회계.png') :
                success_count(self)
                return False
            
            time.sleep(0.5)

            if not find_and_click.img_click('결의및전표관리.png') :
                success_count(self)
                return False
            
            time.sleep(0.5)
            
            if not find_and_click.img_click('결의서전표간편입력.png') :
                success_count(self)
                return False
            
            gui.press('enter') # 결의서전표간편입력 메뉴 진입 엔터
            
            ### 일반전표(수입, 지출) 처리
            for row_i in range(0, len(income_list)) :
                rows = income_list[row_i]
                # 한 건씩 등록(수입)
                if auto_save_by_one(self, rows) == False : continue
                # 상태값 변경
                else : 
                    status_change_true(self, row_index)
                    row_index += 1
            
            time.sleep(1)

            for row_i in range(0, len(expense_list)) :
                rows = expense_list[row_i]
                # 한 건씩 등록(지출)
                if auto_save_by_one(self, rows) == False : continue
                # 상태값 변경
                else : 
                    status_change_true(self, row_index) 
                    row_index += 1

        if len(personnel_expense_list) > 0 :
            if not find_and_click.img_click('간편입력.png') :
                success_count(self)
                return False
            
            time.sleep(0.5)

            if not find_and_click.img_click('급여대장등록.png') :
                success_count(self)
                return False
            
            gui.press('enter')

            time.sleep(2)

            ################################################################################
            # 소스 수정. 한 번만 수행할 것이므로 밖으로 빼둠
            # 사업 딱 1번만 설정함. 어차피 같은 사업으로 처리할 것이기 때문임.

            ##### 재귀함수 처리 : 급여대장 등록 화면 오픈 Start #################################################################################################
            # 급여대장 등록 화면이 늦게 뜨는 경우가 있음. 재귀함수로 확인이 필요함.
            self.recursive_cnt = 1
            def resolution_payroll_view(cnt) :
                if cnt == 0 and not find_and_click.find_img_flag('급여대장_화면.png') :
                    applogger.debug('-- fail 급여대장_화면 오픈 확인 불가 --')
                    self.recursive_cnt += 1
                    time.sleep(1)
                    resolution_payroll_view(1) if self.recursive_cnt == 999 else resolution_payroll_view(0)
                else : return
            resolution_payroll_view(0)
            if self.recursive_cnt == 999 :
                applogger.debug(f'급여대장 등록 화면이 확인되지 않습니다. 관리자 확인이 필요합니다.')
                return False
            ##################################################################################################################################################

            if not find_and_click.img_click('급여대장_사업명.png') : return False

            gui.press('tab')

            if find_and_click.find_img_flag('급여대장_사업명_선택하세요2.png') :
                project_num = int(self.project_num.text())
                for _ in range(0, project_num): 
                    gui.press('down')
                    gui.press('enter') # 안내창 닫기

            if not find_and_click.img_click('급여대장_사업명.png') : return False
    
            time.sleep(1)

            # 행 추가 버튼
            if not find_and_click.img_click('행추가.png') : return False

            ##### 재귀함수 처리 : 직원정보검색 Start ##########################################################################################################
            # 1. 직원정보 검색 타이틀 확인
            self.recursive_cnt = 1
            def resolution_staff_search_popup(cnt) :
                if cnt == 0 and find_and_click.find_img_flag('직원정보검색_타이틀.png') == False :
                    applogger.debug('-- fail 직원정보검색 팝업 확인 불가 --')
                    self.recursive_cnt += 1
                    time.sleep(1)
                    resolution_staff_search_popup(1) if self.recursive_cnt == 999 else resolution_staff_search_popup(0)
                else : return

            resolution_staff_search_popup(0)
            if self.recursive_cnt == 999 : 
                applogger.debug(f'직원정보검색 팝업창이 확인되지 않습니다. 관리자 확인이 필요합니다.')
                return False
            
            # 2. 직원선택이 필요한데, 팝업이 로딩중인데 선택하려고 하면 다음 작업을 수행할 수 없으므로 팝업창의 선택 버튼이 활성화가 되는지 먼저 확인한다.
            # (1) 일단 한번 tab을 하여 활성화 시도를 하고,
            gui.press('tab')
            # (2) 재귀함수를 통해 버튼이 활성화되지 않았으면 계속적인 호출을 시도함.
            self.recursive_cnt = 1
            def resolution_staff_search_ok_btn_popup(cnt) :
                if cnt == 0 and not find_and_click.find_img_flag('선택2.png') :
                    applogger.debug('-- fail 직원정보검색 팝업창의 선택버튼 활성화 확인 불가 --')
                    self.recursive_cnt += 1
                    time.sleep(1)
                    resolution_staff_search_ok_btn_popup(1) if self.recursive_cnt == 999 else resolution_staff_search_ok_btn_popup(0)
                else :
                    # 활성화 확인이 되었으므로 선택처리함
                    gui.press('enter')
                    return

            resolution_staff_search_ok_btn_popup(0)
            if self.recursive_cnt == 999 : 
                applogger.debug(f'직원정보검색 팝업창의 선택 활성화 버튼이 확인되지 않습니다. 관리자 확인이 필요합니다.')
                return False
            ##### 재귀함수 처리 : 직원정보검색 End ############################################################################################################

            ##### 재귀함수 처리 : 전체선택 체크박스 Start ############################################################################################################
            self.recursive_cnt = 1
            def resolution_all_check(cnt) :
                if cnt == 0 and not find_and_click.find_img_flag('전체선택체크박스.png') :
                    applogger.debug('-- fail 급여대장 전체선택 체크박스 확인 불가 --')
                    self.recursive_cnt += 1
                    time.sleep(1)
                    resolution_all_check(1) if self.recursive_cnt == 999 else resolution_all_check(0)
                else :
                    # 전체선택 체크박스 선택
                    find_and_click.img_click('전체선택체크박스.png')
                    return

            resolution_all_check(0)
            if self.recursive_cnt == 999 : 
                applogger.debug(f'전체선택 체크박스 확인이 어렵습니다. 관리자 확인이 필요합니다.')
                return False
            ##### 재귀함수 처리 : 전체선택 체크박스 End ############################################################################################################

            # 재귀함수 사용 전 소스코드
            # 대상자 선택(앞줄 1명)
            # gui.press('tab')
            # gui.press('enter')

            # if not find_and_click.img_click('전체선택체크박스.png') : return False

            time.sleep(1)
            ################################################################################

            # 인건비(지출) 처리
            for row_i in range(0, len(personnel_expense_list)) :
                rows = personnel_expense_list[row_i]
                # 한 건씩 등록(인건비(지출))
                if auto_save_payroll_one(self, rows) == False : continue
                # 상태값 변경
                else : 
                    status_change_true(self, row_index)
                    row_index += 1
        
        success_count(self)
    except Exception as e:
        logging.error('전표정보 자동저장(auto_save) Exception : ', str(e))



# 한 건씩 처리하기
def auto_save_by_one(self, rows) :
    time.sleep(1)
    ##### 재귀함수 처리 : 조회 #######################################################################################################
    # 일단 조회
    gui.hotkey('alt', 'f3') # 조회
    # 조회하고 난 뒤에 내부오류 발생이 일어날 수도 있음. enter로 메세지박스를 끄고 다시 조회
    def search_inner_error(cnt) :
        if cnt == 0 :
            if find_and_click.find_img_flag('내부서버오류입니다.png') :
                time.sleep(0.5)
                gui.press('enter')
                gui.hotkey('alt', 'f3') # 한번더 조회
                search_inner_error(0)
            elif find_and_click.find_down_img_flag('loading.png') :
                time.sleep(1)
                gui.hotkey('alt', 'f3')
                search_inner_error(0)
            else : return
        else : return
    
    search_inner_error(0)
    # self.recursive_cnt = 0
    # def search_count(cnt) :
    #     if cnt == 0 and not find_and_click.find_img_flag('조회총건수.png') : 
    #         applogger.debug('-- fail 조회 --')
    #         self.recursive_cnt += 1
    #         time.sleep(1)
    #         gui.hotkey('alt', 'f3') # 한번더 조회
    #         search_count(1) if self.recursive_cnt == 999 else search_count(0)
    #     else : return
    # search_count(0)
    # if self.recursive_cnt == 999 :
    #     applogger.debug(f'해당 전표는 등록하지 못하였습니다. (결의번호 : {rows[10]})')
    #     return False
    #################################################################################################################################
    gui.sleep(1)
    for _ in range(0,3) : gui.hotkey('alt', 'f2') # 신규

    ##### 재귀함수 처리 : 신규 #######################################################################################################
    # 조회하는 동안 응답대기 현상이 일어날 수 있으므로 '신규' 단축키를 이용했을 때 여전히 요소들이 비활성화되어 있으면 단축키를 다시 실행한다.
    def new_box(cnt) :
        if cnt == 0 and find_and_click.find_img_flag('신규비활성화.png') :
            applogger.debug('-- fail 신규화면 표시 -- ')
            time.sleep(1)
            gui.hotkey('alt', 'f2') # 신규
            new_box(0)
        else : return

    new_box(0)
    ##################################################################################################################################
    time.sleep(3)
    find_and_click.img_right_click('결의일자타이틀.png')
    gui.hotkey('ctrl', 'a')
    gui.press('backspace')
    pyperclip.copy(rows[3])
    gui.hotkey('ctrl', 'v')

    for _ in range(0,2) : gui.press('tab')

    if rows[0] == '수입' :
        gui.press('down')
    else :
        gui.press('down')
        gui.press('down')

    gui.press('tab')

    project_num = int(self.project_num.text())
    if project_num != None :
        for _ in range(0, project_num) : gui.press('down')
    
    gui.press('tab')
    gui.hotkey('ctrl', 'a')
    gui.press('backspace')

    if rows[0] == '수입' :
        if rows[1] == '반납' : pyperclip.copy('-' + rows[7])
        else : pyperclip.copy(rows[6])
    else :
        if rows[1] == '반납' : pyperclip.copy('-' + rows[6]) ## '반납'인 경우, 마이너스 금액으로 작성
        else : pyperclip.copy(rows[7]) ## '정상'인 경우, 금액 그대로 작성

    time.sleep(0.5)
    gui.hotkey('ctrl', 'v')
    time.sleep(0.5)

    for _ in range(0,3) : gui.press('tab')

    gui.press('enter')

    ##### 재귀함수 처리 : 계정과목 팝업 Start ##############################################################################################
    self.recursive_cnt = 1 # 재귀함수 호출 수
    def popup_check_msgbox(cnt) :
        if cnt == 0 :
            if rows[0] == '수입' : 
                if find_and_click.find_img_flag('계정코드팝업_세입.png') == False :
                    applogger.debug('-- fail 계정코드팝업_세입 --')
                    self.recursive_cnt += 1
                    time.sleep(1)
                    popup_check_msgbox(1) if self.recursive_cnt == 999 else popup_check_msgbox(0)
                else : return
            else :
                if find_and_click.find_img_flag('계정코드팝업_세출.png') == False :
                    applogger.debug('-- fail 계정코드팝업_세출 --')
                    self.recursive_cnt += 1
                    time.sleep(1)
                    popup_check_msgbox(1) if self.recursive_cnt == 999 else popup_check_msgbox(0)
                else : return
        else : return

    popup_check_msgbox(0) # 재귀함수 호출
    if self.recursive_cnt == 999 : # 재귀함수 호출 수가 999번이 될 동안 이미지를 찾지 못했다면 해당 1 row 작업은 넘어가도록 한다.
        find_and_click.img_click('엑스박스.png')
        applogger.debug(f'해당 전표는 등록하지 못하였습니다. (결의번호 : {rows[10]})')
        return False
    ##### 재귀함수 처리 : 계정과목 팝업 End ################################################################################################
    
    time.sleep(1)


    ##### 재귀함수 처리 : 계정과목 팝업의 선택버튼 Start ######################################################################################
    # (1) 일단 한번 tab을 하여 활성화 시도를 하고,
    gui.press('tab')
    # (2) 재귀함수를 통해 활성화되지 않았으면 계속적인 호출을 시도함
    self.recursive_cnt = 1
    def popup_ok_box(cnt) :
        if cnt == 0 and find_and_click.find_img_flag('선택2.png') == False :
            applogger.debug('-- fail popup ok btn -- ')
            self.recursive_cnt += 1
            time.sleep(1)
            popup_ok_box(1) if self.recursive_cnt == 999 else popup_ok_box(0)
        else : return

    # [2-1] 재귀함수 시작
    popup_ok_box(0)
    # [2-2] 재귀함수의 호출 수가 999번이 될 동안 이미지를 찾지 못했다면 해당 1 row 작업은 넘어가도록 한다.
    if self.recursive_cnt == 999 : 
        find_and_click.img_click('엑스박스.png')
        applogger.debug(f'해당 전표는 등록하지 못하였습니다. (결의번호 : {rows[10]})')
        return False
    ##### 재귀함수 처리 : 계정과목 팝업의 선택버튼 End ########################################################################################
    
    for _ in range(0,3) : gui.press('tab')

    # for _ in range(0,4) : gui.press('tab')

    if rows[4] == '공공요금 및 제세공과금' : pyperclip.copy('공공요금 및 각종 세금공과금')
    else : pyperclip.copy(rows[4])

    gui.hotkey('ctrl', 'v')
    
    time.sleep(0.5)

    ##### 재귀함수 처리 : 계정과목 팝업에서 조회했을 때 로딩 확인 Start #######################################################################################
    # 먼저 enter 작업
    gui.press('enter')
    time.sleep(1)
    self.recursive_cnt = 0
    def account_subject_popup_loading(cnt) :
        if cnt == 0 and find_and_click.find_img_flag('조회된데이터가없습니다.png') :
            applogger.debug('-- fail 계정과목 조회 후 로딩 확인 불가 --')
            self.recursive_cnt += 1
            time.sleep(1)
            account_subject_popup_loading(1) if self.recursive_cnt == 999 else account_subject_popup_loading(0)
        else : return
        
    account_subject_popup_loading(0)
    if self.recursive_cnt == 999 :
        find_and_click.img_click('엑스박스.png')
        applogger.debug(f'해당 전표는 등록하지 못하였습니다. (결의번호 : {rows[10]})')
        return False
    #######################################################################################################################################################
    
    ##### 재귀함수 처리 : 계정과목 팝업에서 조회 후 선택2 활성화 확인 Start #######################################################################################
    self.recursive_cnt = 0
    gui.press('tab') 
    popup_ok_box(0)
    if self.recursive_cnt == 999 : 
        find_and_click.img_click('엑스박스.png')
        applogger.debug(f'해당 전표는 등록하지 못하였습니다. (결의번호 : {rows[10]})')
        return False
    gui.press('enter')
    #######################################################################################################################################################


    # for _ in range(0,2):
    #     gui.press('enter')
    #     time.sleep(1)

    ##### 재귀함수 처리 : 대상자 팝업 Start ##################################################################################################
    if rows[4] == '본인부담금수입' :
        if find_and_click.img_right_in_click('대상자.png') == False : return False

        self.recursive_cnt = 1
        def target_popup_check_msgbox(cnt) :
            if cnt == 0 and find_and_click.find_img_flag('대상자팝업.png') == False :
                applogger.debug('-- fail 대상자 팝업 --')
                self.recursive_cnt += 1
                time.sleep(1)
                target_popup_check_msgbox(1) if self.recursive_cnt == 999 else target_popup_check_msgbox(0)
            else : return

        target_popup_check_msgbox(0)
        if self.recursive_cnt == 999 : 
            find_and_click.img_click('엑스박스.png')
            applogger.debug(f'해당 전표는 등록하지 못하였습니다. (결의번호 : {rows[10]})')
            return False
    ##### 재귀함수 처리 : 대상자 팝업 End ####################################################################################################

        time.sleep(0.7)
        # if find_and_click.img_right_in_click('대상자.png') == False : return False
        # time.sleep(1)
        gui.press('enter')
    
    time.sleep(1)
    
    ##### 재귀함수 처리 : 상대계정 팝업 Start #################################################################################################
    if rows[9] != '' :
        if not find_and_click.img_right_in_click('상대계정박스.png') :
            self.recursive_cnt = 1
            def oppo_popup_check_msgbox(cnt) :
                if cnt == 0 and find_and_click.find_img_flag('상대계정팝업.png') == False :
                    applogger.debug('-- fail 상대계정팝업 --')
                    self.recursive_cnt += 1
                    time.sleep(1)
                    oppo_popup_check_msgbox(1) if self.recursive_cnt == 999 else oppo_popup_check_msgbox(0)
                else : return

            oppo_popup_check_msgbox(0)
            if self.recursive_cnt == 999 :
                find_and_click.img_click('엑스박스.png')
                applogger.debug(f'해당 전표는 등록하지 못하였습니다. (결의번호 : {rows[10]})')
                return False
    ##### 재귀함수 처리 : 상대계정 팝업 End ###################################################################################################
        
        time.sleep(0.7)

    ##### 재귀함수 처리 : 상대계정 팝업의 선택박스 Start #######################################################################################
        # 상대계정과목 팝업이 오랫동안 활성화되지 않을 수도 있음. 재귀함수를 통해 활성화됐는지 확인할 것.
        # (1) 일단 한번 tab을 하여 활성화 시도를 하고,
        gui.press('tab')
        # (2) 재귀함수를 통해 활성화되지 않았으면 계속적인 호출을 시도함
        self.recursive_cnt = 1
        def popup_ok_box(cnt) :
            if cnt == 0 and find_and_click.find_img_flag('선택2.png') == False :
                applogger.debug('-- fail 상대계정과목 선택버튼 -- ')
                self.recursive_cnt += 999
                time.sleep(1)
                popup_ok_box(1) if self.recursive_cnt == 999 else popup_ok_box(0)
            else : return

        popup_ok_box(0)
        if self.recursive_cnt == 999 :
            find_and_click.img_click('엑스박스.png')
            applogger.debug(f'해당 전표는 등록하지 못하였습니다. (결의번호 : {rows[10]})')
            return False
    ########################################################################################################################################

        for _ in range(0,2) : gui.press('tab')

        pyperclip.copy(rows[9])
        gui.hotkey('ctrl', 'v')
        
    ##### 재귀함수 처리 : 상대계정 팝업에서 조회했을 때 로딩 확인 Start #######################################################################################
        # 먼저 enter 작업
        gui.press('enter')
        time.sleep(1)
        self.recursive_cnt = 0
        def oppo_account_subject_popup_loading(cnt) :
            if cnt == 0 and find_and_click.find_img_flag('조회된데이터가없습니다.png') :
                applogger.debug('-- fail 상대계정과목 조회 후 로딩 확인 불가 --')
                self.recursive_cnt += 1
                time.sleep(1)
                oppo_account_subject_popup_loading(1) if self.recursive_cnt == 999 else oppo_account_subject_popup_loading(0)
            else : 
                gui.press('enter')
                return
            
        oppo_account_subject_popup_loading(0)
        if self.recursive_cnt == 999 :
            find_and_click.img_click('엑스박스.png')
            applogger.debug(f'해당 전표는 등록하지 못하였습니다. (결의번호 : {rows[10]})')
            return False
    #######################################################################################################################################################
        
        # for _ in range(0,2) :
        #     gui.press('enter')
        #     time.sleep(0.5)

    time.sleep(1)

    ##### 재귀함수 처리 : 적요 input Start ###################################################################################################
    if len(rows[5]) > 0 :
        self.recursive_cnt = 1
        def summary_check(cnt) :
            if cnt == 0 and find_and_click.find_img_flag('결의서적요.png') == False :
                applogger.debug('-- fail summary check --')
                self.recursive_cnt += 1
                time.sleep(1)
                summary_check(1) if self.recursive_cnt == 999 else summary_check(0)
            else :
                find_and_click.img_click('결의서적요.png')
                return
            
        summary_check(0)
        if self.recursive_cnt == 999 : 
            applogger.debug(f'해당 전표는 등록하지 못하였습니다. (결의번호 : {rows[10]})')
            return False
    ########################################################################################################################################
        
        time.sleep(0.7)

        pyperclip.copy(rows[5])
        gui.hotkey('ctrl', 'v')
        gui.press('tab')

    time.sleep(1)

    ##### 재귀함수 처리 : 저장하시겠습니까 Start ##############################################################################################
    gui.hotkey('alt', 'f8') # 저장
    # 내부서버오류입니다 안내메세지가 발견될 수 있으므로 재귀함수로 저장 기능을 수행할 것.
    # 원인1 : 단축키나 버튼을 클릭했을 때 오류 안내메세지가 발생
    # 해결 : 한번 더 단축키 사용
    self.recursive_cnt = 1
    def save_confirm_box(cnt) :
        if cnt == 0 and not find_and_click.find_img_flag('저장하시겠습니까.png') :
            applogger.debug('-- fail 저장하시겠습니까 --')
            self.recursive_cnt += 1
            time.sleep(1)
            # 한번더 실행
            gui.hotkey('alt', 'f8') # 저장
            save_confirm_box(1) if self.recursive_cnt == 999 else save_confirm_box(0)
        else :
            # 찾으면 엔터
            gui.press('enter')
            time.sleep(3)
            # 여기서 내부서버 오류 발생됨. 엔터를 쳐서 메세지박스를 끄고 다시 재귀함수 호출.
            if find_and_click.find_img_flag('내부서버오류입니다.png') :
                time.sleep(0.5)
                gui.press('enter')
                gui.hotkey('alt', 'f8') # 한번더 저장
                save_confirm_box(0)
            else : return
        
    save_confirm_box(0)
    if self.recursive_cnt == 999 :
        # 어떠한 안내메세지를 띄우는 로직은 아니므로 다음 건수로 넘어간다.
        applogger.debug(f'해당 전표는 등록하지 못하였습니다. (결의번호 : {rows[10]})')
        return False
    #########################################################################################################################################
    ##### 재귀함수 처리 : 저장 Start ##########################################################################################################
    self.recursive_cnt = 1
    def save_check_msgbox(cnt) :
        if cnt == 0 and find_and_click.find_img_flag('성공적으로저장하였습니다.png') == False :
            applogger.debug('-- fail 저장 --')
            self.recursive_cnt += 1
            time.sleep(1)
            save_check_msgbox(1) if self.recursive_cnt == 999 else save_check_msgbox(0)
        else :
            gui.press('enter')
            return

    save_check_msgbox(0)
    if self.recursive_cnt == 999 : 
        applogger.debug(f'해당 전표는 등록하지 못하였습니다. (결의번호 : {rows[10]})')
        return False
    ########################################################################################################################################



# 한건씩 처리(인건비(지출))
def auto_save_payroll_one(self, rows):
    time.sleep(1)

    if not find_and_click.img_click('급여대장_사업명.png') : return False
    for _ in range(0,2) : gui.press('tab')
    gui.press('enter')

    time.sleep(0.5)

    if not find_and_click.customization_payroll_year_img_click(self) : return False
    if not find_and_click.img_click('회계연도_선택.png') : return False

    time.sleep(1)

    ##### 재귀함수 처리 : 인건비(지출) 결의서 팝업 Start ##############################################################################################
    self.recursive_cnt = 1
    def resolution_popup_box(cnt) :
        pass
        if cnt == 0 and not find_and_click.find_img_flag('급여대장_결의일자.png') :
            applogger.debug('-- fail 급여대장_결의일자 --')
            self.recursive_cnt += 1
            time.sleep(1)
            resolution_popup_box(1) if self.recursive_cnt == 999 else resolution_popup_box(0)
        else :
            # 결의일자 찾으면 클릭
            find_and_click.img_right_click('급여대장_결의일자.png')
            return

    resolution_popup_box(0)
    if self.recursive_cnt == 999 :
        find_and_click.img_click('엑스박스.png')
        applogger.debug(f'해당 전표는 등록하지 못하였습니다. (결의번호 : {rows[10]})')
        return False
    ################################################################################################################################################

    # 결의일자
    # if not find_and_click.img_right_click('급여대장_결의일자.png') : return False # 기존소스 대신 재귀함수 이용
    time.sleep(1)
    gui.hotkey('ctrl', 'a')
    gui.press('backspace')
    pyperclip.copy(rows[3])
    gui.hotkey('ctrl', 'v')

    # 행추가 버튼
    if not find_and_click.img_click('행추가.png') : return False

    ##### 재귀함수 처리 : 인건비(지출) 결의서의 계정과목 팝업 Start ##############################################################################################
    self.recursive_cnt = 1
    def resolution_account_subject_box(cnt) :
        if cnt == 0 and not find_and_click.find_img_flag('급여대장_계정과목_타이틀.png') :
            applogger.debug('-- fail 급여대장_계정과목_타이틀 --')
            self.recursive_cnt += 1
            time.sleep(1)
            resolution_account_subject_box(1) if self.recursive_cnt == 999 else resolution_account_subject_box(0)
        else : 
            # 찾으면 클릭
            find_and_click.img_bottom_right_in_click('급여대장_계정과목_타이틀.png')
            return

    resolution_account_subject_box(0)
    if self.recursive_cnt == 999 :
        find_and_click.img_right_in_click('계정코드박스_타이틀.png') # 팝업이 2개나 띄워져 있기 때문에 엑스박스.png로 찾아서 클릭하기엔 무리가 있음.
        applogger.debug(f'해당 전표는 등록하지 못하였습니다. (결의번호 : {rows[10]})')
        return False
    ##########################################################################################################################################################
    # 계정명 : 직접 데이터를 입력하면 안내창이 뜨므로, 아이콘을 눌러서 처리해야 한다.
    # if not find_and_click.img_bottom_right_in_click('급여대장_계정과목_타이틀.png') : return False # 기존소스 대신 재귀함수 이용
    # time.sleep(2)


    ##### 재귀함수 처리 : 인건비(지출) 결의서의 계정과목 팝업의 선택박스 활성화 Start ###############################################################################
    # 1. 먼저 tab을 눌러 버튼 활성화를 해두고
    gui.press('tab')
    # 2. 재귀함수를 통해 해당 이미지가 존재하는지 확인하여 있으면 계정과목 처리 완료
    self.recursive_cnt = 1
    def resolution_account_subject_ok_box(cnt) :
        if cnt == 0 and not find_and_click.find_img_flag('선택2.png') :
            applogger.debug('-- fail 계정과목 팝업창의 선택 버튼 활성화 확인 불가 --')
            self.recursive_cnt += 1
            time.sleep(1)
            resolution_account_subject_ok_box(1) if self.recursive_cnt == 999 else resolution_account_subject_ok_box(0)
        else : return

    resolution_account_subject_ok_box(0)
    if self.recursive_cnt == 999 :
        find_and_click.img_right_in_click('계정코드박스_타이틀.png') # 팝업이 2개나 띄워져 있기 때문에 엑스박스.png로 찾아서 클릭하기엔 무리가 있음.
        applogger.debug(f'해당 전표는 등록하지 못하였습니다. (결의번호 : {rows[10]})')
        return False
    ############################################################################################################################################################

    # 선택 버튼 활성화 확인되면 나머지 처리
    for _ in range(0,3) : gui.press('tab')
    pyperclip.copy(rows[4])
    gui.hotkey('ctrl', 'v')
    # for _ in range(0,2) : gui.press('enter') # 확인, 창닫기      

    ##### 재귀함수 처리 : 계정과목 팝업에서 조회했을 때 로딩 확인 Start #######################################################################################
    # 먼저 enter 작업
    gui.press('enter')
    time.sleep(1)
    self.recursive_cnt = 0
    def resolution_account_subject_popup_loading(cnt) :
        if cnt == 0 and find_and_click.find_img_flag('조회된데이터가없습니다.png') :
            applogger.debug('-- fail 계정과목 조회 후 로딩 확인 불가 --')
            self.recursive_cnt += 1
            time.sleep(1)
            resolution_account_subject_popup_loading(1) if self.recursive_cnt == 999 else resolution_account_subject_popup_loading(0)
        else : 
            # gui.press('enter')
            return
        
    resolution_account_subject_popup_loading(0)
    if self.recursive_cnt == 999 :
        find_and_click.img_right_in_click('계정코드박스_타이틀.png')
        applogger.debug(f'해당 전표는 등록하지 못하였습니다. (결의번호 : {rows[10]})')
        return False
    #######################################################################################################################################################

    ##### 재귀함수 처리 : 급여대장의 계정과목 팝업에서 조회 후 선택2 활성화 확인 Start #######################################################################################
    self.recursive_cnt = 0
    gui.press('tab')
    resolution_account_subject_ok_box(0)
    if self.recursive_cnt == 999 : 
        find_and_click.img_right_in_click('계정코드박스_타이틀.png')
        applogger.debug(f'해당 전표는 등록하지 못하였습니다. (결의번호 : {rows[10]})')
        return False
    gui.press('enter')
    #######################################################################################################################################################


    # for _ in range(0,4): gui.press('tab')
    # pyperclip.copy(rows[4])
    # gui.hotkey('ctrl', 'v')
    # for _ in range(0,2) : gui.press('enter') # 확인, 창닫기

    # 금액
    for _ in range(0,3) : gui.press('tab')
    pyperclip.copy(rows[7]) # 지출금액만 취급하므로 idx 7번만 사용함.
    gui.hotkey('ctrl', 'v')

    gui.press('tab')

    # 적요
    if len(rows[5]) > 0 : 
        pyperclip.copy(rows[5])
        gui.hotkey('ctrl', 'v')

    gui.press('tab')

    # 상대계정
    pyperclip.copy(rows[9]) # 지출금액만 취급하므로 idx 9번만 사용함.
    gui.hotkey('ctrl', 'v')
    gui.press('tab')

    # time.sleep(2)

    ##### 재귀함수 처리 : 인건비(지출) 결의서의 상대계정과목 팝업 Start ##############################################################################################
    self.recursive_cnt = 1
    def resolution_oppo_account_subject_box(cnt) :
        if cnt == 0 and not find_and_click.find_img_flag('상대계정박스_타이틀.png') :
            applogger.debug('-- fail 상대계정박스_타이틀 --')
            self.recursive_cnt += 1
            time.sleep(1)
            resolution_oppo_account_subject_box(1) if self.recursive_cnt == 999 else resolution_oppo_account_subject_box(0)
        else : return

    resolution_oppo_account_subject_box(0)
    if self.recursive_cnt == 999 :
        find_and_click.img_click('상대계정박스_타이틀.png') # 팝업이 2개나 띄워져 있기 때문에 엑스박스.png로 찾아서 클릭하기엔 무리가 있음.
        applogger.debug(f'해당 전표는 등록하지 못하였습니다. (결의번호 : {rows[10]})')
        return False
    ##############################################################################################################################################################

    ##### 재귀함수 처리 : 인건비(지출) 결의서의 상대계정과목 팝업 선택박스 활성화 Start ################################################################################
    # 1. 먼저 선택 버튼을 활성화 처리한다.
    gui.press('tab')
    # 2. 재귀함수를 통해 버튼을 확인한다.
    self.recursive_cnt = 1
    def resolution_oppo_account_subject_ok_box(cnt) :
        if cnt == 0 and not find_and_click.find_img_flag('선택2.png') :
            applogger.debug('-- fail 상대계정과목 팝업창의 선택 버튼 활성화 확인 불가 --')
            self.recursive_cnt += 1
            time.sleep(1)
            resolution_oppo_account_subject_ok_box(1) if self.recursive_cnt == 999 else resolution_oppo_account_subject_ok_box(0)
        else : return

    resolution_oppo_account_subject_ok_box(0)
    if self.recursive_cnt == 999 :
        find_and_click.img_click('상대계정박스_타이틀.png') # 팝업이 2개나 띄워져 있기 때문에 엑스박스.png로 찾아서 클릭하기엔 무리가 있음.
        applogger.debug(f'해당 전표는 등록하지 못하였습니다. (결의번호 : {rows[10]})')
        return False
    ##############################################################################################################################################################

    if rows[9] == '장기요양급여수입' or rows[9].replace(' ', '') == '가산금수입': gui.press('enter')

    ##### 재귀함수 처리 : 인건비(지출) 결의서의 저장 및 팝업닫기 End   ##############################################################################################
    # 1. 저장 ing
    self.recursive_cnt = 1
    def resolution_save(cnt) :
        if cnt == 0 and not find_and_click.find_img_flag('급여대장_저장.png') :
            applogger.deubg('-- fail 급여대장_저장 --')
            self.recursive_cnt += 1
            time.sleep(1)
            resolution_save(1) if self.recursive_cnt == 999 else resolution_save(0)
        else :
            # 저장기능 수행
            find_and_click.img_click('급여대장_저장.png')
            gui.press('enter')
            return

    resolution_save(0)
    if self.recursive_cnt == 999 :
        find_and_click.img_click('엑스박스.png')
        applogger.debug(f'해당 전표는 등록하지 못하였습니다. (결의번호 : {rows[10]})')
        return False
    
    # 2. 성공적으로 저장하였습니다 안내문구 확인
    self.recursive_cnt = 1
    def resolution_save_check(cnt) :
        if cnt == 0 and find_and_click.find_img_flag('성공적으로저장하였습니다.png') == False :
            applogger.debug('-- fail 성공적으로 저장하였습니다 안내문구 미확인 --')
            self.recursive_cnt += 1
            time.sleep(1)
            resolution_save_check(1) if self.recursive_cnt == 999 else resolution_save_check(0)
        else :
            gui.press('enter')
            return
        
    resolution_save_check(0)
    if self.recursive_cnt == 999 :
        find_and_click.img_click('엑스박스.png')
        applogger.debug(f'해당 전표는 등록하지 못하였습니다. (결의번호 : {rows[10]})')
        return False
    
    time.sleep(1)

    # 3. 팝업창 닫기
    for _ in range(0,5) : gui.press('tab')
    gui.press('enter') # 닫기

    time.sleep(1)
    ##### 재귀함수 처리 : 인건비(지출) 결의서의 저장 및 팝업닫기 End   ##############################################################################################
    
    # 기존소스
    # 마무리
    # if not find_and_click.img_click('급여대장_저장.png') : return False
    # gui.press('enter')
    # time.sleep(5)
    # gui.press('enter')
    # time.sleep(1.5)
    # if not find_and_click.img_click('급여대장_닫기.png') : return False
    # time.sleep(1.5)


        


    








##### 성능을 위하여 리팩토링
def auto_save_simple(self, excel_list):
    for row_i in range(0, len(excel_list)) :
        # for문 대신 컨트롤할 변수
        rows = excel_list[row_i]

        for _ in range(0, 2): 
            gui.hotkey('alt', 'f3') # 조회
            time.sleep(0.35)

        for _ in range(0, 2): 
            gui.hotkey('alt', 'f2') # 신규
            time.sleep(0.35)
        
        #결의일자#
        gui.hotkey('ctrl', 'a')
        gui.press('backspace')
        pyperclip.copy(rows[3])
        gui.hotkey('ctrl', 'v')

        for _ in range(0, 2): 
            gui.press('tab')
            time.sleep(0.35)
        
        #결의구분#
        if rows[0] == '수입':
            gui.press('down')
        else: 
            gui.press('down')
            gui.press('down')

        time.sleep(0.35)
        gui.press('tab')

        #사업 ---> 이거 하려면 화면단부터 수정해야 함#
        project_num = int(self.project_num.text())
        if project_num != None: 
            for _ in range(0, project_num): 
                gui.press('down')
                time.sleep(0.35)

        gui.press('tab')

        #금액
        # - 반납구분이 '반납'인 경우 '-' 붙여서 금액 입력할 것.
        gui.hotkey('ctrl', 'a')
        gui.press('backspace')
        
        # '수입'일 때
        if rows[0] == '수입':
            ## '반납'인 경우, 마이너스 금액으로 작성
            if rows[1] == '반납': pyperclip.copy('-' + rows[7])
            ## '정상'인 경우, 금액 그대로 작성
            else: pyperclip.copy(rows[6])
        # '지출'일 때
        else:
            ## '반납'인 경우, 마이너스 금액으로 작성
            if rows[1] == '반납': pyperclip.copy('-' + rows[6])
            ## '정상'인 경우, 금액 그대로 작성
            else: pyperclip.copy(rows[7])

        time.sleep(0.5)
        gui.hotkey('ctrl', 'v')
        
        #계정코드#
        for _ in range(0, 3): 
            gui.press('tab')
            time.sleep(0.35)

        gui.press('enter') # 팝업창 오픈

        for _ in range(0, 4): 
            gui.press('tab')
            time.sleep(0.35)

        #계정코드명 수정: W4C 프로그램에서는 '공공요금 및 제세공과금'에 대한 과목명이 다르게 관리되므로 별도로 변경함
        if rows[4] == '공공요금 및 제세공과금':
            pyperclip.copy('공공요금 및 각종 세금공과금')
        else:
            pyperclip.copy(rows[4])

        gui.hotkey('ctrl', 'v') # 계정코드 입력

        for _ in range(0, 2): 
            gui.press('enter')
            time.sleep(0.35)

        #대상자#
        if rows[4] == '본인부담금수입':
            for _ in range(0, 3): 
                gui.press('tab')
                time.sleep(0.35)

            for _ in range(0, 2): 
                gui.press('enter') # 팝업창 오픈, 선택까지
                time.sleep(0.35)

        #상대계정(지출결의서)#
        if rows[9] != '':
            for _ in range(0, 3): 
                gui.press('tab')
                time.sleep(0.35)

            gui.press('enter') # 팝업창 오픈

            for _ in range(0, 3): 
                gui.press('tab')
                time.sleep(0.35)

            pyperclip.copy(rows[9])
            gui.hotkey('ctrl', 'v')

            for _ in range(0, 2): 
                gui.press('enter') # 무조건 첫번째 상대계정 선택
                time.sleep(1)

        #적요#
        if len(rows[5]) > 0:
            for _ in range(0, 6): 
                gui.press('tab')
                time.sleep(0.35)

            pyperclip.copy(rows[5])
            gui.hotkey('ctrl', 'v')

            time.sleep(0.35)

            gui.press('tab')

        gui.sleep(0.5)

        #1건 저장 프로세스#
        gui.hotkey('alt', 'f8') # 저장
        gui.sleep(0.35)

        for _ in range(0, 2):
            gui.press('enter')
            gui.sleep(0.35)
        
        #성공 확인됨. Flag값 수정하기#
        status_change_true(self, rows)
    else : return False





def auto_save_payroll(self, excel_list):
    # 사업을 딱 1번만 설정함. 어차피 같은 사업처리할 것이기 때문.
    find_and_click.img_click('급여대장_사업명.png')
    gui.press('tab')

    if find_and_click.find_img_flag('급여대장_사업명_선택하세요2.png'):
        project_num = int(self.project_num.text())
        for _ in range(0, project_num): 
            gui.press('down')
            gui.press('enter') # 안내창 닫기
    
    gui.sleep(0.5)

    #데이터 수 만큼 움직이기#
    for row_i in range(0, len(excel_list)):
        rows = excel_list[row_i]

        gui.hotkey('alt', 'f3') # 조회

        #행추가#
        find_and_click.img_click('급여대장_사업명.png')
        gui.keyDown('shift')
        for _ in range(0, 4): gui.press('tab')
        gui.keyUp('shift')
        gui.press('enter')

        gui.sleep(0.7) # 팝업창 오픈

        #대상자 선택버튼 클릭#
        gui.press('tab')
        gui.press('enter')

        #전체 선택박스 체크#
        find_and_click.img_click('전체선택체크박스.png')

        #회계연도 선택#
        find_and_click.img_click('급여대장_사업명.png')
        for _ in range(0, 2): gui.press('tab')
        gui.press('enter')
        find_and_click.customization_payroll_year_img_click(self)
        find_and_click.img_click('회계연도_선택.png')

        time.sleep(1.0) #팝업창 오픈
        
        #결의일자#
        for _ in range(0, 6): gui.press('tab')
        gui.hotkey('ctrl', 'a')
        gui.press('backspace')
        pyperclip.copy(rows[3])
        gui.hotkey('ctrl', 'v')

        #행추가#
        for _ in range(0, 10): gui.press('tab')
        gui.press('enter')

        #계정명 : 직접 데이터를 입력하면 안내창이 뜨므로, 아이콘을 눌러서 처리해야 한다.#
        find_and_click.img_bottom_right_in_click('급여대장_계정과목_타이틀.png')
        for _ in range(0, 4): gui.press('tab')
        pyperclip.copy(rows[4])
        gui.hotkey('ctrl', 'v')
        for _ in range(0, 2): gui.press('enter') # 확인, 창닫기

        #금액#
        for _ in range(0, 3): gui.press('tab')
        pyperclip.copy(rows[7]) # 지출금액만 취급하므로 idx 7번만 사용함.
        gui.hotkey('ctrl', 'v')

        gui.press('tab')

        #적요#
        if len(rows[5]) > 0:
            pyperclip.copy(rows[5])
            gui.hotkey('ctrl', 'v')

        gui.press('tab')

        #상대계정#
        pyperclip.copy(rows[9]) # 지출금액만 취급하므로 idx 7번만 사용함.
        gui.hotkey('ctrl', 'v')
        gui.press('tab')

        gui.sleep(0.5)
        if rows[9] == '장기요양급여수입': gui.press('enter')
        gui.sleep(1.0)


        #1건 마무리 프로세스#
        find_and_click.img_click('급여대장_저장.png')
        for _ in range(0, 2):
            gui.press('enter') # 팝업창 확인
            gui.sleep(0.2)
        # 급여대장 지출결의서 닫기
        for _ in range(0, 5): gui.press('tab')
        gui.press('enter')

        #성공 확인됨. Flag값 수정하기#
        status_change_true(self, rows)











# 상태 테이블값 true로 변환
def status_change_true(self, idx):
    # data      = idx
    # excel_tb  = self.excel_tb
    status_tb = self.status_tb

    # 굳이 for문 돌지 않고 status_tb의 해당 idx에 'Success'로 변경하기
    status_tb.setItem(idx, 0, QTableWidgetItem('Success'))

    # excel_tb 위젯을 for문 돌아서 결의번호(10)가 동일하면 그 index를 이용하여 status_tb의 fail를 success로 변경하기
    # for r_idx in range(status_tb.rowCount()):
    #     if data[10] == excel_tb.item(r_idx, 10).text(): status_tb.setItem(r_idx, 0, QTableWidgetItem('Success'))

# success, fail 체크하여 안내 메세지 띄우기
def success_count(self) :
    # Success 갯수 체크
    status_tb     = self.status_tb
    success_count = 0
    fail_count    = 0
    row_count     = status_tb.rowCount()

    for idx in range(0, row_count):
        if status_tb.item(idx,0).text() in 'Success' :
            success_count += 1
        else: 
            fail_count += 1

    gui.alert(f'전표정보 자동업로드 작업을 마무리합니다. \n성공횟수: {success_count} \n실패횟수: {fail_count}')
    state.running_stop_flag = 'stop'