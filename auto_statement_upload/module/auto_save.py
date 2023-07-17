# 결의서/전표 자동저장 작업
import os
import pyautogui as gui # 운영체제 제어
from PyQt5.QtWidgets import QTableWidgetItem
import pyperclip # 데이터 복사 및 붙여넣기 
import time
import logging # 로그
from enum import Enum


# 공통 경로
imgDirPath = os.path.dirname(__file__).replace('module', 'img') + os.sep


'''
    @param self      # PyQT5
    @param excelList # makeExcelData()를 통해 갖고있는 데이터
'''
def autoSave(self, excelList):
    try:
        maxColumnCnt = len(excelList[0])
        
        for rowI in range(0, len(excelList)) :
            rows = excelList[rowI]

            gui.hotkey('alt', 'f3') # 조회
            time.sleep(0.5)
            gui.hotkey('alt', 'f2') # 신규
            time.sleep(0.5)
            
            for i in range(0, maxColumnCnt):
                data = rows[i]

                if i==0:
                    imgRightClick('결의구분타이틀.png')
                    if data == '수입': imgClick('수입결의서TXT.png')
                    else: imgClick('지출결의서TXT.png')
                    continue
                elif i==2:
                    imgRightClick('사업타이틀.png')
                    customizationProjectImgClick(self)
                    continue
                elif i==3:
                    # 결의일자 활성화
                    imgRightClick('결의일자타이틀.png')
                    # 결의일자 삽입
                    gui.hotkey('ctrl', 'a')
                    gui.press('backspace')
                    pyperclip.copy(data)
                    gui.hotkey('ctrl', 'v')
                    continue
                elif i==4:
                    imgClick('계정코드박스.png')
                    gui.press('tab')
                    pyperclip.copy(data)
                    gui.hotkey('ctrl', 'v')

                    # 본인부담금수입인 경우, 대상자 항목이 오픈됨.
                    if data == '본인부담금수입':
                        imgRightInClick('대상자.png')

                    continue
                elif i==5:
                    imgClick('결의서적요.png')
                    gui.hotkey('ctrl', 'a')
                    gui.press('backspace')
                    pyperclip.copy(data)
                    gui.hotkey('ctrl', 'v')
                    continue
                elif i==6 or i==7:
                    imgRightClick('금액타이틀.png')
                    # 수입이던지, 지출이던지간에 어차피 금액은 다시 세팅될것이므로 그냥 두자...
                    gui.hotkey('ctrl', 'a')
                    gui.press('backspace')
                    pyperclip.copy(data)
                    gui.hotkey('ctrl', 'v')
                    continue
                # elif i==8:
                #     imgClick('자금원천박스.png')
                #     if data == '보조금': imgClick('보조금TXT.png')
                #     elif data == '자부담': imgClick('자부담TXT.png')
                #     elif data == '후원금': imgClick('후원금TXT.png')
                #     elif data == '수익사업': imgClick('수익사업TXT.png')
                #     continue
                elif i==9:
                    if data != '':
                        imgClick('상대계정박스.png')
                        gui.press('tab')
                        pyperclip.copy(data)
                        gui.hotkey('ctrl', 'v')
                        time.sleep(0.2)
                        continue
                    else:
                        continue

            if i == maxColumnCnt-1:
                time.sleep(0.5)
                # 이미지 인식이 잘 되도록 포커스 초기화 하기 위한 클릭 수행
                imgLeftClick('조회.png')
                time.sleep(0.5)
                gui.hotkey('alt', 'f8') # 저장
                time.sleep(0.5)
                # 이미지 인식이 잘 되도록 포커스 초기화 하기 위한 클릭 수행
                screenCenterClick() # imgClick('저장시주의사항아이콘.png')
                time.sleep(0.5)
                imgLeftClick('저장취소.png')
                time.sleep(0.5)
                # 이미지 인식이 잘 되도록 포커스 초기화 하기 위한 클릭 수행
                screenCenterClick() # imgClick('저장시주의사항아이콘.png')
                time.sleep(0.5)
                imgClick('성공저장확인.png')
                time.sleep(0.5)
                #### 성공 확인됨. Flag값 수정하기
                statusChangeTrue(self, rowI)
                time.sleep(0.5)
            
    except Exception as e:
        logging.debug('자동저장(auto_save) Exception : ', e)
# def autoSave End #



#1 이미지 찾아서 가운데 클릭 기능
def imgClick(imgNm):
    img = gui.locateOnScreen(imgDirPath + imgNm)

    if img is not None: 
        center = gui.center(img)
        gui.click(center)
    else:
        gui.alert(f'찾는 이미지 : {imgNm}\n찾고자 하는 이미지가 존재하지 않습니다. \n관리자 확인이 필요합니다.')
        return False
# def imgClick End #


#2 이미지 찾아서 이미지의 오른쪽 위치 클릭 기능
def imgRightClick(imgNm):
    img = gui.locateOnScreen(imgDirPath + imgNm)

    if img is not None:
        moveX = (img.left + img.width) + 10
        moveY = img.top + img.height // 2
        gui.click(moveX, moveY)
    else:
        gui.alert(f'찾는 이미지 : {imgNm}\n찾고자 하는 이미지가 존재하지 않습니다. \n관리자 확인이 필요합니다.')
        return False
# def imgRightClick End #


# 이미지 찾아서 이미지의 왼쪽 위치 클릭 기능
def imgLeftClick(imgNm):
    img = gui.locateOnScreen(imgDirPath + imgNm)

    if img is not None:
        moveX = img.left - 30
        moveY = img.top + img.height // 2
        gui.click(moveX, moveY)
    else:
        gui.alert(f'찾는 이미지 : {imgNm}\n찾고자 하는 이미지가 존재하지 않습니다. \n관리자 확인이 필요합니다.')
        return False


#3 사용자가 올린 사업명 이미지 경로를 찾아서 가운데 클릭
'''
    @param self : PyQt5
'''
def customizationProjectImgClick(self):
    imgPath = self.file_projectImg_path.toPlainText()
    imgNm = self.file_projectImg_nm.toPlainText().split('.')[0]

    clickImg = gui.locateOnScreen(imgPath)

    if clickImg is not None: 
        center = gui.center(clickImg)
        gui.click(center)
    else:
        gui.alert(f'찾는 이미지 : {imgNm}\n찾고자 하는 이미지가 존재하지 않습니다. \n관리자 확인이 필요합니다.')
        return False


#4 사용자가 올린 계좌명 이미지 경로를 찾아서 가운데 클릭
'''
    @param self : PyQt5
'''
def customizationManageImgClick(self):
    findManageImg = gui.locateOnScreen(imgDirPath + '계좌번호(선택).png')
    imgNm = self.file_projectImg_nm.toPlainText().split('.')[0]

    # 만약, 사업이 세팅되어 계좌번호가 자동적으로 세팅되어 있지 않다면 이미지를 찾아서 클릭할 것.
    if findManageImg == None:
        imgPath = self.file_manageImg_path.toPlainText()
        imgRightClick('계좌번호타이틀.png')
        clickImg = gui.locateOnScreen(imgPath)

        if clickImg is not None:
            center = gui.center(clickImg)
            gui.click(center)
        else:
            gui.alert(f'찾는 이미지 : {imgNm}\n찾고자 하는 이미지가 존재하지 않습니다. \n관리자 확인이 필요합니다.')
            
 

#5 화면 정 가운데 클릭
def screenCenterClick():
    # 화면 크기 가져오기
    screen_width, screen_height = gui.size()

    # 정 가운데 좌표 계산
    click_x = screen_width // 2
    click_y = screen_height // 2

    # 클릭 실행
    gui.click(click_x, click_y)


#6 상태 테이블값 true로 변환
def statusChangeTrue(self, rowI): self.status_tb.setItem(rowI, 0, QTableWidgetItem('Success'))


#7 이미지를 찾아서 이미지의 오른쪽 끝을 클릭하는 기능
def imgRightInClick(imgNm):
    img = gui.locateOnScreen(imgDirPath + imgNm)

    if img is not None:
        moveX = img.left + img.width + 10
        moveY = img.top + img.height // 2
        gui.click(moveX, moveY)
    else:
        gui.alert(f'찾는 이미지 : {imgNm}\n찾고자 하는 이미지가 존재하지 않습니다. \n관리자 확인이 필요합니다.')
        return False