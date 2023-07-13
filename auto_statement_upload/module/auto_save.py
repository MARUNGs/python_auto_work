# 결의서/전표 자동저장 작업
import os
import pyautogui as gui # 운영체제 제어
import pyperclip # 데이터 복사 및 붙여넣기 
import time
import logging # 로그
from enum import Enum


# 공통 경로
imgDirPath = os.path.dirname(__file__).replace('module', 'img') + os.sep


# enum
class ProjectType(Enum):
    tp01 = '일반사업'
    tp02 = '보조금사업'
    tp03 = '후원금사업'
    tp04 = '특별회계사업'
    tp05 = '복지수당사업'
    tp06 = '기능보강사업'
    tp07 = '종사자처우개선비사업'
    tp08 = '주야간보호'
    tp09 = '방문요양'
    tp10 = '방문목욕'
    tp11 = '단기보호'
    tp12 = '방문간호'
    tp13 = '노인요양시설(개정법)'
    tp14 = '노인요양공동생활가정'
    tp15 = '복지용구제공사업소'

# 기본 딜레이 설정
gui.PAUSE = 0.2
gui.FAILSAFE = False


'''
    @param self      # PyQT5
    @param excelList # makeExcelData()를 통해 갖고있는 데이터
'''
def autoSave(self, excelList):
    try:
        maxColumnCnt = len(excelList[0])

        imgClick('신규.png')
        # time.sleep(0.2)
        
        
        for rows in excelList:
            for i in range(0, maxColumnCnt):
                data = rows[i]

                if i==0:
                    imgRightClick('결의구분타이틀.png')
                    # time.sleep(0.2)

                    if data == '수입': imgClick('수입결의서TXT.png')
                    else: imgClick('지출결의서TXT.png')

                    # time.sleep(0.2)
                    continue
                elif i==2:
                    imgRightClick('사업타이틀.png')
                    # time.sleep(0.2)

                    customizationImgClick(self)
                    # time.sleep(0.2)
                    
                    continue
                elif i==3:
                    # 결의일자 활성화
                    imgRightClick('결의일자타이틀.png')
                    # time.sleep(0.2)

                    # 결의일자 삽입
                    gui.hotkey('ctrl', 'a')
                    gui.press('backspace')
                    pyperclip.copy(data)
                    gui.hotkey('ctrl', 'v')
                    # time.sleep(0.2)
                    continue
                elif i==4:
                    imgClick('계정코드박스.png')
                    # time.sleep(0.2)

                    gui.press('tab')
                    pyperclip.copy(data)
                    gui.hotkey('ctrl', 'v')
                    # time.sleep(0.2)
                    continue
                elif i==5:
                    imgClick('결의서적요.png')
                    # time.sleep(0.2)

                    gui.hotkey('ctrl', 'a')
                    gui.press('backspace')
                    pyperclip.copy(data)
                    gui.hotkey('ctrl', 'v')
                    # time.sleep(0.2)
                    continue
                elif i==6 or i==7:
                    imgRightClick('금액타이틀.png')
                    # time.sleep(0.2)
                    
                    # 수입이던지, 지출이던지간에 어차피 금액은 다시 세팅될것이므로 그냥 두자...
                    gui.hotkey('ctrl', 'a')
                    gui.press('backspace')
                    pyperclip.copy(data)
                    gui.hotkey('ctrl', 'v')
                    # time.sleep(0.2)
                    continue
                elif i==8:
                    imgClick('자금원천박스.png')
                    # time.sleep(0.2)

                    if data == '보조금': imgClick('보조금TXT.png')
                    elif data == '자부담': imgClick('자부담TXT.png')
                    elif data == '후원금': imgClick('후원금TXT.png')
                    elif data == '수익사업': imgClick('수익사업TXT.png')
                    # time.sleep(0.2)
                    continue
                elif i==9:
                    if data != '':
                        imgClick('상대계정박스.png')
                        # time.sleep(0.2)

                        gui.press('tab')
                        pyperclip.copy(data)
                        gui.hotkey('ctrl', 'v')
                        # time.sleep(0.2)
                        continue
                    else:
                        continue
                elif i==11:
                    # 계좌명의 경우, 사업을 설정하면 자동적으로 매핑되는데 매핑되었는지 안 되었는지를 확인하여 처리하면 될 듯.
                    print('아직 안했어')
    except Exception as e:
        logging.debug('autoSave Exception : ', e)
# def autoSave End #



# 이미지 찾아서 가운데 클릭 기능
def imgClick(imgNm):
    img = gui.locateOnScreen(imgDirPath + imgNm)

    if img is not None: 
        center = gui.center(img)
        gui.click(center)
    else:
        gui.alert('찾고자 하는 이미지가 존재하지 않습니다. \n관리자 확인이 필요합니다.')
        return False
# def imgClick End #


# 이미지 찾아서 이미지의 오른쪽 위치 클릭 기능
def imgRightClick(imgNm):
    img = gui.locateOnScreen(imgDirPath + imgNm)

    if img is not None:
        img_right = img.left + img.width
        moveX = img_right + 10
        moveY = img.top + img.height // 2
        gui.click(moveX, moveY)
    else:
        gui.alert('찾고자 하는 이미지가 존재하지 않습니다. \n관리자 확인이 필요합니다.')
        return False
# def imgRightClick End #



# 사용자가 올린 이미지 경로를 찾아서 가운데 클릭
'''
    @param self : PyQt5
'''
def customizationImgClick(self):
    imgPath = self.file_projectImg_path.toPlainText()
    imgNm = self.file_projectImg_nm.toPlainText().split('.')[0]

    #일반사업
    if imgNm == ProjectType.tp01.value: 
        clickImg = gui.locateOnScreen(imgPath)

        if clickImg is not None: 
            center = gui.center(clickImg)
            gui.click(center)
        else:
            gui.alert('찾고자 하는 이미지가 존재하지 않습니다. \n관리자 확인이 필요합니다.')
            return False
    elif imgNm == ProjectType.tp02:
        imgClick()
# def customizationImgClick End #