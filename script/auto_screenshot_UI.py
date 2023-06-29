# 자동화 주제 : 화면 스크린샷 자동화 수행
# - 자동화 수행을 위한 파이썬 UI 생성필요(내장 라이브러리 Tkinter 이용)
######################################################################
from tkinter import *
import pyautogui as gui
import time

# 루트화면 생성
tk = Tk()

# 루트화면 크기 설정
tk.geometry("300x150+0+0")

# 텍스트 표시
label = Label(tk, text="스크린샷 자동화 프로그램")

# 레이블 배치 실행
label.pack()

# 기능 1. 버튼 생성(버튼 클릭 시, 메세지 팝업을 띄우고 'OK'버튼 클릭 시 자동 스크린샷 수행)
def btn_event() :
    # 메세지 팝업 띄우기
    # ok_or_cancle_btn = gui.confirm(
    #     text = "스크린샷 자동화를 수행하시겠습니까?",
    #     title = "스크린샷 자동화 안내창",
    #     buttons = ["OK", "CANCLE"]
    # )

    # 메세지 팝업의 버튼 값에 따라 분기처리하여 수행함.
    # if ok_or_cancle_btn == "OK" :
        screen_show_fn()
    # else :
    #     gui.alert("스크린샷 자동화를 취소합니다.")

    

# 기능 2. 스크린샷 기능
def screen_show_fn() :
    # 마우스를 W4C 프로그램으로 옮긴다.
    # 아이콘 위치 : 스티커메모, 계산기, 스크린샷, 엑셀, 파일탐색기, 슬랙, 카카오톡, 크롬, "W4C"
    
    ##### W4C 프로그램이 띄워져 있다는 전제로 수행한다.

    # 윈도우 스크린샷 프로그램 실행
    gui.hotkey("winleft", "shift", "s")
    time.sleep(0.2)

    # 마우스 위치 이동 2(내가 스크린샷 할 영역 시작)
    gui.moveTo(325, 189, 0.2)
    # 마우스 드래그(내가 스크린샷 할 영역 끝)
    gui.dragTo(1585, 947, 0.2)

    # 한 텀 쉬었다가..
    time.sleep(0.5)

    # 새 알림 클릭
    gui.moveTo(1878, 1053, 0.2)
    gui.click()
    time.sleep(0.5)

    # 가장 위에 있는 클립보드 클릭
    gui.moveTo(1678, 212, 0.2)
    gui.click()
    time.sleep(0.7)
    
    # 캡처 및 스케치 복사
    gui.moveTo(1072, 310, 0.2)
    gui.click()
    print("SUCCESS...")


# 버튼에 함수 삽입 : Button(메인루프 객체, text="내용", command = 실행할 함수)
button = Button(tk, text="스크린샷 시작", command = btn_event)
button["text"] = "PLAY FUNCTION AUTO SCREENSHOT"

# button.pack(side:배치설정, padx:좌우여백설정, pady:상하여백설정)
button.pack(side = LEFT, padx = 10, pady = 10)

# 메인루프 실행
tk.mainloop()