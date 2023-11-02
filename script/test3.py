import pyautogui
# import psutil
# import sys
# from PyQt5.QtWidgets import QApplication, QWidget, QPushButton

w4c_window = pyautogui.getWindowsWithTitle('사회복지시설정보시스템(1W)')[0]
if w4c_window.isActive == False: w4c_window.activate()
for _ in range(0, 1000): pyautogui.press('enter')







# first_process = None

# while True :
#     for p in psutil.process_iter(attrs=['pid', 'name', 'status', 'cmdline']) :
#         if p.info['name'].lower() == 'python.exe' :
#             first_process = p.info
#             break
#         else :
#             continue
#     break

# print('--------------- process status값 확인 (1) 시작 ------------------')
# print(first_process)
# print('--------------- process status값 확인 (1) 끝 ------------------')

# # PyQt5 애플리케이션 생성
# app = QApplication(sys.argv)

# # 메인 창 생성
# window = QWidget()
# window.setWindowTitle("간단한 PyQt5 창 예제")
# window.setGeometry(100, 100, 400, 200)  # 창 위치와 크기 설정


# # 버튼 생성
# button = QPushButton('클릭하세요', window)
# button.setGeometry(150, 70, 100, 30)  # 버튼 위치와 크기 설정

# # 버튼 클릭 시 실행할 함수
# def on_button_click():
#     for _ in range(0, 10000000) :
#         now_process = psutil.Process(first_process['pid'])
#         # print(now_process.pid)
#         # print(now_process.status)
#         # pyautogui.press('enter')

#         process_status = now_process.status()
#         print(f"프로세스 상태: {process_status}")


#         import ctypes

#         # Windows API 함수 및 데이터 타입 가져오기
#         user32 = ctypes.windll.user32
#         cursor_id = user32.GetSystemMetrics(106)  # 106은 모래시계 모양의 ID입니다.

#         if cursor_id == 106:
#             cursor_info = "모래시계 모양"
#         else:
#             cursor_info = "다른 모양"

#         print(f"현재 마우스 커서 모양: {cursor_info}")



        
#         if process_status == psutil.STATUS_STOPPED:
#             print("프로세스가 응답하지 않습니다.")
#         pyautogui.press('enter')


# # 버튼 클릭 이벤트 연결
# button.clicked.connect(on_button_click)


# # 창을 화면에 표시
# window.show()

# # PyQt5 애플리케이션 루프 시작
# sys.exit(app.exec_())

























# from pynput.keyboard import Key, Listener, KeyCode


# def on_key_release(key):
#     # 키 릴리스 이벤트 처리
#     print(f'Key released: {KeyCode.from_char(key)}')

# def on_key_press(key):
#     # 키 프레스 이벤트 처리
#     print(f'Key pressed: {KeyCode.from_char(key)}')

# # 키보드 리스너 생성
# with Listener(on_press=on_key_press, on_release=on_key_release) as listener:
#     listener.join()

