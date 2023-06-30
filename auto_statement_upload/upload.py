# 주제: 마음손 전표정보 업로드 자동화

# [ 작업순서 ]
# 1. 업로드할 파일명은 미리 복사하여 입력창에 작성해두어야 한다.
# 2. 입력한 파일명을 기준으로 파일을 검색한다.
# 3. 검색된 파일명이 맞으면 enter 키 입력을 수행한다.
# 4. 파일명에 존재하는 확장자를 추출하여 xlsx인지 xls인지 확인한다.
# 5. xls인 경우, 호환성 여부 안내창을 인식한다.



########## import list ##############################################################################
import os               # 운영체제 정보
from tkinter import *   # 파이썬 UI 구현
import pyautogui as gui # 마우스 & 키보드 제어




########## 참고한 블로그 ###############################################################################
# Tkinter 위젯 배치 :  https://camplee.tistory.com/32
# Tkinter 위젯 x,y 배치 : https://cosmosproject.tistory.com/610
# Tkinter 여러가지 설정(읽어보기) : https://m.blog.naver.com/PostView.naver?isHttpsRedirect=true&blogId=yheekeun&logNo=220701094001
# 파일이 열려있는가 확인 : https://cyworld.tistory.com/3053
# 


########## Setting window(window) base ################################################################
 # 루트화면 생성
window = Tk()

# 창 이름 설정
window.title('SmartBiz(마음손) - 전표등록 자동업로드 작업 프로그램')
# 루트화면 크기, 위치 설정
window.geometry('600x400+0+0')
# 루트화면 크기 조절 불가
window.resizable(False, False)




########## Event function ##############################################################################
'''
    #1 open_file_fn() : 파일열기 버튼 이벤트
        처리 1. 파일명이 .xlsx 또는 .xls 문자열이 포함하면 start_auto_fn() 이벤트를 수행한다.
        처리 2. 위의 문자열이 포함하지 않으면 이벤트를 종료한다.
    #2 start_auto_fn() : 전표정보 자동 업로드 시작 이벤트
'''
# 1
def open_file_fn() :
    # 1. 파일명의 확장자가 .xlsx 또는 .xls가 맞는지 확인 
    print(upload_file_nm)
    
    if upload_file_nm is None :
        gui.alert('작업할 파일명을 먼저 입력하세요.')
        return False
    else :
        print(upload_file_nm in '.xlsx') or (upload_file_nm in '.xls')

        if (upload_file_nm in '.xlsx') or (upload_file_nm in '.xls'):
            gui.alert('upload_file_fn success')

            # 확인 차 alert창으로 프로세스를 잠시 멈춤
            gui.alert("전표등록 자동업로드 작업을 시작합니다.")
            start_auto_fn()
        else : 
            gui.alert("전표등록 자동업로드 작업을 취소합니다.")

# 2
def start_auto_fn() :
    gui.alert('start_auto_fn success')


########## Init UI Form ##############################################################################
# 루트화면에 보여질 정적 요소들
Label(window, text='-------------------- 프로그램 사용 안내글 --------------------').place(x=5, y=20)
Label(window, text='0. 해당 프로그램은 마음손 사용자 외에는 사용 불가합니다.').place(x=5, y=40)
Label(window, text='1. 자동 업로드 처리할 첨부파일명을 입력하세요.').place(x=5, y=60)
# Label(window).grid(row=1, column=0)
# upload_file_nm_label = Label(window, text='첨부파일명').grid(row=2, column=0)

# 업로드파일명 입력창
# upload_file_nm = Entry(window).grid(row=2, column=1)

# 버튼1. 파일열기
# open_file_btn = Button(window, text='업로드파일 자동탐색 후 열기', command=open_file_fn).grid(row=3, column=0)
# 버튼2. 전표정보 자동 업로드 시작
# start_auto_btn = Button(window, text='마음손 전표내역 업로드 시작', command=start_auto_fn).grid(row=3, column=1)





window.mainloop()