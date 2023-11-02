# 상태값 저장
check = False # 체크해야 할 사항들에 대한 체크여부
popup = False # 팝업 체크여부
# finish = False # 프로세스 완전 끝남여부


# 정보 저장
running_stop_flag = None # 프로그램 실행 및 종료 flag (상태값 종류 : running, stop)
excel_obj = None # 엑셀 테이블 저장 장소
process = None # 실행할 때 프로세스의 pid info 변수
