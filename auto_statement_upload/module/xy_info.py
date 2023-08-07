# 이미지 좌표 저장(속도개선)
def xy_info_map():
    return {
        'init_point': (448,139), # 상대좌표(0, 0)
        'init_resolution': (1920,1080), # 기본 해상도 정보
        # 메뉴이동 이미지 좌표 #
        'move_menu1_depth1': (780,174),
        'move_menu1_depth2': (810,204),
        'move_menu1_depth3': (550,402),
        'move_menu2_depth1': (1189,176),
        'move_menu2_depth2': (535,402),
        # 일반전표 이미지 좌표 #
        'search': (1434,310),                          # 조회
        'income_expense_type': (1253,375),             # 결의구분
        'income': (1176,424),                          # 수입결의서
        'expense': (1210,442),                         # 지출결의서
        'project_combobox': (946,398),                 # 사업 combobox
        'cashier_dt': (672,377),                       # 결의일자
        'account_subject_cd_box': (692,426),           # 계정코드 박스
        'subject_box': (1447,422),                     # 대상자 박스
        'popup_account_list': (696,246),               # 팝업_계정코드목록
        'popup_account_select': (932,781),             # 팝업_계정코드_선택
        'popup_account_cd': (1035,284),                # 계정코드 코드명 입력
        'summary': (1238,472),                         # 결의서적요
        'amount': (1191,400),                          # 금액
        'opponent_account_subject_cd_box': (1186,422), # 상대계정과목 박스
        'popup_opponent_account_subject_cd': (796,273),# 상대계정_코드명_입력
        # 인건비(지출) - 급여대장 이미지 좌표 1. 급여대장 화면 #
        'payroll_project_combobox': (1284,836),        # 급여대장 사업명 combobox
        'payroll_project_refresh': (1254,862),         # 선택하세요 항목
        'payroll_add_row_btn': (1315,436),             # 행추가
        'payroll_staff_select_btn': (933,715),         # 직원선택
        'payroll_all_check': (483,480),                # 전체선택
        'payroll_expense_registration': (1386,837),    # 지출결의서 등록
        'payroll_account_year_select_btn': (925,623),  # 회계연도 선택
        # 인건비(지출) - 급여대장 이미지 좌표 2. 결의서/전표 등록/수정 화면 #
        'payroll_cashier_dt': (670,326),               # 결의일자
        'payroll_statement_add_row_btn': (609,495),    # 행추가
        'account_magnifier_icon': (723,542),           # 인건비(지출) 계정과목 아이콘
        'payroll_account_subject_nm': (685,545),       # 계정과목
        # ※ 팝업 계정과목 코드명입력은 위의 작성된 팝업_계정코드목록과 동일해서 같이 이용해도 됨
        # 저장 시 안내창 공통좌표
        'payroll_save': (1095,240),                    # 인건비(지출) 저장
        'save_ok': (928,567),                          # 저장(확인)
        'save_msg_ok': (961,569),                      # 저장완료메세지(확인)
        'close': (1434,239)                            # 닫기
    }