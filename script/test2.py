import win32gui
import win32con

def enum_child_windows(hwnd, lParam):
    child_handles = lParam
    child_handles.append(hwnd)
    return True

def enum_child_controls(hwnd, lParam):
    control_handles = lParam
    control_handles.append(hwnd)
    return True


parent_hwnd = win32gui.FindWindow(None, "XPlatform.exe")                # 부모 윈도우의 핸들 가져오기
child_hwnds = []                                                        # 자식 윈도우 핸들을 저장할 리스트 생성
control_hwnds = []                                                      # 각 자식 윈도우의 컨트롤 핸들을 저장할 리스트 생성
childCyWindowClassList = [] # 자식핸들 중, CyWindowClass 클래스명에 해당하는 자식 윈도우 핸들 리스트 생성

win32gui.EnumChildWindows(parent_hwnd, enum_child_windows, child_hwnds) # 자식 윈도우 핸들 조회


# 각 자식 윈도우의 컨트롤 핸들 조회
for child_hwnd in child_hwnds:
    win32gui.EnumChildWindows(child_hwnd, enum_child_controls, control_hwnds)

# 조회된 컨트롤 핸들 출력
print('------------------------------------------------------')
print('-------------- [자식윈도우 컨트롤 핸들 조회] -----------')
print('------------------------------------------------------')
for control_hwnd in control_hwnds:

    ctrlId = win32gui.GetDlgCtrlID(control_hwnd)
    wndClass = win32gui.GetClassName(control_hwnd)
    wndTxt = win32gui.GetWindowText(control_hwnd)

    print(str(ctrlId) + ', ' + wndClass + ', ' + wndTxt)

    if wndClass == 'ComboBox': 
        childCyWindowClassList.append(control_hwnd)


print('------------------------------------------------------')
print('----------- [자식윈도우 CyWindowClass 조회] -----------')
print('------------------------------------------------------')
for cyWidClas in childCyWindowClassList:
    ctrlId = win32gui.GetDlgCtrlID(control_hwnd)
    wndClass = win32gui.GetClassName(control_hwnd)
    wndTxt = win32gui.GetWindowText(control_hwnd)

    print(str(ctrlId) + ', ' + wndClass + ', ' + wndTxt)