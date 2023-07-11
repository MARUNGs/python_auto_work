from pywinauto import application
import psutil
import win32gui
import pywintypes
import sys



# 부모 윈도우의 핸들을 검사합니다.
class WindowFinder:
	def __init__(self, windowname):
		try:
			win32gui.EnumWindows(self.__EnumWindowsHandler, windowname)		
		except pywintypes.error as e:
			# 발생된 예외 중 e[0]가 0이면 callback이 멈춘 정상 케이스
			if e[0] == 0: pass		
 
	def __EnumWindowsHandler(self, hwnd, extra):
		wintext = win32gui.GetWindowText(hwnd)
		if wintext.find(extra) != -1:
			self.__hwnd = hwnd
			return pywintypes.FALSE # FALSE는 예외를 발생시킵니다.	
 
	def GetHwnd(self):
		return self.__hwnd
 
	__hwnd = 0



# 자식 윈도우의 핸들 리스트를 검사합니다.
class ChildWindowFinder:
	def __init__(self, parentwnd):
		try:
			win32gui.EnumChildWindows(parentwnd, self.__EnumChildWindowsHandler, None)		
		except pywintypes.error as e:
			if e[0] == 0: pass		
 
	def __EnumChildWindowsHandler(self, hwnd, extra):
		self.__childwnds.append(hwnd)
 
	def GetChildrenList(self):
		return self.__childwnds
 
	__childwnds = []
	

# windowname을 가진 윈도우의 모든 자식 윈도우 리스트를 얻어낸다.
def GetChildWindows(windowname):
 
	# TeraCopy의 window handle을 검사한다.
	teracopyhwnd = WindowFinder('TeraCopy').GetHwnd()
 
	# Teracopy의 모든 child window handle을 검색한다.
	childrenlist = ChildWindowFinder(teracopyhwnd).GetChildrenList()
 
	return teracopyhwnd, childrenlist
 
# main 입니다. -  현재 실행중인 프로세스 리스트 확인 가능
def main(argv):
	hwnd, childwnds = GetChildWindows('teraCopy')
	CyWindowClassList = []
    
	print("%X %s" % (hwnd, win32gui.GetWindowText(hwnd)))
	print("HWND     CtlrID\tClass\tWindow Text")
	print("===========================================")
    
	for child in childwnds:
		ctrl_id = win32gui.GetDlgCtrlID(child)
		wnd_clas = win32gui.GetClassName(child)
		wnd_txt = win32gui.GetWindowText(child)
		print('CyWindowClass : ' + "%08X %6d\t%s\t%s" % (child, ctrl_id, wnd_clas, wnd_txt))
		
        # CyWindowClass만 담기
		if wnd_clas == 'CyWindowClass':
			CyWindowClassList.append({
				'child': child,
				'ctrl_id': ctrl_id,
				'wnd_clas': wnd_clas,
				'wnd_txt': wnd_txt
            })
			
	print("===========================================")
	
	
	# list 중에서 사회복지시설정보시스템만 별도로 뺀 값 삽입
	w4cInfo = filterW4CInfo(CyWindowClassList)


	return 0

# def main End #



# list 중에서 사회복지시설정보시스템만 별도로 뺀 값 삽입
def filterW4CInfo(CyWindowClassList):
	pickWindow = None
	for item in CyWindowClassList:
		if item['wnd_txt'] == '사회복지시설정보시스템(1W)':
			pickWindow = item
			break

	return pickWindow
# def filterW4CInfo End #













if __name__ == '__main__':
	sys.exit(main(sys.argv))




def EnumWindowsHandler(hwnd, extra):
    wintxt = win32gui.GetWindowText(hwnd)
    print("%08X: %s" % (hwnd, wintxt))

win32gui.EnumWindows(EnumWindowsHandler, None)









































# 작업할 프로세스 정보 확인
# procPid = None
# for proc in psutil.process_iter():
#     if proc.name() == 'XPlatform.exe':
#         print(proc.name() + ' : ' + str(proc.pid))
#         procPid = proc.pid
#         break

# 이미 실행된 작업할 프로세스 연결
# app = application.Application(backend='win32')
# app.connect(process=procPid)


# # 컨트롤러 출력
# dlg = app['사회복지시설정보시스템(1W)']
# dlg.print_control_identifiers() # 컨트롤러 리스트 트리형태로 모두 출력



