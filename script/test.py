from pywinauto import application
import win32gui
import pywintypes
import sys
import psutil
import win32con


# 프로세스 PID 찾기
proc_pid = None
for proc in psutil.process_iter():
    if proc.name() == 'XPlatform.exe':
        proc_pid = proc.pid
        break

# 프로세스에 연결
app = application.Application().connect(process=proc_pid)

# 메인 윈도우 가져오기
spec = app.window()
print(spec)


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
	teracopyhwnd = WindowFinder(windowname).GetHwnd()
 
	# Teracopy의 모든 child window handle을 검색한다.
	childrenlist = ChildWindowFinder(teracopyhwnd).GetChildrenList()
 
	return teracopyhwnd, childrenlist


# W4C 프로그램 분석
def analysisW4C(child, ctrlId, wndClass, wndTxt):
	print("===============[ W4C 프로그램 ]====================")
	print("%08X %6d\t%s\t%s" % (child, ctrlId, wndClass, wndTxt))
	print("\n")
	print("HWND     CtlrID\t\tClass\t\tWindow Text")

	w4cChildrenList = ChildWindowFinder(child).GetChildrenList()

	for i in range(0, len(w4cChildrenList)):
		w4c_child = w4cChildrenList[i]

		w4c_ctrlId  = win32gui.GetDlgCtrlID(child)
		w4c_wndClass = win32gui.GetClassName(child)
		w4c_wndTxt = win32gui.GetWindowText(child)
		
        


		# # 재귀함수
		# if len(ChildWindowFinder(w4c_child).GetChildrenList()) != None: 
		# 	analysisW4C(w4c_child, w4c_ctrlId, w4c_wndClass, w4c_wndTxt)
		# else:
		# 	continue




def findControls(topHwnd,
                 wantedText=None,
                 wantedClass=None,
                 selectionFunction=None):
    
    def searchChildWindows(currentHwnd):
        results = []
        childWindows = []
        try:
            win32gui.EnumChildWindows(currentHwnd,
                                    #   _windowEnumerationHandler,
                                      None,
                                      childWindows)
        except win32gui.error:
            # This seems to mean that the control *cannot* have child windows,
            # i.e. not a container.
            return
        for childHwnd, windowText, windowClass in childWindows:
	    
            print("%08X \t%s\t%s" % (childHwnd, windowText, windowClass))

            descendentMatchingHwnds = searchChildWindows(childHwnd)
            if descendentMatchingHwnds:
                results += descendentMatchingHwnds

            # if wantedText and \
            #    not _normaliseText(wantedText) in _normaliseText(windowText):
            #     continue
            if wantedClass and \
               not windowClass == wantedClass:
                continue
            if selectionFunction and \
               not selectionFunction(childHwnd):
                continue
            results.append(childHwnd)
        return results

    return searchChildWindows(topHwnd) 





# main 입니다.
def main(argv):
	hwnd, childwnds = GetChildWindows('XPlatform.exe')
	print("%X %s" % (hwnd, win32gui.GetWindowText(hwnd)))
	print("HWND     CtlrID\tClass\tWindow Text")
	print("===========================================")
 
	# 찾고자 하는 객체명
	obj = {}
	
	for child in childwnds:
		ctrlId  = win32gui.GetDlgCtrlID(child)
		wndClass = win32gui.GetClassName(child)
		wndTxt = win32gui.GetWindowText(child)

		if '사회복지시설정보시스템(1W)' in wndTxt:
			print("%08X %6d\t%s\t%s" % (child, ctrlId, wndClass, wndTxt))
		
			obj['child'] = child
			obj['ctrlId'] = ctrlId
			obj['wndClass'] = wndClass
			obj['wndTxt'] = wndTxt
            
			findControls(child)

			break

	
	# analysisW4C(obj['child'], obj["ctrlId"], obj['wndClass'], obj['wndTxt'])

	return 0
 
if __name__ == '__main__':
	sys.exit(main(sys.argv))