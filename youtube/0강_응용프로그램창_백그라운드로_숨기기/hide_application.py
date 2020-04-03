"""
안녕하세요?
파이썬으로 엑셀이나 아래한글 등 응용프로그램 자동화를 해보시면, 백그라운드로 숨기고 싶거나,
화면에 나타나게 하고 싶은 경우가 있죠. 셀레늄 크롤링할 때 크롬의 headless옵션처럼요.
MS제품군, 엑셀 같은 경우는 excel.Visible 변수(bool)를 False로 정의해서 프로그램화면을
백그라운드로 숨기거나 True로 정의해서 나타나게 할 수 있는데, 아래한글 등 관련 메서드가 따로 없는
프로그램의 경우에는 동영상처럼 해주시면 됩니다.
동영상을 요약하면 win32gui.FindWindow(args)로 응용프로그램의 핸들값(int)을 찾고,
win32gui.ShowWindow(args)로 창을 숨기거나 나타나게 합니다.
백그라운드 작업 종료시에는 try-finally 구문 등으로 프로그램을 닫는 코드를 꼭 넣어주셔서
메모리누수가 일어나지 않도록 주의해 주시기 바랍니다.
문의사항은 댓글로 남겨주시면 설명드리겠습니다. 감사합니다.

!pip install pywin32로 관련모듈 설치 필요
"""


import win32com.client as win32
import win32gui
import win32con


# 엑셀 실행하기
excel = win32.Dispatch("Excel.Application")
print(excel.Visible)


# 엑셀 숨기기 해제
excel.Visible = True

# 엑셀 종료
excel.Quit()

# 아래아한글 실행하기
hwp = win32.Dispatch("HWPFrame.HwpObject")

# 아래아한글의 핸들값 찾기
hwnd = win32gui.FindWindow(None, "빈 문서 1 - 한글")
print(hwnd)

# 아래아한글 백그라운드로 숨기기
win32gui.ShowWindow(hwnd, win32con.SW_HIDE)

# 아래아한글이 실행중인지 확인하기
hwp.InitScan()
hwp.GetText()
hwp.GetText()
hwp.GetText()
hwp.GetText()

# 아래아한글 숨기기 해제
win32gui.ShowWindow(hwnd, win32con.SW_SHOW)

# 아래아한글 종료
hwp.Quit()
