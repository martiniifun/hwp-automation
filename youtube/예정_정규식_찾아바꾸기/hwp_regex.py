"""
아래아한글 내에서 "정규식으로 찾기"는 가능하지만
"정규식으로 찾아바꾸기"는 불가능하다.
미묘한 차이 같지만, 이런 경우를 생각해보자.
주민등록번호나 법인등록번호처럼 "여섯자리숫자-일곱자리숫자(예:321012-1234567) 양식을
321012-1****** 처럼 뒤의 여섯자만 *로 마스킹을 하고 싶을 때,
아래아한글의 찾아바꾸기를 이용하려면 조금 번거롭다..
1. 정규식으로 \d\d\d\d\d\d-\d\d\d\d\d\d\d을 찾고,
(일반적으로 반복횟수에 해당하는 {n}은 아래아한글에서 태그지정문법으로 사용중이어서, 저렇게 쓸 수밖에 없다..)
2. 해당 위치로 캐럿을 옮긴 다음(찾기 후 Esc)
3. 우측이동을 8번 하고, 오른쪽 여섯글자를 선택한 후에
4. ******을 입력한다.
5. 위의 과정을 반복한다.(상당히 복잡하다.)

파이썬으로는 아래와 같다.
"""

import os  # 경로 선택 등을 위한 모듈
import re  # 정규표현식 모듈
import tkinter
from tkinter.filedialog import askopenfilename

import win32com.client as win32  # 아래아한글 제어를 위한 모듈

BASE_DIR = os.getcwd()  # 아래아한글은 저장/불러오기 할 때 전체경로를 입력해야 한다. 그러기 위해서는...
# file_name = "d6-d7.hwp"  # 정규식으로 찾아바꾸기 할 문서 -> tkinter로 파일다이얼로그 대체
tkinter.Tk().withdraw()
file_name = askopenfilename(initialdir=BASE_DIR)
re_pattern = re.compile("(\d{6})[-](\d)\d{6}")  # 주민등록번호 패턴. 뒷자리 첫 번째 숫자를 남겨놓기 위해 그룹을 두 개.

if __name__ == '__main__':  # main함수 실행(본 파일을 다른 py에서 import한 경우가 아니라면 아래 코드를 실행
    try:  # 한글이 오류로 닫힐 때에 finally 구문으로 백그라운드의 모든 hwp 인스턴스를 닫기 위함
        hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")  # 한/글 인스턴스 생성
        hwp.RegisterModule("FilePathCheckDLL", "SecurityModule")  # 보안모듈
        hwp.Run("FileNew")  # 화면에 보이는 창 하나 띄우기(백그라운드작업을 원하면 이 라인을 지워도 됨)
        hwp.Open(os.path.join(BASE_DIR, file_name))  # 파일 열기
        try:  # GetText() 과정 중 오류 발생시에도 ReleaseScan()을 실행하기 위함
            hwp.InitScan()  # GetText 메서드를 실행하기 위한 검색초기화 메서드
            while True:
                textdata = hwp.GetText()  # 문자열을 탐색하면서 튜플을 반환함(엔터로 구분) [0]:상태코드, [1]:텍스트
                if textdata[0] == 1:  # 상태코드가 "검색종료"면:
                    break  # while문 종료
                else:  # 상태코드가 "검색종료"가 아닌 동안은:
                    hwp.MovePos(201)  # GetText() 메서드의 결괏값(텍스트) 위치로 이동
                    text = textdata[1].strip()  # 텍스트 좌우의 이스케이프문자열(또는 스페이스) 제거
                    re_text = re_pattern.sub("\g<1>-\g<2>******", text)  # 정규식 변환후 텍스트 생성
                    if not re_text:  # (매칭되는 문자열이 없어) 빈 문자열인 경우에는:
                        pass  # 아무 작업도 하지 말고 패스
                    else:  # 매칭되는 결과(주민등록번호)가 있는 경우에는:
                        hwp.Run("Select")  # Run("Select") 한 번 실행하면 "선택모드 활성화"
                        hwp.Run("Select")  # Run("Select") 두 번 실행하면 "스페이스로 구분된 한 단어" 선택(덮어쓰기 위함)
                        hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)  # 텍스트 삽입메서드 초기화
                        hwp.HParameterSet.HInsertText.Text = re_text  # 삽입할 텍스트(re_text) 입력
                        hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)  # 텍스트 삽입메서드 실행
                        hwp.Run("Cancel")  # 선택모드 종료
        finally:
            hwp.ReleaseScan()  # InitScan() 및 GetText() 후에는 ReleaseScan() 실행-검색종료
            print('정규식 바꾸기 작업을 완료하고 "./result.hwp"로 저장하였습니다.')
    finally:
        hwp.SaveAs(os.path.join(BASE_DIR, "result.hwp"))  # BASE_DIR에 "result.hwp"라는 이름으로 저장하고
        hwp.Quit()  # 한/글 인스턴스 종료(백그라운드 인스턴스 포함)
