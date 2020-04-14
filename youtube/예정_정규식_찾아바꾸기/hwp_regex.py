"""
아래아한글 내에서 "정규식으로 찾기"는 가능하지만
"정규식으로 찾아바꾸기"는 불가능하다.
미묘한 차이 같지만, 이런 경우를 생각해보자.
주민등록번호나 법인등록번호처럼 "여섯자리숫자-일곱자리숫자(예:321012-1234567) 양식을
321012-1****** 처럼 뒤의 여섯자만 *로 마스킹을 하고 싶을 때,
아래아한글의 찾아바꾸기를 이용하려면 조금 번거롭다..
1. 정규식으로 \d\d\d\d\d\d-\d\d\d\d\d\d\d을 찾고,
({n}은 아래아한글에서 태그지정문법으로 사용중이다..)
2. 해당 위치로 캐럿을 옮긴 다음(찾기 후 Esc)
3. 우측이동을 8번 하고, 오른쪽 여섯글자를 선택한 후에
4. ******을 입력한다.
5. 위의 과정을 반복한다.

파이썬으로는 아래와 같다.

"""


import os
import re

import win32com.client as win32


BASE_DIR = os.getcwd()
file_name = "d6-d7.hwp"
re_pattern = re.compile("(\d{6})[-](\d)\d{6}")  # 주민등록번호

if __name__ == '__main__':
    try:
        hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
        hwp.RegisterModule("FilePathCheckDLL", "SecurityModule")
        hwp.Run("FileNew")
        hwp.Open(os.path.join(BASE_DIR, file_name))
        try:
            hwp.InitScan()
            while True:
                textdata = hwp.GetText()
                if textdata[0] == 1:
                    break
                else:
                    hwp.MovePos(201)
                    text = textdata[1].strip()
                    re_text = re_pattern.sub("\g<1>-\g<2>******", text)
                    if not re_text:
                        pass
                    else:
                        hwp.Run("Select")
                        hwp.Run("Select")
                        hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
                        hwp.HParameterSet.HInsertText.Text = re_pattern.sub("\g<1>-\g<2>******", text)
                        hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)
                        hwp.Run("Cancel")

        finally:
            hwp.ReleaseScan()
            print('Scan Released')
    finally:
        hwp.SaveAs(os.path.join(BASE_DIR, "result.hwp"))
        hwp.Quit()