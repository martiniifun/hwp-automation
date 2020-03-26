"""
상장을 대량으로 생성하는 코드입니다.

<대략의 작업순서>
1. 엑셀파일(수상자명단.xlsx)을 pandas로 열어서 데이터프레임으로 변환한 후
2. 한/글파일(award.hwp)을 열어서 상장 개수만큼 페이지를 복사하고
3. 각 페이지의 누름틀(필드)에 해당하는 값을 데이터프레임에서 불러와서 차례대로 삽입.

<참고사항>
작업이 진행되는 과정을 보여주기 위해,
한/글 창을 백그라운드로 숨기는 코드와
작업완료된 한/글 파일을 저장하고 한/글 창을 닫는 등의 코드는
모두 주석으로 처리해 두었습니다.

pandas, pywin32 모듈을 사용하기 위해 pip를 통해 설치해야 함
커맨드창에서, "pip install pandas pywin32" 입력 후 엔터


한/글파일 출처 : 네이버 한글한글아름답게(https://hangeul.naver.com)
"""

import shutil  # 파일복사용 모듈
import win32com.client as win32  # 한/글 열기 위한 모듈
import pandas as pd  # 그 유명한 판다스. 엑셀파일을 다루기 위함
from datetime import datetime as dt  # 작업시간을 측정하기 위함. 지워도 됨.
# import win32gui  # 한/글 창을 백그라운드로 숨기기 위한 모듈


excel = pd.read_excel(r"C:\Users\User\Desktop\수상자명단.xlsx")  # 엑셀로 데이터프레임 생성
hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")  # 한/글 열기
# hwnd = win32gui.FindWindow(None, '빈 문서 1 - 한글')  # 한/글 창의 윈도우핸들값을 알아내서
# win32gui.ShowWindow(hwnd, 0)  # 한/글 창을 백그라운드로 숨김
hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")  # 보안모듈 적용(파일 열고닫을 때 팝업이 안나타남)

shutil.copyfile(r"C:\Users\User\Desktop\award(field).hwp",  # 원본은 그대로 두고,
                r"C:\Users\User\Desktop\award_result.hwp")  # 복사한 파일을 수정하려고 함.

hwp.Open(r"C:\Users\User\Desktop\award_result.hwp")  # 수정할 한/글 파일 열기

start_time = dt.now()  # 작업시간을 측정하기 위해 현재 시각을 start_time 변수에 저장. 없어도 됨...
field_list = [i for i in hwp.GetFieldList().split("\x02")]  # 한/글 안의 누름틀 목록 불러오기
hwp.Run('SelectAll')  # Ctrl-A (전체선택)
hwp.Run('Copy')  # Ctrl-C (복사)
hwp.MovePos(3)  # 문서 끝으로 이동

print('페이지 복사를 시작합니다.')

for i in range(len(excel) - 1):  # 엑셀파일 행갯수-1 만큼 한/글 페이지를 복사(기존에 한쪽이 있으니까)
    hwp.Run('Paste')  # Ctrl-V (붙여넣기)
    hwp.MovePos(3)  # 문서 끝으로 이동

print(f'{len(excel)}페이지 복사를 완료하였습니다.')

for page in range(len(excel)):  # 한/글 모든 페이지를 전부 순회하면서,
    for field in field_list:  # 모든 누름틀에 각각,
        hwp.MoveToField(f'{field}{{{{{page}}}}}')  # 커서를 해당 누름틀로 이동(작성과정을 지켜보기 위함. 없어도 무관)
        hwp.PutFieldText(f'{field}{{{{{page}}}}}',  # f"{{{{{page}}}}}"는 "{{1}}"로 입력된다. {를 출력하려면 {{를 입력.
                         excel[field].iloc[page])  # hwp.PutFieldText("index{{1}}") 식으로 실행될 것.
    print(f'{page + 1}:{excel.name[page]}')  # 현재 입력이 진행되고 있는 한/글문서 페이지번호를 콘솔창에 출력
# hwp.Save()  # 한/글 파일(award_result.hwp)을 저장하고,
# hwp.Quit()  # 한/글 종료. (저장하지 않고 종료하는 방법은 7강에서~)

end_time = dt.now()  # 작업종료 시각. 없어도 무관.
소요시간 = end_time - start_time  # 전체 작업시간을 기록. 없어도 무관.

print(f'작업을 완료하였습니다. 약 {소요시간.seconds}초 소요되었습니다.')  # 작업완료된 후 출력. 끝.