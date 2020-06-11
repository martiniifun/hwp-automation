"""
30칸 좌로 가면 100%
그대로 있으면 50%
30칸 우로 가면 0%
100%를 60칸으로 나눠놨다.
한 칸 이동에 1.7%
예를 들어 14%라면?
50-14 = 36% 우측으로 이동해야 하므로,
36/1.7 = 21! 우측으로 21칸 이동하면 됨

몇칸 이동하는지 계산하려면?
x -= 50
x = int(x/1.7)
for i in range(x):
    hwp.HAction.Run("TableCellBlock");
    if x<0:
	    hwp.HAction.Run("TableResizeCellRight")
	else:
	    hwp.HAction.Run("TableResizeCellLeft")

"""

from time import sleep

import win32com.client as win32
import pandas as pd

def shift(percent, direction):
    percent -= 50
    percent = int(percent / 1.7)
    hwp.HAction.Run("TableCellBlock")
    for i in range(abs(percent)):
        sleep(0.01)
        if direction == "left":
            if percent < 0:
                hwp.HAction.Run("TableResizeCellRight")
            else:
                hwp.HAction.Run("TableResizeCellLeft")
        elif direction == "right":
            if percent < 0:
                hwp.HAction.Run("TableResizeCellLeft")
            else:
                hwp.HAction.Run("TableResizeCellRight")
        else:
            pass

hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
hwp.RegisterModule("FilePathCheckDLL", "SecurityModule")
hwp.Run("FileNew")
hwp.Open(r"C:\Users\smj02\Desktop\학습양식의_활용_샘플.hwp")
df = pd.read_excel(r"C:\Users\smj02\Desktop\학습양식의_활용_자료_샘플.xlsx")

sleep(3)
field_list = [i for i in hwp.GetFieldList().split('\x02') if not i.isdigit()]

for i in field_list:
    if i.endswith('-'):
        hwp.PutFieldText(i, str(int(df[df["구분"] == i[:-1]]["사용안함/거의안함(%)"])) + "%")
    else:
        hwp.PutFieldText(i, str(int(df[df["구분"] == i[:-1]]["가끔사용/자주사용(%)"])) + "%")

for i in range(2*len(df)):
    if i < len(df):
        hwp.MoveToField(str(i))
        shift(df.iloc[i]["사용안함/거의안함(%)"], "left")
    else:
        hwp.MoveToField(str(i))
        shift(df.iloc[i-len(df)]["가끔사용/자주사용(%)"], "right")

print("완료되었습니다.")