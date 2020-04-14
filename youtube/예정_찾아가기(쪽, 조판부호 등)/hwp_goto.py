import win32com.client as win32
hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")

hwp.HAction.GetDefault("Goto", hwp.HParameterSet.HGotoE.HSet)
hwp.HParameterSet.HGotoE.HSet.SetItem("DialogResult", 2) # 2쪽
hwp.HParameterSet.HGotoE.SetSelectionIndex = 1 # "쪽" 찾아가기
hwp.HAction.Execute("Goto", hwp.HParameterSet.HGotoE.HSet)
hwp.Run("Cancel")