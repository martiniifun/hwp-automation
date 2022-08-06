import win32com.client as win32
hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
hwp.XHwpWindows.Item(0).Visible = True


def goto_page(page):
    hwp.HAction.GetDefault("Goto", hwp.HParameterSet.HGotoE.HSet)
    hwp.HParameterSet.HGotoE.SetSelectionIndex = 1
    hwp.HParameterSet.HGotoE.HSet.SetItem("DialogResult", page)
    hwp.HAction.Execute("Goto", hwp.HParameterSet.HGotoE.HSet)

goto_page(3)