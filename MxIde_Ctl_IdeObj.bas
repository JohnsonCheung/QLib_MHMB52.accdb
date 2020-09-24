Attribute VB_Name = "MxIde_Ctl_IdeObj"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Ctl_IXxx."
Sub ClsIWinImm()
DoEvents
IWinImm.Visible = False
Interaction.SendKeys "^{F4}", Wait:=True
End Sub
Sub ClrIWinImm()
DoEvents
IWinImm.Visible = True
IWinImm.SetFocus
DoEvents
Interaction.SendKeys "^{Home}^+{End}", Wait:=True
DoEvents
'Interaction.SendKeys "{Del}" ' Cannot Del!!
'DoEvents
End Sub
Function IBtnEdtClr() As Office.CommandBarButton: Set IBtnEdtClr = CapFst(IPopEdt.Controls, "C&lear"):      End Function
Function IBarStd() As Office.CommandBar:             Set IBarStd = IBarsC("Standard"):                      End Function
Function IBarMnu() As CommandBar:                    Set IBarMnu = IBarMnuzV(CVbe):                         End Function
Function IWinImm() As VbIde.Window:                  Set IWinImm = IWinFst(vbext_wt_Immediate):             End Function
Function IWinLcl() As VbIde.Window:                  Set IWinLcl = IWinFst(vbext_wt_Locals):                End Function
Function IWinObj() As VbIde.Window:                  Set IWinObj = IWinFst(vbext_wt_Browser):               End Function
Function IWinPj() As VbIde.Window:                    Set IWinPj = IWinFst(vbext_wt_ProjectWindow):         End Function
Function IBtnSelAll() As Office.CommandBarButton: Set IBtnSelAll = CapFst(IPopEdt.Controls, "Select &All"): End Function


Private Sub B_IBarMnu()
Dim A As CommandBar
Set A = IBarMnu
Stop
End Sub

Function IBtnNxtStmt() As Office.CommandBarButton: Set IBtnNxtStmt = IPopDbg.Controls("Show Next Statement"): End Function
Function IBtnTileH() As Office.CommandBarButton:     Set IBtnTileH = IPopWin.Controls("Tile &Horizontally"):  End Function
Function IBtnTileV() As Office.CommandBarButton:     Set IBtnTileV = IPopWin.Controls("Tile &Vertically"):    End Function
Function IBtnSav() As Office.CommandBarButton:         Set IBtnSav = IBtnSavzV(CVbe):                         End Function
Function IBtnXls() As Office.CommandBarControl:        Set IBtnXls = IBarStd.Controls(1):                     End Function
Function IBtnCompile() As Office.CommandBarButton: Set IBtnCompile = IPopDbg.CommandBar.Controls(1):          End Function
Function IPopWin() As CommandBarPopup:                 Set IPopWin = IBarMnu.Controls("Window"):              End Function
Function IPopDbg() As CommandBarPopup:                 Set IPopDbg = IBarMnu.Controls("Debug"):               End Function
Function IPopEdt() As CommandBarPopup:                 Set IPopEdt = CapFst(IBarMnu.Controls, "&Edit"):       End Function

Function IBtnSavzV(V As VBE) As CommandBarButton
Dim I As CommandBarControl: For Each I In IBarStdzV(V).Controls
'    Debug.Print I.Caption
    If HasPfx(I.Caption, "&Save") Then Set IBtnSavzV = I: Exit Function
Next
End Function

Function CIWin() As VbIde.Window: Set CIWin = CPne.Window: End Function
