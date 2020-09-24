Attribute VB_Name = "MxIde_Ctl_Prp_ICtlPrp"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Ctl_Fun."
Function IBarsVbe(V As VBE) As Office.CommandBars: Set IBarsVbe = VBE.CommandBars:                      End Function
Function CapyIdeCtl(C As Controls) As String():      CapyIdeCtl = SyItrPrp(C, "Caption"):               End Function
Function CvIdeCtl(V) As CommandBarControl:         Set CvIdeCtl = V:                                    End Function
Function WinMdn(Mdn) As VbIde.Window:                Set WinMdn = Md(Mdn).CodePane.Window:              End Function
Function WinMd(M As CodeModule) As VbIde.Window:      Set WinMd = M.CodePane.Window:                    End Function
Function CapFst(Itr, Caption):                           CapFst = ItoFstPrpEq(Itr, "Caption", Caption): End Function
Function HasIBar(IBarn) As Boolean:                     HasIBar = HasIdeIBarV(CVbe, IBarn):             End Function
Function IsVIbtn(V) As Boolean:                         IsVIbtn = TypeName(V) = "CommandButton":        End Function
Function CvObtn(V) As CommandBarButton:              Set CvObtn = V:                                    End Function
Function CvWin(V) As VbIde.Window:                    Set CvWin = V:                                    End Function
Private Sub B_IPopDbg()
Dim A
Set A = IPopDbg
Stop
End Sub

Function PneCmpn(Cmpn$) As CodePane:          Set PneCmpn = Md(Cmpn).CodePane:                   End Function
Function IBar(IBarn) As CommandBar:              Set IBar = IBarsC(IBarn):                       End Function
Function IBarStdzV(V As VBE) As CommandBar: Set IBarStdzV = V.CommandBars("Standard"):           End Function
Function IBarMnuzV(V As VBE) As CommandBar: Set IBarMnuzV = V.CommandBars("Menu Bar"):           End Function
Sub ClsWinExlMdn(ExlMdn$):                                  ClsWinExlAp IWinImm, WinMdn(ExlMdn): End Sub
Function IBarnyV(V As VBE) As String():           IBarnyV = Itn(V.CommandBars):                  End Function

Function RrccPne(P As CodePane) As Rrcc
Dim R1&, R2&, C1&, C2&
P.GetSelection R1, R2, C1, C2
RrccPne = Rrcc(R1, R2, C1, C2)
End Function

Function IWinFst(A As vbext_WindowType) As VbIde.Window:  Set IWinFst = ItoFstPrpEq(CVbe.Windows, "Type", A):       End Function
Function WinyTy(T As vbext_WindowType) As VbIde.Window():      WinyTy = IntoItwEq(WinyTy, CVbe.Windows, "Type", T): End Function
Function MdnWin$(WinCd As VbIde.Window):                       MdnWin = IsBet(WinCd.Caption, " - ", " (Code)"):     End Function

Function IBarsC() As Office.CommandBars: Set IBarsC = CVbe.CommandBars: End Function
Function InIBarBtn(Bar As CommandBar, BtnCap) As Boolean
Dim C As CommandBarControl
For Each C In Bar.Controls
    If C.Type = msoControlButton Then
        If C.Caption = BtnCap Then InIBarBtn = True: Exit Function
    End If
Next
End Function
