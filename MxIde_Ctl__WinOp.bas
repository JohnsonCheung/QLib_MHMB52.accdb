Attribute VB_Name = "MxIde_Ctl__WinOp"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Ctl__WinOp."

Sub ClsWinAll()
Dim W As VbIde.Window: For Each W In CVbe.Windows
    ClsWin W
Next
TileV
End Sub

Sub VisIWin(W As VbIde.Window): W.Visible = True: End Sub

Sub ClsWinExlAp(ParamArray ExlWinAp())
Dim Av(): Av = ExlWinAp
Dim I: For Each I In Itr(IWinyVis)
    Dim W As VbIde.Window: Set W = I
    If Not HasObj(Av, W) Then
        ClsWin W
    Else
        VisIWin W
    End If
Next
End Sub

Sub ShwDbg()
ClsWinExlAp IWinImm, IWinLcl, CIWin
DoEvents
TileV
End Sub
Private Sub B_ClrImm()
Dim J%
Debug.Print "lskdfjsdlkf"
Wait 2
Debug.Print "lskdfjsdlkf"
Wait
ClrImm
End Sub
Sub ClrImm()
DoEvents
With IWinImm
    .SetFocus
End With
DoEvents
SndKeys "^{HOME}^+{END}"
SndKeys "{DEL}"
DoEvents
End Sub

Sub ClsWin(W As VbIde.Window)
If W.Visible Then W.Close
End Sub

Sub ClrWin(A As VbIde.Window)
DoEvents
IBtnSelAll.Execute
DoEvents
SendKeys " "
IBtnEdtClr.Execute
End Sub
