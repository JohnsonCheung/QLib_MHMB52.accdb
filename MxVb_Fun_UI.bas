Attribute VB_Name = "MxVb_Fun_UI"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Fun_UI."

Function CfmYes(Msg$) As Boolean
CfmYes = UCase(InputBox(Msg)) = "YES"
End Function

Sub PromptCnl(Optional Msg = "Should cancel and check")
If MsgBox(Msg, vbOKCancel) = vbCancel Then Stop
End Sub

Sub Done()
MsgBox "Done"
End Sub
Function Start(Optional Msg$ = "Start?", Optional Tit$ = "Start?") As Boolean
Start = MsgBox(Replace(Msg, "|", vbCrLf), vbQuestion + vbYesNo + vbDefaultButton1, Tit) = vbYes
End Function

Function IsCfm(Msg$, Optional Tit$ = "Please confirm", Optional NoAsk As Boolean) As Boolean
If NoAsk Then IsCfm = True: Exit Function
IsCfm = MsgBox(Msg, vbYesNo + vbQuestion + vbDefaultButton1) = vbYes
End Function
