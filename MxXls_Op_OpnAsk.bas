Attribute VB_Name = "MxXls_Op_OpnAsk"
Option Compare Text
Option Explicit
Const CMod$ = "MxXls_Op_OpnAsk."
Function AskOpnFx(Fx) As Boolean
If NoFfn(Fx) Then Exit Function
Select Case MsgBox("File exists:" & Fx & vbLf & vbLf & _
    "[Yes] = Re-generate and over-write" & vbLf & _
    "[No] = Open existing file" & vbLf & _
    "[Cancel] = Cancel", vbYesNoCancel + vbDefaultButton2 + vbQuestion, "Generate file.")
Case VbMsgBoxResult.vbNo:
    MaxvFx Fx
    AskOpnFx = True
    Exit Function
End Select
End Function

Function AskOpnFxy(Fxy$()) As Boolean
Dim Fx: For Each Fx In Itr(Fxy)
    If NoFfn(Fx) Then Exit Function
Next
Dim A$: A = JnCrLf(Fxy)
Select Case MsgBox("File exists:" & A & vbLf & vbLf & _
    "[Yes] = Re-generate and over-write" & vbLf & _
    "[No] = Open existing file" & vbLf & _
    "[Cancel] = Cancel", vbYesNoCancel + vbDefaultButton2 + vbQuestion, "Generate file.")
Case VbMsgBoxResult.vbNo:
    OpnFxy Fxy
    AskOpnFxy = True
    Exit Function
End Select
End Function
