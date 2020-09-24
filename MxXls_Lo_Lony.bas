Attribute VB_Name = "MxXls_Lo_Lony"
Option Compare Text
Option Explicit
Const CMod$ = "MxXls_Lo_Lony."

Function LonyWs(S As Worksheet) As String(): LonyWs = Itn(S.ListObjects): End Function
Function LonyWb(B As Workbook) As String()
Dim S As Worksheet: For Each S In B.Sheets
    PushIAy LonyWb, LonyWs(S)
Next
End Function
Function LonyFx(Fx) As String()
Dim B As Workbook: Set B = WbFx(Fx)
LonyFx = LonyWb(B)
B.Close
End Function
