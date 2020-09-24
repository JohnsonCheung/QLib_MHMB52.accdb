Attribute VB_Name = "MxXls_Cdny"
Option Compare Text
Const CMod$ = "MxXls_Cdny."
Option Explicit
Function CdnWb$(Fx)
Dim B As Workbook
Set B = WbFx(Fx)
CdnWb = B.CodeName
B.Close False
End Function
Function CdnyWsC():     CdnyWsC = CdnyWs(CWb):   End Function
Function CdnyWbWsC(): CdnyWbWsC = CdnyWbWs(CWb): End Function
Function CdnyWbWs(Fx) As String()
Dim B As Workbook
Set B = WbFx(Fx)
PushI CdnyWbWs, B.CodeName
PushIAy CdnyWbWs, WWCdny(B)
B.Close False
End Function
Function CdnyWs(Fx) As String()
Dim B As Workbook
Set B = WbFx(Fx)
CdnyWs = WWCdny(B)
B.Close False
End Function
Private Function WWCdny(B As Workbook) As String(): WWCdny = SyItp(B.Sheets, "CodeName"): End Function
