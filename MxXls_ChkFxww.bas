Attribute VB_Name = "MxXls_ChkFxww"
Option Compare Text
Option Explicit
Const CMod$ = "MxXls_ChkFxww."
Sub ChkFxww(Fx$, WW$, Optional Kd$ = "Excel file")
End Sub
Function EryFxwwMis(Fx, WW, Optional Kd$ = "Excel file") As String()
ChkFfnExi Fx, Kd
Dim Wny$(): Wny = WnyFx(Fx)
Dim WnyMis$(): WnyMis = SyMinus(SySs(WW), Wny)
    If Si(WnyMis) = 0 Then Exit Function
Dim Er$()
    Dim M$: M = FmtQQ("Missing ? worksheet", Si(WnyMis))
    PushI Er, M
    PushI Er, LyUL(M)
    PushI Er, "Excel File    : [" & Fx & "]"
    PushI Er, "Has worksheets: [" & Wny(0) & "]"
    
    Dim J%: For J = 1 To UB(Wny)
        PushI Er, "                [" & Wny(J) & "]"
    Next
    PushI Er, "Missing       : [" & WnyMis(0) & "]"
    For J = 1 To UB(WnyMis)
    PushI Er, "                [" & WnyMis(J) & "]"
    Next
BrwSyIfNB Er



End Function

Private Function WErFxwMis(Fx, W, Optional FilKd$ = "Excel file") As String()
If HasFxw(Fx, W) Then Exit Function
Erase XX
X FmtQQ("[?] miss ws [?]", FilKd, W)
X vbTab & "Path  : " & Pth(Fx)
X vbTab & "File  : " & Fn(Fx)
X vbTab & "Has Ws: " & TmlAy(WnyFx(Fx))
WErFxwMis = XX
Erase XX
End Function
