Attribute VB_Name = "MxDta_Da_Drs_DrsIO"
Option Compare Text
Option Explicit
Const CMod$ = "MxDta_Da_Drs_DrsIO."

Sub WrtDrs(A As Drs, Ft$)
WrtAy TsyDrs(A), Ft
End Sub

Function DrsFt(Ft$) As Drs: DrsFt = WDrsTsy(LyFt(Ft)): End Function
Private Function WDrsTsy(Tsy$()) As Drs
Const CSub$ = CMod & "WDrsTsy"
Dim A$(): A = SplitCrLf(Tsy)
If Si(A) = 0 Then Thw CSub, "No lines in @Tsy"
Dim O As Drs
O.Fny = SplitTab(A(0))
If Si(A) = 1 Then Exit Function
Dim T() As VbVarType: Stop ' : T = TslTynDr(SplitTab(A(1)))
Dim J&: For J = 2 To UB(A)
    PushI O.Dy, WDrTsl(A(J), T)
Next
WDrsTsy = O
End Function

Private Function WDrTsl(Tsl, T() As VbVarType) As Variant()
WDrTsl = WDrSy(SplitTab(Tsl), T)
End Function

Private Function WDrSy(Sy$(), T() As VbVarType) As Variant()
Dim J%, S: For Each S In Sy
    PushI WDrSy, ValS(CStr(S), T(J))
    J = J + 1
Next
End Function

Function TsyDrs(D As Drs) As String()
PushI TsyDrs, JnTab(D.Fny)
Dim Dr: For Each Dr In Itr(D.Dy)
    PushI TsyDrs, JnTab(Dr)
Next
End Function
Function TsyDt(D As Dt) As String()
PushI TsyDt, "*Dt " & D.Dtn
PushI TsyDt, JnTab(D.Fny)
Dim Dr: For Each Dr In Itr(D.Dy)
    PushI TsyDt, JnTab(Dr)
Next
End Function
