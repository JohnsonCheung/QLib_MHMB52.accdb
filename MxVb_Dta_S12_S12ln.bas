Attribute VB_Name = "MxVb_Dta_S12_S12ln"
'#l:sfx:Lines#
'#lny:sfx:Line-Array#
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Dta_S12_Ln."
Function S12lnS12$(S As S12, Optional Sep$ = " = "):    S12lnS12 = S.S1 & Sep & S.S2:           End Function ' #S12-Line# A
Function S12S12ln(S12ln, Optional Sep$ = "=") As S12:   S12S12ln = Brk(S12ln, Sep):             End Function
Function S12yS12l(S12l$, Optional Sep$ = "=") As S12(): S12yS12l = S12yS12lny(SplitCrLf(S12l)): End Function
Function S12yS12lny(S12lny$(), Optional Sep$ = "=") As S12()
Dim S: For Each S In Itr(S12lny)
    PushS12 S12yS12lny, S12S12ln(S, Sep)
Next
End Function

Function S12lS12y$(S() As S12, Optional Sep$ = " = "): S12lS12y = JnCrLf(S12lnyS12y(S)): End Function '#S12-Lines# Lines of :S12ln
Function S12lnyS12y(S() As S12, Optional Sep$ = " = ") As String()
Dim J&, O$(): For J = 0 To UbS12(S)
     PushIAy S12lnyS12y, S12lnS12(S(J))
Next
End Function

Function S12yT1ry(T1ry$()) As S12()
Dim L: For Each L In Itr(T1ry)
    PushS12 S12yT1ry, S12T1r(L)
Next
End Function
Function S12T1r(T1r) As S12
Dim T1$, R$
R = T1r
T1 = ShfTm(R)
S12T1r = S12(T1, R)
End Function
