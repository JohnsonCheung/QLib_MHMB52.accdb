Attribute VB_Name = "MxVb_Dta_S12_XxxS12"
Option Compare Text
Option Explicit

Function S12yS1ya2(S1$(), S2$()) As S12()
ChkIsEqSi S1, S2, CSub
Dim J&: For J = 0 To UB(S1)
    PushS12 S12yS1ya2, S12(S1(J), S2(J))
Next
End Function
