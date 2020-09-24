Attribute VB_Name = "MxVb_Dta_S12_S12y_S12yMix"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Dta_S12_S12y_Mix."
Function LsyS12yMix(S() As S12) As String()
Dim J&: For J = 0 To UbS12(S)
    With S(J)
        PushI LsyS12yMix, LinesMix(.S1, .S2)
    End With
Next
End Function
