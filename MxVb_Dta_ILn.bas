Attribute VB_Name = "MxVb_Dta_ILn"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Dta_ILn."

Function LyILny(L() As TIxLn) As String()
Dim J%: For J = 0 To UbTIxLn(L)
    PushI LyILny, L(J).Ln
Next
End Function
