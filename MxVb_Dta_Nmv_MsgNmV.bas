Attribute VB_Name = "MxVb_Dta_Nmv_MsgNmV"
Option Compare Text
Const CMod$ = "MxVb_Dta_Nmv_MsgNmV."
Option Explicit

Function MsglyNmV(Nm$, V) As String()
Dim Ly$(): Ly = LyV(V)
PushI MsglyNmV, Nm & ": " & EleFst(Ly)
If Si(Ly) <= 1 Then Exit Function
Dim S$: S = Space(Len(Nm) + 2)
Dim J%: For J = 1 To UB(Ly)
    PushI MsglyNmV, S & Ly(J)
Next
End Function
