Attribute VB_Name = "MxVb_Ay_Nw_PrimySno"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Ay_NwSnoy."

Function FnySno(Pfx$, N%, Optional IsFmZer As Boolean) As String()
Dim I01%: I01 = IIf(IsFmZer, 0, 1)
Dim NDigit%: NDigit = NDig(N + I01 - 1)
Dim J%: For J = 0 + I01 To N + I01 - 1
    PushI FnySno, Pfx & Pad0(J, NDigit)
Next
End Function
Function IntySno(N, Optional Fm = 0) As Integer(): IntySno = SnoyInto(IntyEmp, N, Fm): End Function
Function LngySno(N, Optional Fm = 0) As Long():    LngySno = SnoyInto(LngyEmp, N, Fm): End Function
Function BytySno(N, Optional Fm = 0) As Byte():    BytySno = SnoyInto(BytyEmp, N, Fm): End Function
Private Function SnoyInto(Into, N, Optional Fm = 0)
Dim O: O = Into: Erase O
Dim J&: For J = Fm To N + Fm - 1
    PushI O, J
Next
SnoyInto = O
End Function
