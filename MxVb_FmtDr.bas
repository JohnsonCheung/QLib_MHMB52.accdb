Attribute VB_Name = "MxVb_FmtDr"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_FmtDr."
Function LnDrAli$(Dr, Q As Qmk, W%())
Dim IsAliRy() As Boolean
LnDrAli = LnDr(DrAli(Dr, W, IsAliRy), Q) '#Dta-Row-Line# ret a formatted line of @Dr
End Function
Function LnDr$(Dr, Q As Qmk)
If Si(Dr) = 0 Then Exit Function
Dim O$()
With Q
    PushI O, .Q1
    Dim U%: U = UB(Dr)
    Dim J%: For J = 0 To U - 1
        PushI O, Dr(J)
        PushI O, .Sep
    Next
    PushI O, Dr(J)
    PushI O, .Q2
End With
LnDr = RTrim(Jn(O))
End Function
