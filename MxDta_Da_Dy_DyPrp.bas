Attribute VB_Name = "MxDta_Da_Dy_DyPrp"
Option Compare Text
Option Explicit
Const CMod$ = "MxDta_Da_Dy_DyPrp."

Function WdtyDr(Dr) As Integer()
Dim V: For Each V In Dr
    PushI WdtyDr, WdtLines(CStr(V))
Next
End Function
Function WdtyDy(Dy(), Optional Wdt% = 120) As Integer() ' If MaxWdt=0, there is no limit to the width
If Si(Dy) = 0 Then Exit Function
Dim UCol%: UCol = NDcDy(Dy) - 1
Dim W%()
    ReDim W(UCol)
    Dim J%: For J = 0 To UCol
        Dim Dc$(): Dc = DcStrDy(Dy, J)
        W(J) = WdtLsy(Dc)
    Next

    If Wdt > 0 Then
        For J = 0 To UCol
            W(J) = Min(Wdt, W(J))
        Next
    End If
WdtyDy = W
End Function
