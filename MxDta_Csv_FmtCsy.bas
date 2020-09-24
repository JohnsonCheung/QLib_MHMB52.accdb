Attribute VB_Name = "MxDta_Csv_FmtCsy"
Option Compare Text
Option Explicit
Const CMod$ = "MxDta_Csv_FmtCsy."

Function FmtCsy(Csy$(), Optional IsNoLnFny As Boolean) As String()
FmtCsy = FmtDrs(DrsCsy(Csy))
End Function
Function DrsCsy(Csy$(), Optional IsNoLnFny As Boolean) As Drs
If Si(Csy) = 0 Then Exit Function
Dim NDc%: NDc = WNDcCsy(Csy)
Dim OFny$(): OFny = WFny(Csy(0), IsNoLnFny, NDc)
Dim ODy(): ODy = WDy(Csy, IsNoLnFny, NDc)
DrsCsy = Drs(OFny, ODy)
End Function
Private Function WDy(Csy$(), IsNoLnFny As Boolean, NDc%) As Variant()

End Function
Private Function WNDcCsy%(Csy$())

End Function
Private Function WFny(Csvln$, IsNoLnFny As Boolean, NDc%) As String()

End Function
