Attribute VB_Name = "MxVb_Dta_Dte_StrDte"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Dta_Dte_Str."

Function StrNow15$(): StrNow15 = StrDte15(Now): End Function
Function StrNow14$(): StrNow14 = StrDte14(Now): End Function
Function StrNow$():     StrNow = StrDte(Now):   End Function

Function StrDte15$(D As Date): StrDte15 = Format(D, "YYYYMMDD_HHMMSS"):     End Function
Function StrDte$(D As Date):     StrDte = Format(D, "YYYY-MM-DD HH:MM:SS"): End Function
Function StrDte14$(D As Date): StrDte14 = Format(D, "YYYYMMDDHHMMSS"):      End Function
Function StrNowUnq$()
Static I%: I = I + 1: If I = 1000 Then I = 1
StrNowUnq = StrNow15 & "_" & Pad0(I, 3)
End Function
