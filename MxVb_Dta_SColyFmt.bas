Attribute VB_Name = "MxVb_Dta_SColyFmt"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Dta_SColyFmt."
Function LyDcyStrAp(Sep$, ParamArray DcyStrAp()) As String()
Dim Av(): Av = DcyStrAp
LyDcyStrAp = LyDcyStr(Av, Sep)
End Function
Function LyDcyStr(DcyStr(), Optional Sep$ = " ") As String()
'Assume each element of @DcyStr must be :Sy
Dim O$(): O = DcyStr(0)
Dim J&: For J = 1 To UB(DcyStr)
    O = DcStrJn(O, CvSy(DcyStr(J)), Sep)
Next
LyDcyStr = O
End Function
Function DcStrJn(DcStr1$(), DcStr2$(), Optional Sep$ = " ") As String()
If Si(DcStr1) = 0 Then DcStrJn = DcStr2: Exit Function
If Si(DcStr2) = 0 Then DcStrJn = DcStr1: Exit Function
Dim A$(), B$(): WAsgAB DcStr1, DcStr2, A, B
A = AmAli(A)
Dim J&: For J = 0 To UB(A)
    PushI DcStrJn, A(J) & Sep & B(J)
Next
End Function
Private Sub WAsgAB(DcStr1$(), DcStr2$(), OA$(), OB$())
Dim U1&: U1 = UB(DcStr1)
Dim U2&: U2 = UB(DcStr2)
Dim U&: U = Max(U1, U2)
OA = DcStr1: If U > U1 Then ReDim Preserve OA(U)
OB = DcStr2: If U > U2 Then ReDim Preserve OB(U)
End Sub
