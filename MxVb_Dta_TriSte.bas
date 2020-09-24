Attribute VB_Name = "MxVb_Dta_TriSte"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Dta_TriSte."
Enum eTri: eTriOpn: eTriYes: eTriNo: End Enum
Enum eSel: eSelYN: eSelN: eSelY: End Enum
Function BoolTri(A As eTri) As Boolean
Select Case True
Case A = eTriYes: BoolTri = True
Case A = eTriNo:  BoolTri = False
Case Else: Stop
End Select
End Function

Function HitSel(B As Boolean, S As eSel) As Boolean
Select Case True
Case S = eSelYN: HitSel = True
Case S = eSelY: HitSel = B = True
Case S = eSelN: HitSel = B = False
End Select
End Function
