Attribute VB_Name = "MxVb_Dta_Di_DiCnt"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Dta_DicCnt."
Function CntDupzDiCnt(DiCnt As Dictionary) As Dictionary
Set CntDupzDiCnt = New Dictionary
Dim Cnt&, K
For Each K In DiCnt.Keys
    Cnt = DiCnt(K)
    If Cnt > 1 Then CntDupzDiCnt.Add K, Cnt
Next
End Function

Function VyDiszDiCnt(DiCnt As Dictionary) As String()
Dim K: For Each K In DiCnt.Keys
    If DiCnt(K) = 0 Then PushI VyDiszDiCnt, K
Next
End Function

Function DiCntDrs(A As Drs, C$) As Dictionary
Set DiCntDrs = DiCnt(DcDrs(A, C))
End Function
