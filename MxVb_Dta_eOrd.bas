Attribute VB_Name = "MxVb_Dta_eOrd"
Option Compare Text
Option Explicit
Enum eOrd: eOrdLT: eOrdEQ: eOrdGT: End Enum
Function eOrdDte(A As Date, B As Date) As eOrd: eOrdDte = eOrdV(A, B): End Function
Function eOrdV(A, B) As eOrd
Select Case True
Case A > B: eOrdV = eOrdGT
Case B < A: eOrdV = eOrdLT
Case Else: eOrdV = eOrdEQ
End Select
End Function
