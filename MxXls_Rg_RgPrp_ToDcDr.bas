Attribute VB_Name = "MxXls_Rg_RgPrp_ToDcDr"
Option Compare Text
Option Explicit
Const CMod$ = "MxXls_Rg_Val."
Function DrFstRg(A As Range) As Variant()
DrFstRg = DrFstSq(SqRg(RgR(A, 1)))
End Function
Function DcFstRg(A As Range) As Variant()
DcFstRg = DcFstSq(SqRg(RgC(A, 1).Value))
End Function

Function DcIntAt(At As Range) As Integer(): DcIntAt = DcIntSq(SqRg(RgAtDown(At))): End Function
Function DcStrAt(At As Range) As String():  DcStrAt = DcStrSq(SqRg(RgAtDown(At))): End Function
