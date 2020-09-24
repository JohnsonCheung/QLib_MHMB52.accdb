Attribute VB_Name = "MxXls_Ws_Dta"
Option Compare Text
Option Explicit
Const CMod$ = "MxXls_Ws_Dta."

Function DtaRg(S As Worksheet) As Range
Set DtaRg = S.Range(S.Cells(1, 1), LasCell(S))
End Function

Function DtaSq(S As Worksheet) As Variant()
DtaSq = DtaRg(S).Value
End Function

Function DtaDrs(S As Worksheet) As Drs
DtaDrs = DrsSq(DtaSq(S))
End Function
