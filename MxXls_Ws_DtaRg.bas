Attribute VB_Name = "MxXls_Ws_DtaRg"
Option Compare Text
Option Explicit
Const CMod$ = "MxXls_Ws_DtaRg."
Sub ClrDtarg(A As Worksheet)
DtaDtarg(A).Clear
End Sub

Function DtaDtarg(Ws As Excel.Worksheet) As Range
Set DtaDtarg = Ws.Range(A1Ws(Ws), LasCell(Ws))
End Function
