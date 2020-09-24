Attribute VB_Name = "MxXls_Lo_Dta"
Option Compare Text
Option Explicit
Const CMod$ = "MxXls_Lo_Dta."
Function DrLoCell(Lo As ListObject, Cell As Range) As Variant()
Dim Ix&: Ix = LoRno(Lo, Cell): If Ix = -1 Then Exit Function
DrLoCell = DrFstRg(Lo.ListRows(Ix).Range)
End Function
