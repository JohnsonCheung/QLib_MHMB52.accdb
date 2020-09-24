Attribute VB_Name = "MxXls_Lo_Inf"
Option Compare Text
Option Explicit
Const CMod$ = "MxXls_Lo_Inf."
Public Const LoInfFf$ = "Wsn Lon R C NR NC Ff"

Private Sub B_DrsTLoInf()
Dim B As Workbook: Set B = WbFx(MH.MB52Tp.Tp)
Dim D As Drs: D = DrsTLoInf(B)
ClsWbNoSav B
Stop
BrwDrs D
End Sub
Function DrsTLoInf(B As Workbook) As Drs: DrsTLoInf = DrsFf(LoInfFf, WDyWb(B)): End Function
Private Function WDyWb(B As Workbook) As Variant()
Dim S As Worksheet: For Each S In B.Sheets
    PushIAy WDyWb, WDyWs(S)
Next
End Function
Private Function WDyWs(S As Worksheet) As Variant()
Dim L As ListObject: For Each L In S.ListObjects
    PushI WDyWs, WDr(L)
Next
End Function
Private Function WDr(L As ListObject) As Variant()
Dim Wsn: Wsn = WsnLo(L)
Dim Lon$:: Lon = L.Name
Dim NR&: NR = NRowLo(L)
Dim NC&: NC = L.ListColumns.Count
WDr = Array(WsnLo(L), L.Name, L.Range.Row, L.Range.Column, NR, NC, TmlAy(FnyLo(L)))
End Function
