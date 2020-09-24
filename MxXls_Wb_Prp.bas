Attribute VB_Name = "MxXls_Wb_Prp"
Option Compare Text
Option Explicit
Const CMod$ = "MxXls_Wb_Prp."
Function WsLas(B As Workbook) As Worksheet: Set WsLas = B.Sheets(B.Sheets.Count): End Function
Function WsFst(B As Workbook) As Worksheet: Set WsFst = B.Sheets(1):              End Function

Function FxWb$(B As Workbook)
Dim F$: F = B.FullName
If F = B.Name Then Exit Function
FxWb = F
End Function

Function MainLo(B As Workbook) As ListObject
Dim O As Worksheet, Lo As ListObject
Set O = WsMain(B):              If IsNothing(O) Then Exit Function
Set MainLo = LoWs(O, "T_OupMain")
End Function

Function MainQt(B As Workbook) As QueryTable
Dim L As ListObject: Set L = MainLo(B): If IsNothing(B) Then Exit Function
Set MainQt = L.QueryTable
End Function

Function Ptny(B As Workbook) As String()
Dim Ws As Worksheet
For Each Ws In B.Sheets
    PushIAy Ptny, PtNyWs(Ws)
Next
End Function

Function WsWb(B As Workbook, Wsix) As Worksheet
If HasWsn(B, Wsix) Then Set WsWb = B.Sheets(Wsix)
End Function

Function IsWbGood(Wb As Workbook)
On Error GoTo X
Dim A$: A = Wb.Name
IsWbGood = True
Exit Function
X:
End Function
