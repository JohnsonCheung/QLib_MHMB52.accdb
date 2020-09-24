Attribute VB_Name = "MxXls_Ws_Oln"
Option Compare Text
Option Explicit
Const CMod$ = "MxXls_Ws_Oln."

Sub MiniWbOLvl(Wb As Workbook)
Dim Ws As Worksheet: For Each Ws In Wb.Sheets
    MiniWsOLvl Ws
Next
End Sub
Sub MiniWsOLvl(Ws As Worksheet)
Ws.Outline.ShowLevels 1, 1
End Sub
