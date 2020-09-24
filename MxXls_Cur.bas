Attribute VB_Name = "MxXls_Cur"
Option Compare Text
Option Explicit
Const CMod$ = "MxXls_Cur."
Function CWs() As Worksheet:         Set CWs = Xls.ActiveSheet:    End Function
Function CWb() As Workbook:          Set CWb = Xls.ActiveWorkbook: End Function
Function Xls() As Excel.Application: Set Xls = Excel.Application:  End Function
