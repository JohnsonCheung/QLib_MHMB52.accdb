Attribute VB_Name = "MxXls_Ws_Prp"
Option Compare Text
Option Explicit
Const CMod$ = "MxXls_Ws_Prp."

Function HasLo(S As Worksheet, Lon$) As Boolean
HasLo = HasItn(S.ListObjects, Lon)
End Function

Function LasCell(S As Worksheet) As Range
Set LasCell = S.Cells.SpecialCells(xlCellTypeLastCell)
End Function

Function LasRno&(S As Worksheet)
LasRno = LasCell(S).Row
End Function

Function LasCno%(S As Worksheet)
LasCno = LasCell(S).Column
End Function

Function PtNyWs(S As Worksheet) As String()
PtNyWs = Itn(S.PivotTables)
End Function

Property Get MaxCnoX&(X As Excel.Application)
MaxCnoX = IIf(X.Version = "16.0", 16384, 255)
End Property

Property Get MaxRnoX&(X As Excel.Application)
MaxRnoX = IIf(Xls.Version = "16.0", 1048576, 65535)
End Property
