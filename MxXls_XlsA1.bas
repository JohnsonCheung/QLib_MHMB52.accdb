Attribute VB_Name = "MxXls_XlsA1"
Option Compare Text
Option Explicit
Const CMod$ = "MxXls_XlsA1."
Function A1NwWb(B As Workbook, Optional Wsn$) As Range: Set A1NwWb = A1Ws(WsAdd(B, Wsn)):   End Function  ' Return A1 of a new Ws (with NwWsn) in @B
Function A1Lo(L As ListObject) As Range:                  Set A1Lo = A1Rg(L.DataBodyRange): End Function
Function A1Rg(R As Range) As Range:                       Set A1Rg = R.Cells(1, 1):         End Function
Function A2Ws(S As Worksheet) As Range:                   Set A2Ws = S.Cells(1, 2):         End Function
Function A1Ws(S As Worksheet) As Range:                   Set A1Ws = S.Cells(1, 1):         End Function
