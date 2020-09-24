Attribute VB_Name = "MxXls_Ws_NwWs"
Option Compare Text
Option Explicit
Const CMod$ = "MxXls_Op_Nw."
Function WsAy(Ay, Optional Hdr$ = "Array", Optional Wsn$ = "Sheet1") As Worksheet
Dim O As Worksheet
Set O = WsNw(Wsn)
O.Range("A1").Value = Hdr
Dim R As Range
    Set R = RgSq(SqCol(Ay), O.Range("A2"))
LoRg RgMoreTop(R)
Set WsAy = O
End Function

Function WsLy(Ly$(), Optional Hdr$ = "Ly", Optional Wsn$) As Worksheet
Set WsLy = WsAy(Ly, Hdr, Wsn)
End Function

Function A1Nw(Optional Wsn$) As Range:     Set A1Nw = A1Ws(WsNw(Wsn)):  End Function
Function WsNw(Optional Wsn$) As Worksheet: Set WsNw = WsFst(WbNw(Wsn)): End Function
Function WbNw(Optional Wsn$) As Workbook
Dim O As Workbook
Set O = XlsNw.Workbooks.Add
SetWsn WsFst(O), Wsn
Set WbNw = O
End Function

Private Sub B_XlsNw()
GoSub Tst
Exit Sub
Dim Act As Excel.Application
'{00024500-0000-0000-C000-000000000046}
Tst:
    Set Act = XlsNw
    Stop
    Return
End Sub
Function XlsNw() As Excel.Application
' Don't use New Excel.Application, but why? becase it is not guaratee to create
Dim O As Excel.Application: Set O = CreateObject("Excel.Application")
O.DisplayAlerts = False
Set XlsNw = O
Minvn O
End Function
Function WsDt(A As Dt) As Worksheet
Dim O As Worksheet
Set O = WsNw(A.Dtn)
LoDrs DrsDt(A), A1Ws(O)
WsDt = O
End Function
