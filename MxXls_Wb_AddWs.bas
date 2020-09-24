Attribute VB_Name = "MxXls_Wb_AddWs"
Option Compare Text
Option Explicit
Const CMod$ = "MxXls_Wb_AddWs."

Function WsWbDt(B As Workbook, Dt As Dt) As Worksheet
Dim O As Worksheet
Set O = WsAdd(B, Dt.Dtn)
LoDrs DrsDt(Dt), A1Ws(O)
Set WsWbDt = O
End Function

Function WsEns(B As Workbook, Wsn$, Optional P As ePosWs) As Worksheet
If HasWs(B, Wsn) Then
    Set WsEns = B.Sheets(Wsn)
Else
    Set WsEns = WsAdd(B, Wsn, P)
End If
End Function
Function WsAdd(B As Workbook, Optional Wsn$, Optional Pos As ePosWs, Optional Aft$, Optional Bef$) As Worksheet
DltWsIf B, Wsn
Dim WsRf As Worksheet, IsBef As Boolean
    Select Case True
    Case Pos = ePosWsBeg: IsBef = True:  Set WsRf = WsFst(B)
    Case Pos = ePosWsEnd: IsBef = False: Set WsRf = WsLas(B)
    Case Pos = ePosWsAft: IsBef = False: Set WsRf = WsWb(B, Aft): If IsNothing(WsRf) Then Set WsRf = WsLas(B)
    Case Pos = ePosWsBef: IsBef = True:  Set WsRf = WsWb(B, Bef): If IsNothing(WsRf) Then Set WsRf = WsFst(B)
    Case Else: Stop
    End Select
If IsBef Then
    Set WsAdd = B.Sheets.Add(WsRf)
Else
    Set WsAdd = B.Sheets.Add(, WsRf)
End If
SetWsn WsAdd, Wsn
End Function

Function WbDs(A As Ds) As Workbook
Dim O As Workbook: Set O = WbNw
With WsFst(O)
   .Name = "Ds"
   .Range("A1").Value = A.Dsn
End With
Dim Ay() As Dt: Ay = A.Dty
Dim J&: For J = 0 To DtUB(Ay)
    WsWbDt O, Ay(J)
Next
Set WbDs = O
End Function
