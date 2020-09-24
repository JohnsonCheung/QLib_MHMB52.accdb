Attribute VB_Name = "MxXls_Rg_RgOp_PutAt"
Option Compare Text
Option Explicit
Const CMod$ = "MxXls_Rg_PutDtaAt."

Function PutDbtAt(Db As Database, T, At As Range) As Range
Set PutDbtAt = RgSq(SqT(Db, T), At)
End Function
Private Sub B_PutSSVert()
Dim S As Worksheet: Set S = WsNw
PutSSVert "1 2 3 3 4 5", A1Ws(S)
Maxv S.Application
End Sub
Sub PutAyHori(HoriAy, At As Range):  RgSq SqRow(HoriAy), At:         End Sub
Sub PutAyVert(VertAy, At As Range):  RgSq SqCol(VertAy), At:         End Sub
Sub PutSSHori(HoriSS$, At As Range): PutAyHori SplitSpc(HoriSS), At: End Sub
Sub PutSSVert(VertSS$, At As Range): PutAyVert SplitSpc(VertSS), At: End Sub
