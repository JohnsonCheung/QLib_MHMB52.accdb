Attribute VB_Name = "MxDao_Sql_Fmt_StruSql"
Option Compare Text
Const CMod$ = "MxDao_Sql_Stru."
Option Explicit
Private Sub B_StruSqlC()
Dim Lsy$()
Dim Qd As QueryDef: For Each Qd In CDb.QueryDefs
    Dim Q$: Q = Qd.Sql
    PushI Lsy, LinesUL(Qd.Name) & vbCrLf & LinesUL(JnCrLf(StruSqlC(Q)), "-") & vbCrLf & FmtSql(Q)
Next
BrwLsy Lsy
End Sub

Function StruSqlC(Sql) As String():               StruSqlC = StruSql(CDb, Sql):       End Function
Function StruSql(D As Database, Sql) As String():  StruSql = StruTny(D, TnySql(Sql)): End Function
