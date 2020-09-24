Attribute VB_Name = "MxDao_Sql_Fmt_zIntl_FmtSqlTQp"
Option Compare Text
Option Explicit
Const CMod$ = "MxDao_Sql_FmtSqlTQp."
Private Sub B_FmtSqlTQpDb()
GoSub ZZ
Exit Sub
ZZ:
    BrwAy FmtSqlTQpDb(CDb)
    Return
End Sub
Function FmtSqlStruTQp$(D As Database, Sql)
FmtSqlStruTQp = FmtSqlTQp(Sql) & vbCrLf & JnCrLf(StruSql(D, Sql))
End Function
Function FmtSqlTQp$(Sql)
Dim QSql() As TQp ' Fm Sql
Dim QTb() As TQp  ' Aft Fmt Tb
Dim QFld() As TQp ' Aft Fmt FldLis
QSql = TQpySql(Sql)
QFld = TQpyFmtFld(QSql): 'ShwTimr "TQpyFmtFld...!FmtSqlTQp"
QTb = TQpyFmtTb(QFld):  'ShwTimr "TQpyFmtTb...!FmtSqlTQp"
FmtSqlTQp = SqlTQpy(QTb):  'ShwTimr "SqlTQpy........!FmtSqlTQp"
End Function
Function FmtSqlTQpDb(D As Database) As String()
Dim Lsy$()
Dim Qd As QueryDef: For Each Qd In CDb.QueryDefs
    Dim Q$: Q = Qd.Sql
    PushI Lsy, LinesUL(Qd.Name) & vbCrLf & LinesUL(LinesEndTrim(Q), "-") & vbCrLf & FmtSqlStruC(Q)
Next
FmtSqlTQpDb = FmtLsy(Lsy)
End Function
