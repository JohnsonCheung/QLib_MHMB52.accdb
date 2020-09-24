Attribute VB_Name = "MxDao_Db_Op"
Option Compare Text
Option Explicit
Const CMod$ = "MxDao_Db_Op."

Sub RunqQQ(D As Database, QQ$, ParamArray Ap()): Dim Av(): Av = Ap: RunqQQAv D, QQ, Av:   End Sub
Sub RunqQQC(QQ$, ParamArray Ap()):               Dim Av(): Av = Ap: RunqQQAv CDb, QQ, Av: End Sub
Sub RunqQQAv(D As Database, QQ$, Av()):          Runq D, FmtQQAv(QQ, Av):                 End Sub 'Ret : Run the %Sql by building from &FmtQQ(@QQ,@Av) in @D
Sub RunqC(Q$):                                   Runq CDb, Q:                             End Sub
Sub Runq(D As Database, Q)
Const CSub$ = CMod & "Runq"
On Error GoTo X
D.Execute Q
Exit Sub
X:
    ThwSql CSub, D, Q, Err.Description
End Sub
Sub ThwSql(Fun$, D As Database, Q, Er$)
Dim T$: T = Tmpn("#Qry_")
Dim S$: S = FmtSql(Q)
Dim Stru$(): Stru = StruSql(D, Q)
Crtq D, T, Q
Thw Fun, Er, "Sql Stru Db TmpQryNm", S, Stru, D.Name, T
End Sub
Sub RunqSqy(D As Database, Sqy$())
Dim Q: For Each Q In Sqy
    Runq D, Q
Next
End Sub
Sub RunqSqyC(Sqy$()): RunqSqy CDb, Sqy: End Sub

Sub Crtq(D As Database, Qn$, Sql)
On Error GoTo X
D.QueryDefs.Append QdNw(Qn, Sql)
Exit Sub
X:
    Debug.Print "Crtq: Cannot create query: " & Err.Description
End Sub
Sub DltqTmpC(): DltqTmp CDb: End Sub
Sub DltqTmp(D As Database)
Dim Q: For Each Q In Itr(QnyTmp(D))
    D.QueryDefs.Delete Q
Next
End Sub

Function QdNw(Qn$, Sql) As Dao.QueryDef
Set QdNw = New Dao.QueryDef: QdNw.Name = Qn: QdNw.Sql = Sql
End Function

Function QnTmpCrt$(D As Database, Sql$, Optional QnPfx$)
Dim N$: N = Tmpn("#Qry_" & QnPfx)
Crtq D, N, Sql
QnTmpCrt = N
End Function

Function QnyTmpC() As String():             QnyTmpC = QnyTmp(CDb):            End Function
Function QnyTmp(D As Database) As String():  QnyTmp = AwPfx(Qny(D), "#Qry_"): End Function
Sub DrpQryHshC():                                     DrpQryHsh CDb:          End Sub

Sub DrpQryC(Qn):               DrpQry CDb, Qn:        End Sub
Sub DrpQry(D As Database, Qn): D.QueryDefs.Delete Qn: End Sub
Sub DrpQryHsh(D As Database)
Dim Q: For Each Q In Itr(QnyTmp(D))
    DrpQry D, Q
Next
End Sub
