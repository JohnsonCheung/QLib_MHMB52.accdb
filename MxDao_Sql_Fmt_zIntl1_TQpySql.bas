Attribute VB_Name = "MxDao_Sql_Fmt_zIntl1_TQpySql"
Option Compare Text
Option Explicit
Const CMod$ = "MxDao_Sql_FmtSql_SqlAdjKww."
Private Sub B_QpySql()
GoSub ZZ
Exit Sub
ZZ:
    Dim Lyy()
    Dim Qn: For Each Qn In QnyC
        Dim Sql$: Sql = SqlQnC(Qn)
        PushI Lyy, QpySql(Sql)
    Next
    BrwLyy Lyy
    Return
End Sub

Private Function QpySql(Sql) As String()
Dim O$: O = RplCrLf(Sql)
Dim Kwyy(): Kwyy = Qpkwyy
Dim Kwy: For Each Kwy In Kwyy
    Dim P12y() As P12: P12y = P12yKwy(O, Kwy)
    Dim I%: For I = UbP12(P12y) To 0 Step -1
        Dim P As P12: P = P12y(I)
        O = RplP12(O, P, JnSpc(Kwy), Inl:=True)
        O = StrInsAtCrLf(O, P.P1)
    Next
Next
QpySql = SplitCrLf(TrimWhiteL(O))
'D QpySql: Stop
End Function

Function TQpySql(Sql) As TQp()
Dim Qpy$(): Qpy = QpySql(Sql): 'ShwTimr "QpySql...!TQpySql"
Dim Qp: For Each Qp In Qpy
    'Debug.Print Qp
    PushTQp TQpySql, TQpQp(Qp)
Next
TQpySql = TQpyTrimQpry(TQpySql): 'ShwTimr "TQpySql...!TQpySql"
End Function
Private Function TQpyTrimQpry(Q() As TQp) As TQp()
Dim O() As TQp: O = Q
Dim J%: For J = 0 To UbTQp(Q)
    O(J).Qpr = Trim(O(J).Qpr)
Next
TQpyTrimQpry = O
End Function

