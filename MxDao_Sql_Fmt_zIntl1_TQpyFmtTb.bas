Attribute VB_Name = "MxDao_Sql_Fmt_zIntl1_TQpyFmtTb"
Option Compare Database
Option Explicit

Function TQpyFmtTb(QSql() As TQp) As TQp()
Dim QTb() As TQp  ' Those TQp from QSql which contain Tbn
Dim B As Bei    ' Pointing to QSql for each contains Tb-joining-together, ie: Upd | Fm | Left Join | Inner Join
    B = W_BeiJnTQpy(QSql)
    QTb = TQpyWhBei(QSql, B) ' Stop
Dim QTbFmt() As TQp
    QTbFmt = W_QTbFmt(QTb)
TQpyFmtTb = TQpyRplBix(QSql, B.Bix, QTbFmt) ' Put QTbFmt back to QSql and return
End Function
Private Function W_QTbFmt(QTb() As TQp) As TQp()
'@@Ret Aft formatting QTb
Dim QpyTbFmt$()     ' Each Qp are (Tb: Fm + Jn) and they are formatted
QpyTbFmt = W_QpyFmtQTb(QTb)
If Not W_IsQpyRndTripOK(QpyTbFmt) Then Stop
W_QTbFmt = TQpyQpy(QpyTbFmt)
'Dmp QpyTbFmt: Stop
End Function
Private Function W_IsQpyRndTripOK(Qpy$()) As Boolean
Dim Q() As TQp: Q = TQpyQpy(Qpy)
Dim QpyR$(): QpyR = QpyTQpy(Q)
W_IsQpyRndTripOK = IsEqAy(QpyR, Qpy)
If Not W_IsQpyRndTripOK Then
    Dmp LyUL("Orig"): Dmp Qpy
    Dmp LyUL("RndTrip"): Dmp QpyR
    Stop
End If
End Function

Private Function W_QpyFmtQTb(QTb() As TQp) As String()
'@QTb are those only contains Tbn: Which is [Upd Fm InrJn LeftJn]
Dim LndyFm() ' Each LndyFm-0   2 ele-columns: ETb EAs
             '      LdnyFm-rst 4 ele-columns: ETb EAs EOn EEq
LndyFm = W_LndyFm(QTb)
W_QpyFmtQTb = FmtLndy(LndyFm, NoIx:=True, Sep:="")
'Dmp W_QpyFmtQTb: Stop
End Function
Private Function W_LndyFm(QTb() As TQp) As Variant()
If SiTQp(QTb) = 0 Then Exit Function
Dim DrFm$()
    DrFm = W_Dr3QpFm(QTb(0)) ' Assum QTb(0) is either 'Fm' or 'Upd'-QTp
Dim OLndy()
    PushI OLndy, DrFm
    Dim J%: For J = 1 To UbTQp(QTb)
        PushI OLndy, W_Dr5Jn(QTb(J)) ' Assume rest of QTb, (1-UB) is either 'InrJn' or 'LeftJn'
    Next
W_LndyFm = OLndy
End Function

Private Function W_Dr3QpFm(QFm As TQp) As String()
If QFm.Qpt <> eQptFm And QFm.Qpt <> eQptUpd Then ThwPm CSub, "@TQpFm.Qpt must be 'From' or 'Update'", "@TQpFm.Qpt", EnmsQpt(QFm.Qpt)
'Update|From ((( xx As x     'assume all ( has not space between
'kkkkkkkkkkkkkkk ttaaaaa
Dim L$: L = QFm.Qpr
Dim Qp$: Qp = StrTQp(QFm)
Dim EKw$, ETb$, EAs$
    EKw = EnmtQpt(QFm.Qpt) & " " & ShfChrPfxAll(L, "(")
    ETb = W_ShfETb(L, Qp)
    EAs = W_ShfEAs(L, Qp)
W_Dr3QpFm = Sy(EKw, ETb, EAs)
End Function
Private Function W_Dr5Jn(Jn As TQp) As String()
'Left|Inner Join XXX As x On xxx = xxx)))  ' It is either 'Left Join' or 'Inner Join'
'kkkkkkkkkkkkkkk tttaaaaaooooooorrrrrrrrr  ' aaaa is optional
Dim L$: L = Jn.Qpr
Dim Qp$: Qp = StrTQp(Jn)
Dim EKw$, ETb$, EAs$, EOn$, EEq$
    EKw = EnmtQpt(Jn.Qpt) & " "
    ETb = W_ShfETb(L, Qp)
    EAs = W_ShfEAs(L, Qp)
    EOn = W_ShfEOn(L, Qp)
    EEq = W_ShfEEq(L, Qp)
W_Dr5Jn = Sy(EKw, ETb, EAs, EOn, EEq)
End Function
Private Function W_ShfETb$(OLn$, Qp$)
OLn = LTrim(OLn)
W_ShfETb = ShfSqlTm(OLn): If W_ShfETb = "" Then Thw CSub, "Given @OLn should be able to shf a Tbn", "@Qp", Qp
OLn = LTrim(OLn)
End Function
Private Function W_ShfEAs$(OLn$, Qp$)
Dim L$: L = OLn
If Not IsShfAs(OLn) Then Exit Function
OLn = LTrim(OLn)
Dim N$: N = ShfNm(OLn): If N = "" Then Thw CSub, "There is 'As ', but no name after", "@Qp", Qp
OLn = LTrim(OLn)
W_ShfEAs = " As " & N
End Function
Private Function W_ShfEOn$(OLn$, Qp$)
If Not IsShfOn(OLn) Then Thw CSub, "Kw 'On ' is missing", "@OLn", OLn
Dim B As S12: B = Brk1(OLn, "=")
If B.S2 = "" Then Thw CSub, "After 'On' there is no '='", "@Qp", Qp
OLn = "=" & LTrim(B.S2)
W_ShfEOn = " On " & B.S1
End Function
Private Function W_ShfEEq$(OLn$, Qp$)
If ChrFst(OLn) <> "=" Then Thw CSub, "ChrFst-@OLn should be '='", "@OLn", OLn
W_ShfEEq = " = " & Trim(Mid(OLn, 2))
OLn = ""
End Function
Private Function W_BeiJnTQpy(QSql() As TQp) As Bei
' Bei of those @QSql which contains Tbn joining togeterh.  Note: from those TQp-contains-Tbn-joining togeter are in a series, so Bei is good to use
Static QptyTbJn() As eQpt ' Qpt which contains 'Tbn-Joining-Together'
    If Si(QptyTbJn) = 0 Then
        QptyTbJn = Lngy(eQptFm, eQptUpd, eQptInrJn, eQptLeftJn) ' Tbn-joining-together
    End If
W_BeiJnTQpy = BeiSubay(eQptyTQpy(QSql), QptyTbJn)
End Function


Private Sub B_W_Dr5Jn()
GoSub T1
Exit Sub
Dim TQpJn As TQp, Ept$(), Act$()
T1:
    TQpJn = TQpQp("Inner Join qPHL1 ON qPHL3.PHL1 = qPHL1.PHL1")
    Ept = Sy( _
        "Inner Join ", _
        "qPHL1", _
        "", _
        " On qPHL3.PHL1", _
        " = qPHL1.PHL1")
    GoTo Tst
T2:
    TQpJn = TQpQp("Left Join X    As x On     xxx     =     aaa")
    Ept = SyAp( _
        "Left Join ", _
        "X", _
        "As x", _
        " On xxx", _
        " = aaa")
    Stop
    GoTo Tst
Tst:
    Act = W_Dr5Jn(TQpJn)
    ChkEq Ept, Act
    Return
End Sub


