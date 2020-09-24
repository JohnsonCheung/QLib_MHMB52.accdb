Attribute VB_Name = "MxDao_Sql_Fmt_zIntl1_TnySql"
Option Compare Text
Option Explicit
Function TnySql(Sql) As String()
Dim QSql() As TQp: QSql = TQpySql(Sql)
Dim Iy%(): Iy = IyTbnTQpy(QSql)
Dim QTbn() As TQp: QTbn = TQpyWhIy(QSql, Iy)
Dim O$()
    Dim J%: For J = 0 To UbTQp(QTbn)
        With QTbn(J)
            Select Case .Qpt
            Case eQptUpd: PushI O, TbnQpr_Fm_or_Upd(.Qpr)
            Case eQptFm: PushI O, TbnQpr_Fm_or_Upd(.Qpr)
            Case eQptInto: PushI O, RmvBktSq(.Qpr)
            Case eQptInrJn, eQptLeftJn: PushI O, TbnQpr_Jn(.Qpr)
            Case Else: Thw CSub, "Given @Sql does not started with [Update Insert Select Delete]", "Sql", Sql
            End Select
        End With
    Next
TnySql = O
End Function
Private Function IyTbnTQpy(QSql() As TQp) As Integer()
' Iy of those @QSql which contains Tbn.  It is used in TnySql
Static QptyTbn() As eQpt ' Qpt which contains 'Tbn'
    If Si(QptyTbn) = 0 Then
        QptyTbn = Lngy(eQptFm, eQptUpd, eQptInrJn, eQptLeftJn, eQptInto) ' Contains Tbn
    End If
IyTbnTQpy = IyEley(eQptyTQpy(QSql), QptyTbn)
End Function

Private Function TbnQpr_Fm_or_Upd$(Qpr$)
Dim O$, O1$, O2$
O1 = RmvChrPfxAll(Qpr, "(")
O2 = Trim(BefOrAll(O1, " As "))
O = RmvBktSq(O2)
TbnQpr_Fm_or_Upd = O
End Function
Private Function TbnQpr_Jn$(QprJn$)
'Left Join [XX] As A On ....
'QprJn=    ....
Dim PosAs%
    PosAs = PosSsub(QprJn, " As ")
    If PosAs > 0 Then
        Dim A$: A = Left(QprJn, PosAs - 1)
        TbnQpr_Jn = RmvBktSq(A)
        Exit Function
    End If
'Left Join XX On ...
Dim PosOn%
    PosOn = PosSsub(QprJn, " On ")
    If PosOn > 0 Then
        Dim B$: B = Left(QprJn, PosOn - 1)
        TbnQpr_Jn = RmvBktSq(B)
        Exit Function
    End If
ThwPm CSub, "There is no ' On ' nor ' As ' in @QprJn", "@QprJn", QprJn
End Function
