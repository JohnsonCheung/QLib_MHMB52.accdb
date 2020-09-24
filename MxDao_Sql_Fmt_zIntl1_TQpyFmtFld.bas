Attribute VB_Name = "MxDao_Sql_Fmt_zIntl1_TQpyFmtFld"
Option Compare Database
Option Explicit

Function TQpyFmtFld(QSql() As TQp) As TQp()
Dim Iy%(): Iy = W_IyTQpy(QSql)
Dim QFld() As TQp ' each may be [SElect Distinct | Select]
Dim QFmt() As TQp ' Aft formatting TQpyFld
    QFld = TQpyWhIy(QSql, Iy)
    QFmt = W_QFldFmt(QFld)
TQpyFmtFld = TQpyRplIy(QSql, Iy, QFmt)
End Function
Private Function W_QFldFmt(FldLis() As TQp) As TQp()
Dim U&: U = UbTQp(FldLis)
Dim O() As TQp: ReDim O(U)
Dim J&: For J = 0 To U
    O(J) = W_TQpFmt(FldLis(J))
Next
W_QFldFmt = O
End Function
Private Function W_TQpFmt(FldLis As TQp) As TQp
Dim O$
    With FldLis
        Select Case .Qpt
        Case eQptSet: O = FmtFldLisSet(.Qpr)
        Case eQptSel: O = FmtFldLisSel(.Qpr)
        Case eQptSelDis: O = FmtFldLisSelDis(.Qpr)
        Case Else: ThwPm CSub, "@FldLis.Qpt must be [Set|Sel|SelDis]", "FldLis.Qpt .Qpr", EnmsQpt(.Qpt), .Qpr
        End Select
    End With
W_TQpFmt = FldLis
W_TQpFmt.Qpr = O
End Function
Private Function W_IyTQpy(Q() As TQp) As Integer()
'Only [eQpt? Set Sel SelDis] will have *FldLis
'@@Ret those Iy-to-@Q
Dim J%: For J = 0 To UbTQp(Q)
    Select Case Q(J).Qpt
    Case eQptSel, eQptSelDis, eQptSet: PushI W_IyTQpy, J
    End Select
Next
End Function

