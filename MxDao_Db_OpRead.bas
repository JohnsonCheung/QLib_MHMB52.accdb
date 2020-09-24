Attribute VB_Name = "MxDao_Db_OpRead"
Option Compare Text
Option Explicit
Const CMod$ = "MxDao_Db_OpRead."

Private Sub B_ValQ()
Dim D As Database
Ept = CByte(18)
Act = ValQ(D, "Select Y from [^YM]")
C
End Sub
Private Sub B_ValCQQ()
MsgBox ValQC(SqlSelFf("DD YY MM", "@OH", BeprFeq("YY", 20)))
MsgBox ValCQQ("Select DD from[@OH]where YY=?", 20)
End Sub
Function ValQ(D As Database, Q):  ValQ = ValRs(D.OpenRecordset(Q)): End Function
Function ValQC(Q):               ValQC = ValQ(CDb, Q):              End Function
Function ValCQQ(QQSql$, ParamArray Ap())
Dim Av(): If UBound(Ap) >= 0 Then Av = Ap
ValCQQ = ValQ(CDb, FmtQQAv(QQSql, Av))
End Function
Function ValQQ(D As Database, QQSql$, ParamArray Ap())
Dim Av(): Av = Ap
ValQQ = ValQ(D, FmtQQAv(QQSql, Av))
End Function
Function ValRs(A As Dao.Recordset, Optional F = 0)
If HasRec(A) Then ValRs = Nz(A.Fields(F).Value, Empty)
End Function
Function IdSskv&(D As Database, T, Sskv):            IdSskv = ValSskv(D, T, T & "Id", Sskv):                        End Function
Function ValSskv(D As Database, T, F, Sskv):        ValSskv = ValRs(Rs(D, SqlSelFldWhFeq(F, T, Sskn(D, T), Sskv))): End Function
Function ValSkvy(D As Database, T, F, Skvy()):      ValSkvy = ValQ(D, SqlSelFld(F, T, BeprSkvy(D, T, Skvy))):       End Function
Function ValF(D As Database, T, F, Optional Bepr$):    ValF = ValQ(D, SqlSelFld(F, T, Bepr)):                       End Function
Function ValTF(D As Database, TF$, Optional Bepr$):   ValTF = ValQ(D, SqlSelTFdot(TF, Bepr)):                       End Function
Function ValArs(A As ADODB.Recordset)
If NoRecArs(A) Then Exit Function
Dim V: V = A.Fields(0).Value
If IsNull(V) Then Exit Function
ValArs = V
End Function
Function ValCnq(A As ADODB.Connection, Q):                                   ValCnq = ValArs(A.Execute(Q)):      End Function
Function SqQ(D As Database, Q$, Optional IsInlFldn As Boolean) As Variant():    SqQ = SqRs(Rs(D, Q), IsInlFldn): End Function
