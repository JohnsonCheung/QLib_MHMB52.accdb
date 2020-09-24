Attribute VB_Name = "MxDao_Db_ToRs"
Option Compare Text
Option Explicit
Const CMod$ = "MxDao_Db_ToRs."

Function RsTFeqC(TF$, TFeqv):                                  Set RsTFeqC = RsTFeq(CDb, TF, TFeqv):           End Function
Function RsTFeq(D As Database, TF$, TFeqv):                     Set RsTFeq = Rs(D, SqlSelTFeq(TF, TFeqv)):     End Function
Function RsQnC(Qn) As Dao.Recordset:                             Set RsQnC = RsQn(CDb, Qn):                    End Function
Function RsQn(D As Database, Qn) As Dao.Recordset:                Set RsQn = Qd(D, Qn).OpenRecordset:          End Function
Function RsQC(Q) As Dao.Recordset:                                Set RsQC = Rs(CDb, Q):                       End Function
Function RsQ(D As Database, Q) As Dao.Recordset:                   Set RsQ = Rs(D, Q):                         End Function
Function RsTbFeq(D As Database, T, F$, Feqv) As Dao.Recordset: Set RsTbFeq = Rs(D, SqlSelStarFeq(T, F, Feqv)): End Function
Function Rs(D As Database, Q) As Dao.Recordset
Const CSub$ = CMod & "Rs"
On Error GoTo X
Set Rs = D.OpenRecordset(Q)
Exit Function
X: ThwSql CSub, D, Q, Err.Description
End Function
Function RsQQC(QQ$, ParamArray Ap()) As Dao.Recordset
Dim Av():  Av = Ap
Set RsQQC = Rs(CDb, FmtQQAv(QQ, Av))
End Function
