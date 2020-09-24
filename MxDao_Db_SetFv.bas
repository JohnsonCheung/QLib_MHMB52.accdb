Attribute VB_Name = "MxDao_Db_SetFv"
Option Compare Text
Option Explicit
Const CMod$ = "MxDao_Db_SetFv."

Sub SetPvQ(D As Database, Q, V)
ValRs(D.OpenRecordset(Q)) = V
End Sub

Sub SetPvRs(A As Dao.Recordset, V)
If NoRec(A) Then
    A.AddNew
Else
    A.Edit
End If
A.Fields(0).Value = V
A.Update
End Sub

Sub SetPvRsF(Rs As Dao.Recordset, Fld, V)
With Rs
    .Edit
    .Fields(Fld).Value = V
    .Update
End With
End Sub

Sub SetPvSsk(D As Database, T, F$, SskvSet(), V)
ValRs(Rs(D, SqlSelFldWhFeq(F, T, Sskn(D, T), SskvSet))) = V
End Sub
