Attribute VB_Name = "MxIde_Src_Cac_Db_Mth"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_SrcFtcac_Db_Mth."

Function SkFnyWiSqlQPfx(D As Database, T) As String()
Dim F: For Each F In Itr(FnySk(D, T))
    PushI SkFnyWiSqlQPfx, ChrQuoSqlzDaoTy(DaotyF(D, T, F)) & F
Next
End Function

Sub IupTbl(D As Database, T, Drs As Drs) ' Ins or upd by @Drs has FnySk
Dim Dy(): Dy = Drs.Dy
If Si(Dy) = 0 Then Exit Sub
Dim R As Dao.Recordset, Q$, Sql$, Dr
'Sql = SqlSel_T_Wh(T, BeprzFnySqlQPfxSy(FnySk(D, T), SkSqlQPfxSy(D, T)))
For Each Dr In Dy
    Q = FmtQQAv(Sql, CvAv(Dr))
    Set R = Rs(D, Q)
    If HasRec(R) Then
        UpdRs R, Dr
    Else
        InsRs R, Dr
    End If
Next
End Sub
