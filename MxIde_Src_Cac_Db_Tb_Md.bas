Attribute VB_Name = "MxIde_Src_Cac_Db_Tb_Md"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_SrcFtcac_Db_Tb_Md."
Sub RfhTbMdP(): RfhTbMd CPj: End Sub
Sub RfhTbMd(P As VBProject)
Dim Id&: Id = Pjid(P.Name)
Dim C As VBComponent: For Each C In P.VBComponents
    WRfh Id, C
Next
End Sub
Private Sub WRfh(Pjid&, C As VBComponent)
Dim M As CodeModule: Set M = C.CodeModule
Dim Mdn$: Mdn = C.Name
Dim CmpTy&: CmpTy = C.Type
Dim Rs As Dao.Recordset: Set Rs = RsSkvap(CDb, "Md", Pjid, Mdn)
If HasRec(Rs) Then
    WUpdIfNeed Rs, CmpTy
Else
    WSetSrcIns Pjid, Mdn, CmpTy
End If
End Sub
Private Sub WSetSrcIns(Pjid&, Mdn$, CmpTy&)
DoCmd.RunSQL SqlInsFfDr("Md", "PjId Mdn CmpTy", Array(Pjid, Mdn, CmpTy))
End Sub
Private Sub WUpdIfNeed(Rs As Dao.Recordset, CmpTy&)
With Rs
Select Case True
Case !CmpTy <> CmpTy
    .Edit
    !CmpTy = CmpTy
    .Update
End Select
End With
End Sub
