Attribute VB_Name = "MxIde_Src_Cac_Db_Tb_Mth_Op"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_SrcFtcac_Db_Tb_Mth_Op."

Sub RfhTbMthC(): RfhTbMth CDb: End Sub
Private Sub B_RfhTbMthId()
RfhTbMthId CDb, RfhIdLasC, CPj
End Sub

Sub RfhTbMthId(D As Database, RfhId&, Pj As VBProject)
'Do : delete reocrds to @D.Mth for those record @RfhId by
'     insert records to @D.Mth from $$Md->Mdl
Dim Ny$(), Ty$(): AsgTDcStr12 TDcStr12T(D, "Md", "Mdn CmpTy", "UpdId=" & RfhId), Ny, Ty
Dim Pjn$: Pjn = CPjn
Dim N: For Each N In Itr(Ny)
    D.Execute FmtQQ("Delete * from Mth where Mthn='?' and Pjn='?'", N, Pjn)
Next
'InsTblDrs D, "Mth", DrsTMthD(D, Pjn, ny)
End Sub

Sub RfhTbMth(D As Database)
Dim Pj As VBProject: Set Pj = CPj
Dim RfhId&: RfhId = NwRfhId(D)
'RfhTbMd  D, RfhId, Pj
RfhTbMthId D, RfhId, Pj

'Upd $$Lib from $$Md
RunqC "Delete * from Lib"
RunqC "Insert Into Lib Select Distinct Lib from Md"

'Upd $$Pj from $$Md
RunqC "Delete * from Pj"
RunqC "Insert Into Lib Select Distinct Pj from Md"
End Sub

Function RfhIdLasC&():             RfhIdLasC = RfhIdLas(CDb):         End Function
Function RfhIdLas&(D As Database):  RfhIdLas = IdRecLas(D, "RfhHis"): End Function

Function NwRfhId&(D As Database)
With D.TableDefs("RfhHis").OpenRecordset
    .AddNew
    NwRfhId = !RfhId
    .Update
    .Close
End With
End Function

Private Sub InsTbMth(D As Database, Pjn$, Mdny$(), ShtCmpTy$())
Dim N, J&: For Each N In Itr(Mdny)
    Dim Src$(): Src = SplitCrLf(MdlTbMd(D, Pjn, N))
    Dim Dr(): Dr = DrMdn(Pjn, ShtCmpTy(J), CStr(N), Si(Src))
    Dim Drs As Drs: Stop 'Drs = DrsTMthcS(Src, Dr)
    Stop 'InsTblDrs D, "Mth", DrsTMth(D, Pjn, Mdny)
    J = J + 1
Next
End Sub
