Attribute VB_Name = "MxIde_Mth_Op_CpyMth"
Option Compare Text
Const CMod$ = "MxIde_Mth_Op_Cpy."
Option Explicit

Sub CpyMthTo(M As CodeModule, Mthn, MthnTo$)
Const CSub$ = CMod & "CpyMthTo"
Dim S$(): S = SrcM(M)
If Mthix(S, MthnTo) <> 0 Then Inf CSub, "AsMth exist.", "Mdn Mthn AsMth", Mdn(M), Mthn, MthnTo: Exit Sub
Dim IxFm&: IxFm = Mthix(S, Mthn)
If IxFm <> -1 Then Inf CSub, "MthnFm does not exist.", "Mdn MthnFm MthTo", Mdn(M), Mthn, MthnTo: Exit Sub
M.InsertLines IxFm + 1, WNewl$(S, IxFm&, Mthn, MthnTo)
End Sub
Private Function WNewl$(Src$(), IxFm&, MthnFm, MthnTo$)
Stop 'MthlRen Src(0), MthnFm, MthnTo
End Function

Sub CpyMth(Mthn, M As CodeModule, MdTo As CodeModule)
Const CSub$ = CMod & "CpyMth"
Dim S$(): S = SrcM(M)
Dim IxFm&: IxFm = Mthix(S, Mthn)
If IxFm = -1 Then Thw CSub, "MdTo has mthn FmM", "Mthn MdFm MdTo", Mthn, Mdn(M), Mdn(MdTo)
If Mthix(SrcM(MdTo), Mthn) > 0 Then Thw CSub, "MdTo already has mthn", "Mthn MdFm MdTo", Mthn, Mdn(M), Mdn(MdTo)
MdTo.AddFromString MthlIx(S, IxFm)
End Sub

Sub CpyMthAsVer(M As CodeModule, Mthn, Ver%)
Const CSub$ = CMod & "CpyMthAsVer"
'Ret True if copied
Dim VerMthn$, Newl$, L$, Oldl$
If Not HasMthnM(M, Mthn) Then Inf CSub, "No from-mthn", "Md Mthn", Mdn(M), Mthn: Exit Sub
VerMthn = Mthn & "_Ver" & Ver
'NewL
    L = MthlNmM(M, Mthn)
    Newl = Replace(L, "Sub " & Mthn, "Sub " & VerMthn, Count:=1)
'Rpl
    RplMth M, VerMthn, Newl
End Sub
