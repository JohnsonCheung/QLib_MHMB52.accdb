Attribute VB_Name = "MxIde_Mth_S12yMdnMthl"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Mth_S12yMdnMthl."

Function S12yMdnMthlMthnPC(Mthn, Optional ShtMthTy$) As S12(): S12yMdnMthlMthnPC = S12yMdnMthlMthnP(CPj, Mthn, ShtMthTy): End Function
Function S12yMdnMthlMthnP(P As VBProject, Mthn, Optional ShtMthTy$) As S12()
Dim C As VBComponent: For Each C In P.VBProject
    Dim S$(): S = SrcCmp(C)
    Dim Ix&(): Ix = MthixyMthn(S, Mthn, ShtMthTy)
    Dim I: For Each I In Itr(Ix)
        PushS12 S12yMdnMthlMthnP, S12(C.Name, MthlIx(S, I))
    Next
Next
End Function

Function S12yMdnMthlMthnyP(P As VBProject, Mthny$()) As S12()
Dim C As VBComponent: For Each C In P.VBComponents
    Dim S$(): S = SrcCmp(C)
    Dim Ix&(): Ix = MthixyMthny(S, Mthny)
    Dim I: For Each I In Itr(Ix)
        PushS12 S12yMdnMthlMthnyP, S12(C.Name, MthlIx(S, I))
    Next
Next
End Function
Private Function WMthl$(N As TMth, Mdn$)
Dim S$(): S = SrcMdn(Mdn)
Stop 'WMthl = Mthl(S, Mthix(S, N.Nm, N.ShtTy))
End Function
