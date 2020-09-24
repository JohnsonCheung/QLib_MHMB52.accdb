Attribute VB_Name = "MxIde_Mth_Dup_MMthlNm"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Mth_Dup_MMthlNm."
Function MMthlsyNmP(P As VBProject, Mthn, Optional ShtMthTy$) As String()
Dim C As VBComponent: For Each C In P.VBComponents
    PushIAy MMthlsyNmP, MMthlsyNmM(C.CodeModule, Mthn, ShtMthTy)
Next
End Function

Function MMthlsyNm(Mdn$, Src$(), Mthn, Optional ShtMthTy$) As String():
Dim Ix: For Each Ix In Itr(MthixyMthn(Src, Mthn, ShtMthTy))
 Stop '   PushI LyUL(Mdn) & vbCrLf & S12yMdnMthl, Mthl(Src, Ix)
Next
End Function
Function MMthlsyNmM(M As CodeModule, Mthn, Optional ShtMthTy$) As String(): MMthlsyNmM = MMthlsyNm(Mdn(M), SrcM(M), Mthn, ShtMthTy): End Function

Function MMthlsyNmPC(Mthn, Optional ShtMthTy$) As String(): MMthlsyNmPC = MMthlsyNmP(CPj, Mthn, ShtMthTy): End Function
