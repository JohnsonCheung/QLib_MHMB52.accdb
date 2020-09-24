Attribute VB_Name = "MxIde_MthMthlMthlsy"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_MthMthlMthlsy."

Private Sub B_Mthlsy(): VcLsy Mthlsy(SrcPC): End Sub
Sub VcMthlsyPC():       Vc FmtLsy(MthlsyPC): End Sub
Sub VcMthlsyMC():       Vc FmtLsy(MthlsyMC): End Sub

Function MthlsyMC() As String(): MthlsyMC = MthlsyM(CMd): End Function
Function MthlsyPC() As String(): MthlsyPC = MthlsyP(CPj): End Function
Function MthlsyP(P As VBProject) As String()
Dim C As VBComponent: For Each C In P.VBComponents
    PushIAy MthlsyP, MthlsyM(C.CodeModule)
Next
End Function
Function MthlsyM(M As CodeModule) As String(): MthlsyM = Mthlsy(SrcM(M)): End Function
Function Mthlsy(Src$()) As String()
Dim Ix: For Each Ix In Itr(Mthixy(Src))
    PushI Mthlsy, MthlIx(Src, Ix)
Next
End Function
Function MthlsyNmPC(Mthn, Optional ShtMthTy$) As String(): MthlsyNmPC = MthlsyMthnP(CPj, Mthn, ShtMthTy): End Function
Function MthlsyMthnP(P As VBProject, Mthn, Optional ShtMthTy$) As String()
Dim S$(): S = SrcPC
Dim Ix&(): Ix = MthixyMthn(S, Mthn, ShtMthTy)
Dim I: For Each I In Itr(Ix)
    PushI MthlsyMthnP, MthlIx(S, I)
Next
End Function
