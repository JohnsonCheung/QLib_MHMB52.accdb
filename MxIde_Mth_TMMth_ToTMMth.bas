Attribute VB_Name = "MxIde_Mth_TMMth_ToTMMth"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Mth_TMMth_ToTMMth."
Private Sub B_TMMthyPC()
Dim A() As TMMth: A = TMMthyPC
Stop
End Sub
Function TMMthyPC() As TMMth(): TMMthyPC = TMMthyP(CPj): End Function
Function TMMthyP(P As VBProject) As TMMth()
Dim C As VBComponent: For Each C In P.VBComponents
    PushTMMthy TMMthyP, TMMthyM(C.CodeModule)
Next
End Function
Function TMMthyMC() As TMMth(): TMMthyMC = TMMthyM(CMd): End Function
Function TMMthyM(M As CodeModule) As TMMth()
Dim S$(): S = SrcM(M)
Dim Ix: For Each Ix In Itr(Mthixy(S))
    PushTMMthy TMMthyM, WTMMthySrc(S, Mdn(M))
Next
End Function
Private Function WTMMthySrc(Src$(), Mdn$) As TMMth()
Stop
End Function
