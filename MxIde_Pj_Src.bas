Attribute VB_Name = "MxIde_Pj_Src"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Pj_Src."
Private Sub B_SrcPC():                        VcAy TIndSrcPC: End Sub
Function TIndSrcPC() As String(): TIndSrcPC = TIndSrcP(CPj):  End Function
Function TIndSrcP(P As VBProject) As String()
Dim C As VBComponent: For Each C In P.VBComponents
    PushIAy TIndSrcP, TIndSrcM(C.CodeModule)
Next
End Function
Function TIndSrcM(M As CodeModule) As String()
PushI TIndSrcM, Mdn(M)
PushIAy TIndSrcM, AmAddPfx(SrcM(M), "  ")
End Function
