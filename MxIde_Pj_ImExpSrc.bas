Attribute VB_Name = "MxIde_Pj_ImExpSrc"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Pj_Exp."

Sub ExpSrcPC(): ExpSrcPthP CPj, PthSrcPC: End Sub
Sub ExpSrcPthP(P As VBProject, PthTo$)
Dim J%, N%: N = P.VBComponents.Count
Dim C As VBComponent: For Each C In P.VBComponents
    DoEvents
    If J Mod 100 = 0 Then Debug.Print "Exporting:"; J + 1; "of"; N; C.Name; " ...."
    J = J + 1
    ExpSrcM C.CodeModule, PthTo, NoInf:=True
Next
End Sub
Sub ExpSrcM(M As CodeModule, PthTo$, Optional NoInf As Boolean)
Dim T As vbext_ComponentType: T = M.Parent.Type
Select Case True
Case T = vbext_ct_ClassModule, T = vbext_ct_StdModule
    If Not NoInf Then Debug.Print Now, "Exp Md", Mdn(M)
    M.Parent.Export FfnSrcM(M, PthTo)
Case Else: Exit Sub
End Select
End Sub

Sub ImpSrcPC(PthFm$): ImpSrcP CPj, PthFm: End Sub
Sub ImpSrcP(P As VBProject, PthFm$)
Dim I: For Each I In Itr(WFfnyBas(PthFm))
    Debug.Print Now, "Imp Md", Fn(I)
    P.VBComponents.Import I
Next
End Sub

Private Function WFfnyBas(Pth$) As String(): WFfnyBas = Ffny(Pth, "*.bas"): End Function
