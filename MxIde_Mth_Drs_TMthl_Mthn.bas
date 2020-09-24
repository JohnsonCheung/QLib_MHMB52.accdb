Attribute VB_Name = "MxIde_Mth_Drs_TMthl_Mthn"
Option Compare Text
Const CMod$ = "MxIde_Mth_Drs_TMthl_Mthn."
Option Explicit

Private Sub B_DrsTMthlMthnPC():                                        BrwDrs DrsTMthlMthnPC("X_Dy"): End Sub
Function DrsTMthlMthnPC(Mthn) As Drs:                 DrsTMthlMthnPC = DrsTMthlMthnP(CPj, Mthn):      End Function
Function DrsTMthlMthnM(M As CodeModule, Mthn) As Drs:  DrsTMthlMthnM = DrsFf(FfTMthl, X_Dy(M, Mthn)): End Function
Function DrsTMthlMthnP(P As VBProject, Mthn) As Drs
Dim ODy()
    Dim C As VBComponent: For Each C In P.VBComponents
        PushIAy ODy, X_Dy(MdCmp(C), Mthn)
    Next
DrsTMthlMthnP = DrsFf(FfTMthl, ODy)
End Function

Private Function X_Dy(M As CodeModule, Mthn) As Variant()
Dim S$(): S = SrcM(M)
Dim N$: N = Mdn(M)
Dim Ix&: For Ix = 0 To UB(S)
    If IsLnMthn(S(Ix), Mthn) Then
        PushI X_Dy, Array(N, Ix + 1, MthlIx(S, Ix))
    End If
Next
End Function
