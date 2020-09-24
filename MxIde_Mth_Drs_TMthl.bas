Attribute VB_Name = "MxIde_Mth_Drs_TMthl"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Mth_Drs_TMthl."

Public Const FfTMthl$ = "Mdn L Mthl" ' #Mth--Lines#

Function DrsTMthlPC() As Drs: DrsTMthlPC = DrsTMthlP(CPj): End Function
Function DrsTMthlP(P As VBProject) As Drs
Dim ODy()
    Dim C As VBComponent: For Each C In P.VBComponents
        PushIAy ODy, X_Dy(MdCmp(C))
    Next
DrsTMthlP = DrsFf(FfTMthl, ODy)
End Function
Function DrsTMthlM(M As CodeModule) As Drs: DrsTMthlM = DrsFf(FfTMthl, X_Dy(M)): End Function

Private Function X_Dy(M As CodeModule) As Variant()
Dim S$(): S = SrcM(M)
Dim N$: N = Mdn(M)
Dim Ix&: For Ix = 0 To UB(S)
    If IsLnMth(S(Ix)) Then
        PushI X_Dy, Array(N, Ix + 1, MthlIx(S, Ix))
    End If
Next
End Function
