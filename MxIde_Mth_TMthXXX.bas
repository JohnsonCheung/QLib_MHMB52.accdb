Attribute VB_Name = "MxIde_Mth_TMthXXX"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Mth_TMthXXX."

Function TMthmdnyPC() As TMthmdn(): TMthmdnyPC = TMthmdnyP(CPj): End Function
Function TMthmdnyP(P As VBProject) As TMthmdn()
Dim C As VBComponent: For Each C In P.VBComponents
    Dim M() As TMth: M = TMthyM(C.CodeModule)
    Dim N$: N = C.Name
    Dim J%: For J = 0 To UbTMth(M)
        PushTMthMd TMthmdnyP, TMthmdn(N, M(J))
    Next
Next
End Function
Function IsEqTMthMdn(A As TMthmdn, B As TMthmdn) As Boolean
With A
    Select Case True
    Case .Mdn <> B.Mdn, IsEqTMth(.TMth, B.TMth)
    Case Else: IsEqTMthMdn = True
    End Select
End With
End Function
