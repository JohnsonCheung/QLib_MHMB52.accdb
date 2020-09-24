Attribute VB_Name = "MxIde_Pj_Mxn"
Option Compare Text
Const CMod$ = "MxIde_Pj_Mxn."
Option Explicit

Private Sub B_MxnyPC():                 Dmp MxnyPC: End Sub
Function MxnyPC() As String(): MxnyPC = MxnyP(CPj): End Function
Function MxnyP(P As VBProject) As String()
Dim M: For Each M In MdnyP(P)
    Dim A$: A = Bef(M, "_")
    If HasPfx(A, "Mx") Then
        PushNoDup MxnyP, A
    End If
Next
End Function
