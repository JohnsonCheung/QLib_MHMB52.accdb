Attribute VB_Name = "MxIde_Pj_Pjf"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Pj_Pjf."
Function PjfyVC() As String(): PjfyVC = PjfyV(CVbe): End Function
Function PjfyV(V As VBE) As String()
Dim P As VBProject: For Each P In V.VBProject
    PushI PjfyV, Pjf(P)
Next
End Function
Function PjFm$(M As CodeModule): PjFm = Pjf(PjM(M)): End Function
